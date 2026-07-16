#!/usr/bin/env python3
"""美兰运营分析看板 — 数据提取脚本（内联数据模式）"""
import pandas as pd
import json, re, os
from datetime import datetime, timedelta
from collections import defaultdict

SRC = '/home/Vic/.hermes/tmp_data/美兰数据源收集表.xlsx'
OUT_FILE = '/home/Vic/dewu-reports/MeilanAnalysis/2026/index.html'
BRAND = '美兰'

def cn(v):
    if pd.isna(v): return 0.0
    s = str(v).strip()
    # Handle double-dots like '263..47' → '263.47'
    s = s.replace('..', '.')
    m = re.match(r'([\d.]+)', s)
    return float(m.group(1)) if m else 0.0

def nd(v):
    if pd.isna(v): return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    try:
        n = float(s)
        if 40000 < n < 60000:
            return (datetime(1899,12,30) + timedelta(days=int(n))).strftime('%Y-%m-%d')
    except: pass
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S']:
        try: return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except: pass
    return s[:10] if len(s)>=10 else None

def nm(v):
    d = nd(v)
    return d[:7] if d else None

def ss(v):
    if pd.isna(v): return ''
    return str(v).strip()

print(f"📂 {SRC}")

# ============================================================
# 1. 商详访客
# ============================================================
print("📊 [1/6] 商详访客...")
uv = pd.read_excel(SRC, sheet_name='商详访客数据', engine='calamine')
uv = uv[['日期', '货号', '支付订单金额', '支付订单量', '商详访问指数（UV）']].copy()
uv.columns = ['date_raw', 'goods', 'gmv', 'orders', 'uv']
uv['date_str'] = uv['date_raw'].apply(nd)
uv = uv.dropna(subset=['date_str'])
uv['gmv'] = uv['gmv'].apply(cn)
uv['orders'] = uv['orders'].apply(cn)
uv['uv'] = uv['uv'].apply(cn)
uv = uv[uv['goods'].apply(lambda x: ss(x) not in ['','nan'])]

db = uv.groupby('date_str').agg(GMV=('gmv','sum'), UV=('uv','sum'), orders=('orders','sum')).reset_index()
db['brand'] = BRAND
DAILY_BRAND = db[['date_str','brand','GMV','UV','orders']].to_dict('records')

dg = uv.groupby(['date_str','goods']).agg(GMV=('gmv','sum'), UV=('uv','sum'), orders=('orders','sum')).reset_index()
dg['brand'] = BRAND
DAILY_GOODS = dg.rename(columns={'goods':'商品货号'})[['date_str','brand','商品货号','GMV','UV','orders']].to_dict('records')

ALL_DATES = sorted(set(r['date_str'] for r in DAILY_BRAND))
ALL_GOODS = sorted(set(r['商品货号'] for r in DAILY_GOODS))
ALL_MONTHS = sorted(set(d[:7] for d in ALL_DATES))
print(f"   {len(DAILY_BRAND)} days, {len(DAILY_GOODS)} day-goods, {ALL_DATES[0]}~{ALL_DATES[-1]}")

# ============================================================
# 2. 交易订单 + 售后
# ============================================================
print("📦 [2/6] 交易+售后...")
ord_df = pd.read_excel(SRC, sheet_name='交易订单', engine='calamine')
ord_df = ord_df[['货号', '数量', '出价金额（元）', '订单状态', '下单日期']].copy()
ord_df.columns = ['goods', 'qty', 'amount', 'status', 'date_raw']
ord_df['date_str'] = ord_df['date_raw'].apply(nd)
ord_df = ord_df.dropna(subset=['date_str','goods'])
ord_df['amount'] = ord_df['amount'].apply(cn)
ord_df['qty'] = ord_df['qty'].apply(lambda x: int(cn(x)) if cn(x)>0 else 1)

ret_df = pd.read_excel(SRC, sheet_name='售后订单', engine='calamine')
ret_df = ret_df[['商品货号', '售后原因', '买家申请售后时间']].copy()
ret_df.columns = ['goods', 'reason', 'apply_time']
ret_df['date_str'] = ret_df['apply_time'].apply(nd)
ret_df = ret_df.dropna(subset=['date_str','goods'])
ret_returns = ret_df[ret_df['reason'].astype(str).str.contains('退货|退款|换货', na=False)]

gr = defaultdict(lambda:{'total':0,'returned':0})
for _,r in ord_df.iterrows():
    g=ss(r['goods'])
    if g: gr[g]['total']+=r['qty']
for _,r in ret_returns.iterrows():
    g=ss(r['goods'])
    if g: gr[g]['returned']+=1

GOODS_RATE=[]
for g,v in gr.items():
    if v['total']>=4:
        rate=round(v['returned']/v['total']*100,2)
        GOODS_RATE.append({'货号':g,'total':v['total'],'returned':v['returned'],'return_rate':rate})
GOODS_RATE.sort(key=lambda x:x['return_rate'],reverse=True)

today = ALL_DATES[-1] if ALL_DATES else '2026-06-01'
def anomaly(window):
    cutoff = (datetime.strptime(today,'%Y-%m-%d')-timedelta(days=window)).strftime('%Y-%m-%d')
    ro = ord_df[ord_df['date_str']>=cutoff]; ho = ord_df[ord_df['date_str']<cutoff]
    rr = ret_returns[ret_returns['date_str']>=cutoff]; hr = ret_returns[ret_returns['date_str']<cutoff]
    goods_set = set()
    for df in [ro,ho]:
        for g in df['goods']: 
            gs=ss(g); 
            if gs: goods_set.add(gs)
    results=[]
    for g in goods_set:
        rt = len(ro[ro['goods'].astype(str).str.strip()==g])
        ht = len(ho[ho['goods'].astype(str).str.strip()==g])
        if rt<3: continue
        rrc = len(rr[rr['goods'].astype(str).str.strip()==g])
        hrc = len(hr[hr['goods'].astype(str).str.strip()==g])
        rr2 = rrc/rt*100 if rt else 0; hr2 = hrc/ht*100 if ht else 0
        d = rr2-hr2
        results.append({'goods':g,'hist_rate':round(hr2,1),'recent_rate':round(rr2,1),'delta':round(d,1),
                       'total_orders':ht+rt,'recent_orders':rt,'is_alert':d>=3})
    results.sort(key=lambda x:x['recent_rate'],reverse=True)
    return results

ANOMALY_7D = anomaly(7); ANOMALY_30D = anomaly(30)
print(f"   GOODS_RATE:{len(GOODS_RATE)} ANOMALY_7D:{len(ANOMALY_7D)} ANOMALY_30D:{len(ANOMALY_30D)}")

# ============================================================
# 3. 大盘
# ============================================================
print("📈 [3/6] 大盘...")
mkt = pd.read_excel(SRC, sheet_name='大盘日韩表指数', engine='calamine')
mkt = mkt[['日期', '日韩表大盘成交GMV指数']].copy()
mkt.columns = ['date_raw','idx']
mkt['date_str'] = mkt['date_raw'].apply(nd)
mkt = mkt.dropna(subset=['date_str'])
mkt['idx'] = mkt['idx'].apply(cn)
MARKET_MAP = {}
for _,r in mkt.iterrows():
    MARKET_MAP[r['date_str']] = round(r['idx']*10000,2)
print(f"   {len(MARKET_MAP)} dates")

# ============================================================
# 4. 社区投放（仅4种有效状态）
# ============================================================
print("📢 [4/6] 社区投放...")
comm = pd.read_excel(SRC, sheet_name='社区投放任务', engine='calamine')
comm = comm[['任务月份', '实际任务金额', '动态发布时间', '匹配货号', '曝光', '阅读数', '互动数', '商详访问', '任务状态']].copy()
comm.columns = ['task_m_raw','amount','pub_raw','goods_raw','exposure','reads','interact','visits','status']
comm = comm[comm['status'].astype(str).str.strip().isin(['待验收','待寄回','待收货','完成'])]
comm['task_month'] = comm['task_m_raw'].apply(nm)
comm['pub_date'] = comm['pub_raw'].apply(nd)
comm['goods'] = comm['goods_raw'].apply(ss)
comm = comm.dropna(subset=['goods']); comm = comm[comm['goods']!='']
for c in ['amount','exposure','reads','interact','visits']:
    comm[c] = comm[c].apply(cn)

TASK_RAW=[]; tms=set(); pms=set()
for _,r in comm.iterrows():
    gl=[g.strip() for g in r['goods'].split(',') if g.strip()]
    pa=r['amount']/len(gl) if gl else r['amount']
    for g in gl:
        tm=r['task_month'] or ''; pd2=r['pub_date']; pm=pd2[:7] if isinstance(pd2,str) and len(pd2)>=7 else ''
        TASK_RAW.append({'匹配货号':g,'task_month':tm,'pub_month':pm,'动态发布数量':1,
            '实际任务金额':round(pa,2),'曝光':int(r['exposure']),'阅读数':int(r['reads']),
            '商详访问':int(r['visits']),'互动数':int(r['interact'])})
        if tm: tms.add(tm)
        if pm: pms.add(pm)
TASK_MONTHS=sorted(tms); PUB_MONTHS=sorted(pms)
print(f"   {len(TASK_RAW)} items")

# ============================================================
# 5. 得物推
# ============================================================
print("🚀 [5/6] 得物推...")
push = pd.read_excel(SRC, sheet_name='得物推数据-商品', engine='calamine')
push = push[['时间','商品ID','消耗(元)','直接支付单量(单)','直接支付金额(元)','引导支付单量(单)','引导支付金额(元)','货号']].copy()
push.columns = ['date_raw','gid','cost','do2','dgmv','io2','igmv','goods']
push['goods'] = push['goods'].apply(ss)
push['date_str'] = push['date_raw'].apply(nd); push = push.dropna(subset=['date_str'])
for c in ['cost','dgmv','igmv']: push[c]=push[c].apply(cn)
push['do2']=push['do2'].apply(lambda x:int(cn(x)))
push['io2']=push['io2'].apply(lambda x:int(cn(x)))

PUSH_RAW=[]
for _,r in push.iterrows():
    PUSH_RAW.append({'date_str':r['date_str'],'货号':r['goods'],'消耗':r['cost'],
        '直接支付单量':r['do2'],'直接支付金额':r['dgmv'],
        '引导支付单量':r['io2'],'引导支付金额':r['igmv']})

da=defaultdict(lambda:{'cost':0,'do':0,'dg':0,'io':0,'ig':0,'exp':0,'clk':0})
for r in PUSH_RAW:
    g=r['货号']; da[g]['cost']+=r['消耗']; da[g]['do']+=r['直接支付单量']
    da[g]['dg']+=r['直接支付金额']; da[g]['io']+=r['引导支付单量']; da[g]['ig']+=r['引导支付金额']

DETUI_AGG=[]
for g,v in da.items():
    tg=v['dg']+v['ig']; to2=v['do']+v['io']
    roi=tg/v['cost'] if v['cost']>0 else 0; oc=v['cost']/to2 if to2>0 else 0
    DETUI_AGG.append({'货号':g,'总消耗':round(v['cost'],2),'直接支付单量':v['do'],'直接支付金额':v['dg'],
        '引导支付单量':v['io'],'引导支付金额':v['ig'],'总曝光':v['exp'],'总点击':v['clk'],
        '总支付金额':tg,'总支付单量':to2,'综合ROI':round(roi,2),'订单成本':round(oc,2)})
print(f"   PUSH_RAW:{len(PUSH_RAW)} DETUI_AGG:{len(DETUI_AGG)}")

# ============================================================
# 6. 五分类GMV (from extract_fivecat module)
# ============================================================
print("📊 [6/7] 五分类GMV...")
import sys
sys.path.insert(0, '/home/Vic/dewu-reports')
from extract_fivecat import classify_store, STORE_CONFIG

fc = classify_store('美兰', STORE_CONFIG['美兰'])
FIVECAT = {
    'n_pay': fc['n_pay'], 'gmv_pay': round(fc['gmv_pay']),
    'gmv_ret': round(fc['gmv_ret']), 'gmv_cancel': round(fc['gmv_cancel']),
    'gmv_unclear': round(fc['gmv_unclear']), 'gmv_normal': round(fc['gmv_normal']),
    'n_ret': fc['n_ret'], 'n_cancel': fc['n_cancel'], 'n_unclear': fc['n_unclear'],
    'mature_return_rate': fc['mature_return_rate'],
    'as_start': fc.get('as_start'),
    'daily': fc['daily_gmv'],
}
print(f"   支付OID={FIVECAT['n_pay']}, 退货率={FIVECAT['mature_return_rate']}%, as_start={FIVECAT['as_start']}")
if FIVECAT['mature_return_rate'] is not None:
    print(f"   五分类GMV: 留存¥{FIVECAT['gmv_normal']:,} + 退货¥{FIVECAT['gmv_ret']:,} + 取消¥{FIVECAT['gmv_cancel']:,} + 待确认¥{FIVECAT['gmv_unclear']:,} = ¥{FIVECAT['gmv_pay']:,}")

# ============================================================
# 7. 生成内联数据HTML
# ============================================================
print("💾 [7/7] 生成HTML...")

def js_dumps(obj, indent=0):
    """JSON dump as JS const (no quotes on keys)"""
    s = json.dumps(obj, ensure_ascii=False, default=str)
    # For simple arrays of objects, keep compact format
    return s

HTML = f'''<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>美兰运营分析看板</title><script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script><style>
*{{box-sizing:border-box;margin:0;padding:0}}body{{background:#0f1117;color:#e0e0e0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:20px}}h2{{color:#60a5fa;font-size:18px;margin-bottom:12px;border-bottom:1px solid #2d3348;padding-bottom:8px}}.module{{background:#1a1d2e;border-radius:10px;padding:20px;margin-bottom:20px;border:1px solid #2d3348;overflow:hidden}}.filters{{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;align-items:center}}.filters label{{font-size:13px;color:#94a3b8}}.filters select,.filters input{{background:#0f1117;color:#e0e0e0;border:1px solid #2d3348;border-radius:6px;padding:6px 10px;font-size:13px}}.chart-wrap{{position:relative;height:420px}}.card .goods-name{{max-width:none;overflow:visible;white-space:normal;word-break:break-all;display:block;line-height:1.3;font-size:11px}}.card td{{padding:4px 6px;border-bottom:1px solid #1a1d2e;vertical-align:top}}.cards{{display:flex;gap:12px}}.card{{flex:1;background:#0f1117;border-radius:8px;padding:12px;border:1px solid #2d3348;min-width:0;overflow:hidden}}.card h3{{color:#60a5fa;font-size:13px;margin-bottom:8px}}.card table{{width:100%;font-size:11px;border-collapse:collapse;table-layout:fixed}}.card th:nth-child(1){{width:28px}}.card th:nth-child(2){{width:auto}}.card th:nth-child(3){{width:70px;text-align:right}}.card td:nth-child(3){{text-align:right}}.card th{{text-align:left;padding:6px 8px;color:#94a3b8;border-bottom:1px solid #2d3348;position:sticky;top:0;background:#0f1117}}.scroll{{max-height:420px;overflow-y:auto}}.card .rank{{color:#94a3b8;width:24px}}.alert-table{{width:100%;font-size:13px;border-collapse:collapse;margin-top:12px}}.alert-table th{{text-align:left;padding:8px 10px;color:#94a3b8;border-bottom:1px solid #2d3348;background:#0f1117;position:sticky;top:0}}.alert-table td{{padding:8px 10px;border-bottom:1px solid #1a1d2e}}.alert-row{{background:rgba(248,113,113,0.08)}}.delta-up{{color:#f87171;font-weight:bold}}.btn-group{{display:flex;gap:8px;margin-bottom:12px}}.btn-group button{{padding:6px 16px;border-radius:6px;cursor:pointer;background:#0f1117;color:#94a3b8;border:1px solid #2d3348}}.btn-group button.active{{background:#3b82f6;color:#fff;border-color:#3b82f6}}.roi-high{{color:#22c55e;font-weight:bold}}.roi-mid{{color:#f59e0b}}.roi-low{{color:#f87171}}.no-data{{text-align:center;color:#94a3b8;padding:20px}}
</style></head><body>
<h1 style="color:#60a5fa;font-size:22px;margin-bottom:20px">📊 美兰运营分析看板</h1>
<div class="module"><h2>💰 模块〇：支付GMV五分类结构</h2><div id="m0-cards" class="cards" style="margin-bottom:16px;"></div><div class="chart-wrap"><canvas id="m0-chart"></canvas></div><div style="margin-top:8px;font-size:12px;color:#94a3b8;" id="m0-mature"></div><div style="margin-top:4px;font-size:11px;color:#64748b;" id="m0-asinfo"></div></div>
<div class="module"><h2>📈 模块一：GMV / UV / 订单数 趋势</h2><div class="filters"><label>起始月份</label><select id="m1-month-from"><option value="all">全部</option></select><label>终止月份</label><select id="m1-month-to"><option value="all">全部</option></select><label>日期范围</label><input type="date" id="m1-start"><span style="color:#94a3b8">至</span><input type="date" id="m1-end"><label>货号搜索</label><input type="text" id="m1-goods" placeholder="输入货号..."></div><div class="chart-wrap"><canvas id="m1-chart"></canvas></div></div>
<div class="module"><h2>🏆 模块二：TOP20 排行</h2><div class="filters"><label>日期筛选</label><input type="date" id="m2-start"><span style="color:#94a3b8">至</span><input type="date" id="m2-end"><button onclick="updateM2()" style="background:#3b82f6;color:#fff;border:none;padding:6px 16px;border-radius:6px;cursor:pointer">刷新</button></div><div class="cards"><div class="card"><h3>📊 UV TOP20</h3><div class="scroll"><table><thead><tr><th>#</th><th>货号</th><th>UV</th></tr></thead><tbody id="m2-uv"></tbody></table></div></div><div class="card"><h3>💰 支付金额 TOP20</h3><div class="scroll"><table><thead><tr><th>#</th><th>货号</th><th>金额</th></tr></thead><tbody id="m2-gmv"></tbody></table></div></div><div class="card"><h3>📦 支付订单量 TOP20</h3><div class="scroll"><table><thead><tr><th>#</th><th>货号</th><th>订单</th></tr></thead><tbody id="m2-orders"></tbody></table></div></div></div></div>
<div class="module"><h2>⚠️ 模块三：退货率异常警示</h2><div class="btn-group"><button id="m3-btn-7d" class="active" onclick="switchM3('7d')">近7天异常</button><button id="m3-btn-30d" onclick="switchM3('30d')">近30天异常</button><button id="m3-btn-all" onclick="switchM3('all')">全部退货率</button></div><div class="scroll" style="max-height:360px"><table class="alert-table"><thead><tr><th>货号</th><th>历史退货率</th><th>近期退货率</th><th>变化</th><th>总订单</th><th>近期订单</th></tr></thead><tbody id="m3-body"></tbody></table></div><div id="m3-nodata" class="no-data" style="display:none">✅ 未检测到退货率异常款式</div></div>
<div class="module"><h2>📢 模块四：社区投放任务分析</h2><div class="filters"><label>任务月份</label><select id="m4-task-month"><option value="all">全部</option></select><label>发布月份</label><select id="m4-pub-month"><option value="all">全部</option></select><button onclick="updateM4()" style="background:#3b82f6;color:#fff;border:none;padding:6px 16px;border-radius:6px;cursor:pointer">筛选</button></div><table class="alert-table"><thead><tr><th>匹配货号</th><th>动态发布数量</th><th>实际任务金额(¥)</th><th>曝光</th><th>阅读数</th><th>商详访问</th><th>互动数</th></tr></thead><tbody id="m4-body"></tbody><tfoot id="m4-foot"></tfoot></table></div>
<div class="module"><h2>🚀 模块五：得物推投放分析</h2><div class="filters"><label>日期筛选</label><input type="date" id="m5-start"><span style="color:#94a3b8">至</span><input type="date" id="m5-end"><button onclick="updateM5()" style="background:#3b82f6;color:#fff;border:none;padding:6px 16px;border-radius:6px;cursor:pointer">筛选</button></div><table class="alert-table"><thead><tr><th>货号</th><th>总消耗(¥)</th><th>总支付金额(¥)</th><th>总支付单量</th><th>综合ROI</th><th>订单成本(¥)</th><th>建议</th></tr></thead><tbody id="m5-body"></tbody></table></div>
<script>
const DAILY_BRAND = {js_dumps(DAILY_BRAND)};
const DAILY_GOODS = {js_dumps(DAILY_GOODS)};
const ALL_DATES = {js_dumps(ALL_DATES)};
const ALL_BRANDS = {js_dumps([BRAND])};
const ALL_GOODS = {js_dumps(ALL_GOODS)};
const ALL_MONTHS = {js_dumps(ALL_MONTHS)};
const MARKET_MAP = {js_dumps(MARKET_MAP)};
const ANOMALY_7D = {js_dumps(ANOMALY_7D)};
const ANOMALY_30D = {js_dumps(ANOMALY_30D)};
const GOODS_RATE = {js_dumps(GOODS_RATE)};
const TASK_RAW = {js_dumps(TASK_RAW)};
const TASK_MONTHS = {js_dumps(TASK_MONTHS)};
const PUB_MONTHS = {js_dumps(PUB_MONTHS)};
const DETUI_AGG = {js_dumps(DETUI_AGG)};
const PUSH_RAW = {js_dumps(PUSH_RAW)};
const FIVECAT = {js_dumps(FIVECAT)};

let m1Chart = null, m3Mode = '7d';

function fmt(n,d){{d=d||0;return Number(n||0).toLocaleString('zh-CN',{{maximumFractionDigits:d}});}}
function fmtMoney(n){{return '¥'+fmt(n,0);}}

function initM1(){{
  var fs=document.getElementById('m1-month-from'),ts=document.getElementById('m1-month-to');
  ALL_MONTHS.forEach(function(m){{var t=m.replace('-','年')+'月';var o=document.createElement('option');o.value=m;o.textContent=t;fs.appendChild(o);o=document.createElement('option');o.value=m;o.textContent=t;ts.appendChild(o)}});
  if(ALL_MONTHS.length>=3){{fs.value=ALL_MONTHS[ALL_MONTHS.length-3];ts.value=ALL_MONTHS[ALL_MONTHS.length-1]}}
  document.getElementById('m1-start').value=ALL_DATES[Math.max(0,ALL_DATES.length-90)]||ALL_DATES[0];
  document.getElementById('m1-end').value=ALL_DATES[ALL_DATES.length-1];
  [fs,ts,document.getElementById('m1-start'),document.getElementById('m1-end'),document.getElementById('m1-goods')].forEach(function(el){{el.addEventListener('change',updateM1)}});
  document.getElementById('m1-goods').addEventListener('input',updateM1);
  updateM1()
}}

function updateM1(){{
  var mf=document.getElementById('m1-month-from').value,mt=document.getElementById('m1-month-to').value;
  var im=mf!=='all'&&mt!=='all';
  var d1=im?'':document.getElementById('m1-start').value,d2=im?'':document.getElementById('m1-end').value;
  var gf=document.getElementById('m1-goods').value.trim().toUpperCase();
  var raw=gf?DAILY_GOODS.filter(function(r){{return r['商品货号']&&r['商品货号'].toUpperCase().indexOf(gf)>=0}}):DAILY_BRAND;
  if(im)raw=raw.filter(function(r){{return r.date_str>=mf+'-01'&&r.date_str<=mt+'-31'}});
  else if(d1&&d2)raw=raw.filter(function(r){{return r.date_str>=d1&&r.date_str<=d2}});
  var dates=[],gmv=[],uv=[],orders=[];
  if(im){{var mo={{}};raw.forEach(function(r){{var m=r.date_str.substr(0,7);if(!mo[m])mo[m]={{GMV:0,UV:0,orders:0}};mo[m].GMV+=r.GMV;mo[m].UV+=r.UV;mo[m].orders+=r.orders}});var ks=Object.keys(mo).sort();ks.forEach(function(m){{dates.push(m);gmv.push(mo[m].GMV);uv.push(mo[m].UV);orders.push(mo[m].orders)}})}}
  else if(gf){{var agg={{}};raw.forEach(function(r){{if(!agg[r.date_str])agg[r.date_str]={{GMV:0,UV:0,orders:0}};agg[r.date_str].GMV+=r.GMV;agg[r.date_str].UV+=r.UV;agg[r.date_str].orders+=r.orders}});var ks2=Object.keys(agg).sort();ks2.forEach(function(d){{dates.push(d);gmv.push(agg[d].GMV);uv.push(agg[d].UV);orders.push(agg[d].orders)}})}}
  else raw.forEach(function(r){{dates.push(r.date_str);gmv.push(r.GMV);uv.push(r.UV);orders.push(r.orders)}});
  var mkt;if(im){{mkt=dates.map(function(m){{var sum=0,cnt=0;for(var d in MARKET_MAP){{if(d.substr(0,7)===m){{sum+=MARKET_MAP[d];cnt++}}}}return cnt?sum/cnt:null}})}}else{{mkt=dates.map(function(d){{return MARKET_MAP[d]||null}})}}
  if(m1Chart)m1Chart.destroy();
  m1Chart=new Chart(document.getElementById('m1-chart').getContext('2d'),{{type:'bar',data:{{labels:dates,datasets:[
    {{label:'GMV(元)',data:gmv,backgroundColor:'rgba(96,165,250,0.6)',borderColor:'#60a5fa',borderWidth:1,maxBarThickness:16,yAxisID:'y'}},
    {{label:'UV',type:'line',data:uv,borderColor:'#34d399',yAxisID:'y1',tension:0.3,pointRadius:0}},
    {{label:'订单',type:'line',data:orders,borderColor:'#f59e0b',yAxisID:'y2',tension:0.3,pointRadius:0}},
    {{label:'大盘日韩表指数(万元)',type:'line',data:mkt,borderColor:'#f87171',yAxisID:'y',tension:0.3,pointRadius:0,borderDash:[5,5]}}
  ]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{labels:{{color:'#94a3b8',usePointStyle:true}}}}}},scales:{{x:{{ticks:{{color:'#64748b',maxTicksLimit:15}}}},y:{{type:'linear',position:'left',beginAtZero:true,ticks:{{color:'#60a5fa',callback:function(v){{return v>=10000?(v/10000).toFixed(0)+'w':v}}}},grace:'5%'}},y1:{{type:'linear',position:'right',beginAtZero:true,ticks:{{color:'#34d399'}},grid:{{drawOnChartArea:false}},grace:'5%'}},y2:{{type:'linear',position:'right',beginAtZero:true,ticks:{{color:'#f59e0b'}},grid:{{drawOnChartArea:false}},grace:'5%'}}}}}}}})
}}

function initM2(){{document.getElementById('m2-start').value=ALL_DATES[Math.max(0,ALL_DATES.length-30)]||ALL_DATES[0];document.getElementById('m2-end').value=ALL_DATES[ALL_DATES.length-1];updateM2()}}
function updateM2(){{
  var d1=document.getElementById('m2-start').value,d2=document.getElementById('m2-end').value;
  var agg={{}};DAILY_GOODS.forEach(function(r){{if(r.date_str<d1||r.date_str>d2)return;var g=r['商品货号'];if(!agg[g])agg[g]={{UV:0,GMV:0,orders:0}};agg[g].UV+=r.UV;agg[g].GMV+=r.GMV;agg[g].orders+=r.orders}});
  var arr=Object.entries(agg).map(function(e){{return{{goods:e[0],UV:e[1].UV,GMV:e[1].GMV,orders:e[1].orders}}}});
  var tu=arr.slice().sort(function(a,b){{return b.UV-a.UV}}).slice(0,20);
  var tg=arr.slice().sort(function(a,b){{return b.GMV-a.GMV}}).slice(0,20);
  var to=arr.slice().sort(function(a,b){{return b.orders-a.orders}}).slice(0,20);
  function rt(rs,vk,ff){{return rs.length?rs.map(function(r,i){{return'<tr><td class="rank">'+(i+1)+'</td><td class="goods-name">'+r.goods+'</td><td>'+ff(r[vk])+'</td></tr>'}}).join(''):'<tr><td colspan="3" style="text-align:center;color:#94a3b8;">暂无数据</td></tr>'}}
  document.getElementById('m2-uv').innerHTML=rt(tu,'UV',fmt);
  document.getElementById('m2-gmv').innerHTML=rt(tg,'GMV',fmtMoney);
  document.getElementById('m2-orders').innerHTML=rt(to,'orders',fmt)
}}

function switchM3(m){{m3Mode=m;['7d','30d','all'].forEach(function(x){{var b=document.getElementById('m3-btn-'+x);if(b)b.className=x===m?'active':''}});renderM3()}}
function renderM3(){{
  var d=m3Mode==='7d'?ANOMALY_7D:m3Mode==='30d'?ANOMALY_30D:GOODS_RATE;
  var b=document.getElementById('m3-body'),nd=document.getElementById('m3-nodata');
  if(!d||!d.length){{b.innerHTML='';nd.style.display='block';return}}
  nd.style.display='none';
  var rs=m3Mode==='all'?d.map(function(r){{return'<tr><td>'+r['货号']+'</td><td>-</td><td>'+r.return_rate+'%</td><td>-</td><td>'+r.total+'</td><td>'+r.returned+'</td></tr>'}}):d.map(function(r){{var dt=r.delta||0,c=dt>3?'delta-up':(dt<-3?'delta-down':'');return'<tr class="'+(r.is_alert?'alert-row':'')+'"><td>'+r.goods+'</td><td>'+((r.hist_rate||0).toFixed(1))+'%</td><td>'+((r.recent_rate||0).toFixed(1))+'%</td><td class="'+c+'">'+(dt>0?'+':'')+dt.toFixed(1)+'%</td><td>'+r.total_orders+'</td><td>'+r.recent_orders+'</td></tr>'}});
  b.innerHTML=rs.join('')
}}

function initM4(){{var s1=document.getElementById('m4-task-month'),s2=document.getElementById('m4-pub-month');TASK_MONTHS.forEach(function(m){{var o=document.createElement('option');o.value=m;o.textContent=m.replace('-','年')+'月';s1.appendChild(o)}});PUB_MONTHS.forEach(function(m){{var o=document.createElement('option');o.value=m;o.textContent=m.replace('-','年')+'月';s2.appendChild(o)}});updateM4()}}
function updateM4(){{
  var tm=document.getElementById('m4-task-month').value,pm=document.getElementById('m4-pub-month').value;
  var fd=TASK_RAW.filter(function(r){{if(tm!=='all'&&r.task_month!==tm)return false;if(pm!=='all'&&r.pub_month!==pm)return false;return true}});
  var agg={{}};fd.forEach(function(r){{var g=r['匹配货号'];if(!agg[g])agg[g]={{count:0,amount:0,exp:0,reads:0,visits:0,inter:0}};agg[g].count+=r['动态发布数量'];agg[g].amount+=r['实际任务金额'];agg[g].exp+=r['曝光'];agg[g].reads+=r['阅读数'];agg[g].visits+=r['商详访问'];agg[g].inter+=r['互动数']}});
  var es=Object.entries(agg).sort(function(a,b){{return b[1].amount-a[1].amount}});
  var rs=es.map(function(e){{var v=e[1];return'<tr><td>'+e[0]+'</td><td>'+v.count+'</td><td>'+fmtMoney(v.amount)+'</td><td>'+fmt(v.exp)+'</td><td>'+fmt(v.reads)+'</td><td>'+fmt(v.visits)+'</td><td>'+fmt(v.inter)+'</td></tr>'}});
  document.getElementById('m4-body').innerHTML=rs.length?rs.join(''):'<tr><td colspan="7" style="text-align:center;color:#94a3b8;">暂无数据</td></tr>';
  var tot={{count:0,amount:0,exp:0,reads:0,visits:0,inter:0}};fd.forEach(function(r){{tot.count+=r['动态发布数量'];tot.amount+=r['实际任务金额'];tot.exp+=r['曝光'];tot.reads+=r['阅读数'];tot.visits+=r['商详访问'];tot.inter+=r['互动数']}});
  document.getElementById('m4-foot').innerHTML='<tr style="background:#1e2435;font-weight:bold"><td>📊 合计 ('+es.length+'货号)</td><td>'+tot.count+'</td><td>'+fmtMoney(tot.amount)+'</td><td>'+fmt(tot.exp)+'</td><td>'+fmt(tot.reads)+'</td><td>'+fmt(tot.visits)+'</td><td>'+fmt(tot.inter)+'</td></tr>'
}}

function initM5(){{document.getElementById('m5-start').value=ALL_DATES[Math.max(0,ALL_DATES.length-30)]||ALL_DATES[0];document.getElementById('m5-end').value=ALL_DATES[ALL_DATES.length-1];updateM5()}}
function updateM5(){{
  var d1=document.getElementById('m5-start').value,d2=document.getElementById('m5-end').value;
  var agg={{}};PUSH_RAW.forEach(function(r){{if(r.date_str<d1||r.date_str>d2)return;var g=r['货号'];if(!agg[g])agg[g]={{cost:0,do:0,dg:0,io:0,ig:0}};agg[g].cost+=r['消耗'];agg[g].do+=r['直接支付单量'];agg[g].dg+=r['直接支付金额'];agg[g].io+=r['引导支付单量'];agg[g].ig+=r['引导支付金额']}});
  var as=Object.entries(agg).map(function(e){{var v=e[1],tg=v.dg+v.ig,to=v.do+v.io;return{{goods:e[0],cost:v.cost,gmv:tg,orders:to,roi:v.cost>0?tg/v.cost:0,oc:to>0?v.cost/to:0}}}}).sort(function(a,b){{return b.cost-a.cost}});
  var rs=as.map(function(r){{var c=r.roi>=5?'roi-high':(r.roi>=2?'roi-mid':'roi-low'),a=r.roi>=5?'🔥 高效投放':r.roi>=3?'✅ 良好':r.roi>=1?'⚠️ 关注':r.orders>0?'🔻 低ROI':'⛔ 无转化';return'<tr><td>'+r.goods+'</td><td>'+fmtMoney(r.cost)+'</td><td>'+fmtMoney(r.gmv)+'</td><td>'+r.orders+'</td><td class="'+c+'">'+r.roi.toFixed(1)+'</td><td>'+fmtMoney(r.oc)+'</td><td>'+a+'</td></tr>'}}).join('');
  document.getElementById('m5-body').innerHTML=rs
}}

function initM0(){{
  if (typeof FIVECAT === 'undefined' || !FIVECAT.n_pay) return;
  var fc = FIVECAT;
  var fmtM = function(n){{ return '¥'+(n||0).toLocaleString(); }};
  var pct = function(a,b){{ return b>0 ? (a/b*100).toFixed(1)+'%' : 'N/A'; }};
  var cards = document.getElementById('m0-cards');
  cards.innerHTML = [
    {{label:'支付GMV',val:fmtM(fc.gmv_pay),color:'#60a5fa'}},
    {{label:'留存GMV',val:fmtM(fc.gmv_normal),color:'#22c55e'}},
    {{label:'退货GMV',val:fmtM(fc.gmv_ret),color:'#ef4444'}},
    {{label:'取消GMV',val:fmtM(fc.gmv_cancel),color:'#f59e0b'}},
    {{label:'待确认GMV',val:fmtM(fc.gmv_unclear),color:'#6b7280'}},
  ].map(function(c,i){{
    return '<div class=\"card\" style=\"border-left:3px solid '+c.color+'\">'+
      '<div style=\"font-size:11px;color:#94a3b8\">'+c.label+'</div>'+
      '<div style=\"font-size:18px;font-weight:700;color:'+c.color+'\">'+c.val+'</div>'+
      '</div>';
  }}).join('');
  var matureEl = document.getElementById('m0-mature');
  if (fc.mature_return_rate !== null && fc.mature_return_rate !== undefined) {{
    matureEl.innerHTML = '📊 45天成熟退货率: <b>'+fc.mature_return_rate.toFixed(2)+'%</b> | 支付后取消率: <b>'+pct(fc.n_cancel,fc.n_pay)+'</b> | 原因待确认: <b>'+pct(fc.n_unclear,fc.n_pay)+'</b> | 可靠售后起始: <b>'+(fc.as_start||'N/A')+'</b>';
  }} else {{
    matureEl.innerHTML = '⚠️ 售后明细字段不足，退货与取消指标暂不可计算。';
  }}
  if (fc.daily && fc.daily.length > 0) {{
    var dates = fc.daily.map(function(d){{return d.date;}});
    var ctx = document.getElementById('m0-chart').getContext('2d');
    new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: dates,
        datasets: [
          {{label:'留存',data:fc.daily.map(function(d){{return d.gmv_normal||0;}}),backgroundColor:'#22c55e',stack:'gmv'}},
          {{label:'退货',data:fc.daily.map(function(d){{return d.gmv_ret||0;}}),backgroundColor:'#ef4444',stack:'gmv'}},
          {{label:'取消',data:fc.daily.map(function(d){{return d.gmv_cancel||0;}}),backgroundColor:'#f59e0b',stack:'gmv'}},
          {{label:'待确认',data:fc.daily.map(function(d){{return d.gmv_unclear||0;}}),backgroundColor:'#6b7280',stack:'gmv'}},
        ]
      }},
      options: {{
        responsive: true, maintainAspectRatio: false,
        scales: {{
          x: {{ stacked: true, ticks: {{ color: '#94a3b8', maxTicksLimit: 20, maxRotation: 45, font:{{size:9}} }} }},
          y: {{ stacked: true, ticks: {{ color: '#94a3b8', callback: function(v){{return v>=10000?(v/10000).toFixed(0)+'万':v;}} }} }}
        }},
        plugins: {{
          legend: {{ labels: {{ color: '#94a3b8', usePointStyle: true, padding: 10 }} }},
          tooltip: {{ callbacks: {{ label: function(ctx){{return ctx.dataset.label+': ¥'+ctx.raw.toLocaleString();}} }} }}
        }}
      }}
    }});
  }}
}}

window.addEventListener('load',function(){{
  try{{initM0()}}catch(e){{console.error('M0:',e)}}
  try{{initM1()}}catch(e){{console.error('M1:',e)}}
  try{{initM2()}}catch(e){{console.error('M2:',e)}}
  try{{renderM3()}}catch(e){{console.error('M3:',e)}}
  try{{initM4()}}catch(e){{console.error('M4:',e)}}
  try{{initM5()}}catch(e){{console.error('M5:',e)}}
}});
</script></body></html>'''

os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
with open(OUT_FILE, 'w', encoding='utf-8') as f:
    f.write(HTML)

kb = os.path.getsize(OUT_FILE) / 1024
print(f"\n✅ {OUT_FILE} ({kb:.0f}KB inline HTML)")
print(f"   {ALL_DATES[0]} ~ {ALL_DATES[-1]}")
print(f"   DAILY_BRAND:{len(DAILY_BRAND)} DAILY_GOODS:{len(DAILY_GOODS)}")
print(f"   TASK:{len(TASK_RAW)} PUSH:{len(PUSH_RAW)}")
