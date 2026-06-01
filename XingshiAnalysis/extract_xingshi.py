#!/usr/bin/env python3
"""醒狮运营分析看板 — 数据提取脚本（5模块美兰模板，单文件HTML）"""
import pandas as pd
import json, re, os
from datetime import datetime, timedelta
from collections import defaultdict

SRC = '/mnt/e/Obsidian本地仓库/09-数据源/醒狮/醒狮数据源收集表.xlsx'
OUT_FILE = '/home/Vic/dewu-reports/XingshiAnalysis/2026/index.html'
BRAND = '醒狮'

def cn(v):
    """clean number"""
    if pd.isna(v): return 0.0
    s = str(v).strip()
    m = re.match(r'([\d.]+)', s)
    return float(m.group(1)) if m else 0.0

def nd(v):
    """normalize date to YYYY-MM-DD"""
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
    """normalize month to YYYY-MM"""
    d = nd(v)
    return d[:7] if d else None

def ss(v):
    """safe string"""
    if pd.isna(v): return ''
    return str(v).strip()

print(f"📂 {SRC}")

# ============================================================
# 1. 商详访客 → DAILY_BRAND + DAILY_GOODS
# ============================================================
print("📊 [1/6] 商详访客...")
uv = pd.read_excel(SRC, sheet_name='商详访客数据源', engine='calamine')
uv = uv[['日期', '货号', '支付订单金额', '支付订单量', '商详访问指数（UV）']].copy()
uv.columns = ['date_raw', 'goods', 'gmv', 'orders', 'uv']
uv['date_str'] = uv['date_raw'].apply(nd)
uv = uv.dropna(subset=['date_str'])
uv['gmv'] = uv['gmv'].apply(cn)
uv['orders'] = uv['orders'].apply(cn)
uv['uv'] = uv['uv'].apply(cn)
uv = uv[uv['goods'].apply(lambda x: ss(x) not in ['','nan'])]

# DAILY_BRAND
db = uv.groupby('date_str').agg(GMV=('gmv','sum'), UV=('uv','sum'), orders=('orders','sum')).reset_index()
db['brand'] = BRAND
DAILY_BRAND = db[['date_str','brand','GMV','UV','orders']].to_dict('records')

# DAILY_GOODS
dg = uv.groupby(['date_str','goods']).agg(GMV=('gmv','sum'), UV=('uv','sum'), orders=('orders','sum')).reset_index()
dg['brand'] = BRAND
DAILY_GOODS = dg.rename(columns={'goods':'商品货号'})[['date_str','brand','商品货号','GMV','UV','orders']].to_dict('records')

ALL_DATES = sorted(set(r['date_str'] for r in DAILY_BRAND))
ALL_GOODS = sorted(set(r['商品货号'] for r in DAILY_GOODS))
ALL_MONTHS = sorted(set(d[:7] for d in ALL_DATES))
print(f"   {len(DAILY_BRAND)} days, {len(DAILY_GOODS)} day-goods, {ALL_DATES[0]}~{ALL_DATES[-1]}")

# ============================================================
# 2. 交易订单 + 售后 → 退货率
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

# GOODS_RATE
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

# ANOMALY
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
mkt = mkt[['日期', '日韩表大盘成交指数（万元）']].copy()
mkt.columns = ['date_raw','idx']
mkt['date_str'] = mkt['date_raw'].apply(nd)
mkt = mkt.dropna(subset=['date_str'])
mkt['idx'] = mkt['idx'].apply(cn)
MARKET_MAP = {}
for _,r in mkt.iterrows():
    MARKET_MAP[r['date_str']] = round(r['idx']*10000,2)
print(f"   {len(MARKET_MAP)} dates")

# ============================================================
# 4. 社区投放
# ============================================================
print("📢 [4/6] 社区投放...")
comm = pd.read_excel(SRC, sheet_name='社区投放任务', engine='calamine')
comm = comm[['任务月份', '实际任务金额', '动态发布时间', '货号', '曝光', '阅读数', '互动数', '商详访问']].copy()
comm.columns = ['task_m_raw','amount','pub_raw','goods_raw','exposure','reads','interact','visits']
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
hp = pd.read_excel(SRC, sheet_name='货盘表', engine='calamine')
hp = hp[['SPU_ID','得物货号']].copy(); hp.columns=['spu_id','goods']
spu2goods={}
for _,r in hp.iterrows():
    sid=ss(r['spu_id']); g=ss(r['goods'])
    if sid and g: spu2goods[sid]=g

push = pd.read_excel(SRC, sheet_name='得物推数据', engine='calamine')
push = push[['时间','商品ID','消耗(元)','直接支付单量(单)','直接支付金额(元)','引导支付单量(单)','引导支付金额(元)']].copy()
push.columns = ['date_raw','gid','cost','do2','dgmv','io2','igmv']
push['date_str'] = push['date_raw'].apply(nd); push = push.dropna(subset=['date_str'])
for c in ['cost','dgmv','igmv']: push[c]=push[c].apply(cn)
push['do2']=push['do2'].apply(lambda x:int(cn(x)))
push['io2']=push['io2'].apply(lambda x:int(cn(x)))
push['goods']=push['gid'].apply(lambda x:spu2goods.get(ss(x),ss(x)))

PUSH_RAW=[]
for _,r in push.iterrows():
    PUSH_RAW.append({'date_str':r['date_str'],'货号':r['goods'],'消耗':r['cost'],
        '直接支付单量':r['do2'],'直接支付金额':r['dgmv'],
        '引导支付单量':r['io2'],'引导支付金额':r['igmv']})

# DETUI_AGG
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
# 6. 生成 HTML
# ============================================================
print("🖨️  [6/6] 生成HTML...")

jsc = f"""
const DAILY_BRAND = {json.dumps(DAILY_BRAND, ensure_ascii=False)};
const DAILY_GOODS = {json.dumps(DAILY_GOODS, ensure_ascii=False)};
const ALL_DATES = {json.dumps(ALL_DATES, ensure_ascii=False)};
const ALL_BRANDS = {json.dumps([BRAND], ensure_ascii=False)};
const ALL_GOODS = {json.dumps(ALL_GOODS, ensure_ascii=False)};
const ALL_MONTHS = {json.dumps(ALL_MONTHS, ensure_ascii=False)};
const MARKET_MAP = {json.dumps(MARKET_MAP, ensure_ascii=False)};
const ANOMALY_7D = {json.dumps(ANOMALY_7D, ensure_ascii=False)};
const ANOMALY_30D = {json.dumps(ANOMALY_30D, ensure_ascii=False)};
const GOODS_RATE = {json.dumps(GOODS_RATE, ensure_ascii=False)};
const TASK_RAW = {json.dumps(TASK_RAW, ensure_ascii=False)};
const TASK_MONTHS = {json.dumps(TASK_MONTHS, ensure_ascii=False)};
const PUB_MONTHS = {json.dumps(PUB_MONTHS, ensure_ascii=False)};
const DETUI_AGG = {json.dumps(DETUI_AGG, ensure_ascii=False)};
const PUSH_RAW = {json.dumps(PUSH_RAW, ensure_ascii=False)};
"""

# Read template from existing MeilanAnalysis HTML
meilan_html = '/home/Vic/dewu-reports/MeilanAnalysis/2026/index.html'
with open(meilan_html, 'r', encoding='utf-8') as f:
    template = f.read()

# Extract the <script>...</script> block and replace data
# The MeilanAnalysis has data consts then JS functions
# We need to replace the data part but keep the JS logic

# Find where data ends and JS logic begins
# Pattern: after the last data const, before "let m1Chart"
split_marker = '\nlet m1Chart'
js_start = template.index(split_marker)
html_head = template[:js_start]

# Find where data consts start: after "<script>"
script_start = template.index('<script>')
html_prefix = template[:script_start+8]  # <script>
# Data section starts right after <script>

# Build final
final = html_prefix + jsc + split_marker + template[js_start:]

# Also need to replace title from 美兰 to 醒狮
final = final.replace('美兰运营分析看板', '醒狮运营分析看板')
final = final.replace('<h1 style="color:#60a5fa;font-size:22px;margin-bottom:20px">📊 美兰运营分析看板</h1>',
                      '<h1 style="color:#60a5fa;font-size:22px;margin-bottom:20px">📊 醒狮运营分析看板</h1>')

os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
with open(OUT_FILE, 'w', encoding='utf-8') as f:
    f.write(final)

mb = os.path.getsize(OUT_FILE)/(1024*1024)
print(f"\n✅ {OUT_FILE} ({mb:.1f}MB)")
print(f"   {ALL_DATES[0]} ~ {ALL_DATES[-1]}")
print(f"   DAILY_BRAND:{len(DAILY_BRAND)} DAILY_GOODS:{len(DAILY_GOODS)}")
print(f"   TASK:{len(TASK_RAW)} PUSH:{len(PUSH_RAW)}")
