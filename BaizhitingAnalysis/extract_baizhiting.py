#!/usr/bin/env python3
"""柏治廷运营分析看板 — 数据提取+模板注入脚本"""
import pandas as pd
import json, re, os
from datetime import datetime, timedelta
from collections import defaultdict

SRC = '/home/Vic/.hermes/tmp_data/柏治廷数据.xlsx'
HTML_SRC = '/home/Vic/dewu-reports/BaizhitingAnalysis/2026/index.html'
HTML_OUT = '/home/Vic/dewu-reports/BaizhitingAnalysis/2026/index.html'

def cn(v):
    if pd.isna(v): return 0.0
    s = str(v).strip()
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
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S', '%Y.%m.%d']:
        try: return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except: pass
    return s[:10] if len(s)>=10 else None

def ss(v):
    if pd.isna(v): return ''
    return str(v).strip()

print(f"📂 {SRC}")

# ============================================================
# 1. 商详访客 → DAILY_BRAND + DAILY_GOODS (by date+cat)
# ============================================================
print("📊 [1/4] 商详访客...")
uv = pd.read_excel(SRC, sheet_name='商详访客数据', engine='calamine')
uv = uv[['日期', '商品货号', '支付金额', '支付订单数', '商详访问人数', '类目名称']].copy()
uv.columns = ['date_raw', 'goods', 'gmv', 'orders', 'uv', 'cat_raw']
uv['date_str'] = uv['date_raw'].apply(nd)
uv = uv.dropna(subset=['date_str'])
uv['gmv'] = uv['gmv'].apply(cn)
uv['orders'] = uv['orders'].apply(cn)
uv['uv'] = uv['uv'].apply(cn)
uv = uv[uv['goods'].apply(lambda x: ss(x) not in ['','nan'])]

# Category mapping (simplify from full hierarchy)
def map_cat(name):
    name = str(name).strip()
    if '永生花' in name: return '永生花'
    if '盲盒' in name: return '盲盒'
    if '香薰' in name or '蜡烛' in name or '烛台' in name: return '香薰礼盒'
    if '汽车' in name: return '香薰礼盒'
    return '其他'
uv['cat'] = uv['cat_raw'].apply(map_cat)
uv = uv[uv['cat'] != '其他']
print(f"   类目分布: {uv['cat'].value_counts().to_dict()}")

# DAILY_BRAND: date+cat level
db = uv.groupby(['date_str','cat']).agg(GMV=('gmv','sum'), UV=('uv','sum'), orders=('orders','sum')).reset_index()
DAILY_BRAND = db.to_dict('records')

# DAILY_GOODS: date+cat+goods level  
dg = uv.groupby(['date_str','cat','goods']).agg(GMV=('gmv','sum'), UV=('uv','sum'), orders=('orders','sum')).reset_index()
DAILY_GOODS = dg.to_dict('records')

ALL_DATES = sorted(set(r['date_str'] for r in DAILY_BRAND))
ALL_GOODS = sorted(set(r['goods'] for r in DAILY_GOODS))
ALL_MONTHS = sorted(set(d[:7] for d in ALL_DATES))
ALL_CATS = sorted(set(r['cat'] for r in DAILY_BRAND))
MARKET_CATS = ALL_CATS.copy()
print(f"   {len(DAILY_BRAND)} brand-days, {len(DAILY_GOODS)} day-goods, {ALL_DATES[0]}~{ALL_DATES[-1]}")

# ============================================================
# 2. 交易订单 + 售后 → 退货率 
# ============================================================
print("📦 [2/4] 交易+售后...")
ord_df = pd.read_excel(SRC, sheet_name='交易订单', engine='calamine')
ord_df = ord_df[['货号', '数量', '出价金额（元）', '订单状态', '下单日期']].copy()
ord_df.columns = ['goods', 'qty', 'amount', 'status', 'date_raw']
ord_df['date_str'] = ord_df['date_raw'].apply(nd)
ord_df = ord_df.dropna(subset=['date_str','goods'])
ord_df['amount'] = ord_df['amount'].apply(cn)
ord_df['qty'] = ord_df['qty'].apply(lambda x: int(cn(x)) if cn(x)>0 else 1)

# 柏治廷可能没有售后订单 sheet — 检查
try:
    ret_df = pd.read_excel(SRC, sheet_name='售后订单', engine='calamine')
    ret_df = ret_df[['商品货号', '售后原因', '买家申请售后时间']].copy()
    ret_df.columns = ['goods', 'reason', 'apply_time']
    ret_df['date_str'] = ret_df['apply_time'].apply(nd)
    ret_df = ret_df.dropna(subset=['date_str','goods'])
    ret_returns = ret_df[ret_df['reason'].astype(str).str.contains('退货|退款|换货', na=False)]
    print(f"   售后订单: {len(ret_df)} total, {len(ret_returns)} returns")
except:
    print("   ⚠️ 无售后订单 sheet, 使用交易关闭成功作为退货")
    ret_returns = ord_df[ord_df['status'].astype(str).str.contains('关闭', na=False)].copy()
    ret_returns['goods'] = ret_returns['goods']

gr = defaultdict(lambda:{'total':0,'returned':0})
for _,r in ord_df.iterrows():
    g=ss(r['goods'])
    if g: gr[g]['total']+=r['qty']
for _,r in ret_returns.iterrows():
    g=ss(r['goods'])
    if g: gr[g]['returned']+=1

GOODS_RATE=[]
for g,v in gr.items():
    if v['total']>=2:
        rate=round(v['returned']/v['total']*100,2)
        GOODS_RATE.append({'goods':g,'total':v['total'],'returned':v['returned'],'return_rate':rate})
GOODS_RATE.sort(key=lambda x:x['return_rate'],reverse=True)

today = ALL_DATES[-1]
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
        if rt<2: continue
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
# 3. 大盘指数 → MARKET_MAP (nested: date->cat->value)
# ============================================================
print("📈 [3/4] 大盘...")
mkt = pd.read_excel(SRC, sheet_name='大盘指数', engine='calamine')
# col[0]=月份, col[1]=永生花, col[2]=盲盒, col[3]=香薰礼盒, col[4]=店铺GMV, col[5]=UV
mkt.columns = ['month_raw','eternal','blindbox','scent','store_gmv','store_uv']
mkt['date_str'] = mkt['month_raw'].apply(nd)
mkt = mkt.dropna(subset=['date_str'])
for c in ['eternal','blindbox','scent']:
    mkt[c] = mkt[c].apply(cn)

MARKET_MAP = {}
for _,r in mkt.iterrows():
    MARKET_MAP[r['date_str']] = {
        '永生花': round(r['eternal']*10000, 2),
        '盲盒': round(r['blindbox']*10000, 2),
        '香薰礼盒': round(r['scent']*10000, 2),
    }

# M4: Monthly category data and index values
M4_MONTHS = sorted(set(d[:7] for d in ALL_DATES))
M4_TOTAL_GMV = []
m4_monthly = {}
for m in M4_MONTHS:
    md = [r for r in DAILY_BRAND if r['date_str'][:7] == m]
    total = sum(r['GMV'] for r in md)
    M4_TOTAL_GMV.append(int(total))
    for r in md:
        c = r['cat']
        if c not in m4_monthly:
            m4_monthly[c] = {}
        m4_monthly[c][m] = m4_monthly[c].get(m, 0) + r['GMV']

# M4_CATS: all categories across all data
all_m4_cats = sorted(set(r['cat'] for r in DAILY_BRAND))
M4_CATS = all_m4_cats

# M4_CAT_DATA: {cat: [monthly GMVs]}
M4_CAT_DATA = {}
for c in M4_CATS:
    M4_CAT_DATA[c] = [int(m4_monthly.get(c, {}).get(m, 0)) for m in M4_MONTHS]

# M4_IDX_VALS: {cat: [monthly index values]} - use MARKET_MAP monthly averages
M4_IDX_VALS = {}
IDX_MAP = {'永生花': '永生花', '盲盒': '盲盒', '香薰礼盒': '香薰礼盒'}
for display_cat, idx_key in IDX_MAP.items():
    vals = []
    for m in M4_MONTHS:
        mo_vals = []
        for d, cats in MARKET_MAP.items():
            if d[:7] == m and idx_key in cats:
                mo_vals.append(cats[idx_key])
        vals.append(round(sum(mo_vals)/len(mo_vals), 2) if mo_vals else 0)
    M4_IDX_VALS[display_cat] = vals

M4_CAT_COLORS = {'永生花':'#60a5fa','盲盒':'#f59e0b','香薰礼盒':'#34d399'}

print(f"   MARKET_MAP:{len(MARKET_MAP)} dates, M4_MONTHS:{len(M4_MONTHS)}")

# ============================================================
# 4. 注入到HTML模板
# ============================================================
print("💾 [4/4] 注入HTML...")

with open(HTML_SRC, 'r', encoding='utf-8') as f:
    html = f.read()

def js_dumps(obj):
    return json.dumps(obj, ensure_ascii=False, default=str, separators=(',',':'))

replacements = {
    'DAILY_BRAND': DAILY_BRAND,
    'DAILY_GOODS': DAILY_GOODS,
    'ALL_DATES': ALL_DATES,
    'ALL_GOODS': ALL_GOODS,
    'ALL_MONTHS': ALL_MONTHS,
    'ALL_CATS': ALL_CATS,
    'MARKET_CATS': MARKET_CATS,
    'MARKET_MAP': MARKET_MAP,
    'GOODS_RATE': GOODS_RATE,
    'ANOMALY_7D': ANOMALY_7D,
    'ANOMALY_30D': ANOMALY_30D,
    'M4_CATS': M4_CATS,
    'M4_CAT_DATA': M4_CAT_DATA,
    'M4_IDX_VALS': M4_IDX_VALS,
    'M4_MONTHS': M4_MONTHS,
    'M4_TOTAL_GMV': M4_TOTAL_GMV,
    'M4_CAT_COLORS': M4_CAT_COLORS,
}

for var_name, data in replacements.items():
    pattern = rf'(const {var_name} = )(\[.*?\]|\{{.*?\}})(;)\s*'
    replacement = f'const {var_name} = {js_dumps(data)};'
    new_html, count = re.subn(pattern, replacement, html, count=1, flags=re.DOTALL)
    if count > 0:
        html = new_html
        print(f"   ✅ {var_name}: {len(str(data))} chars")
    else:
        print(f"   ⚠️ {var_name}: NOT FOUND in HTML")

with open(HTML_OUT, 'w', encoding='utf-8') as f:
    f.write(html)

kb = os.path.getsize(HTML_OUT) / 1024 / 1024
print(f"\n✅ {HTML_OUT} ({kb:.1f}MB)")
print(f"   {ALL_DATES[0]} ~ {ALL_DATES[-1]}, {len(ALL_GOODS)} goods, {ALL_CATS} cats")
