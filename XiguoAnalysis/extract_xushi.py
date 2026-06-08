"""
喜过数据提取 v2 - 修复异常检测 + 蓄势模板
"""
import pandas as pd
import json
from datetime import datetime, timedelta
from collections import defaultdict

PATH_XIGUO = '/mnt/e/Obsidian本地仓库/09-数据源/喜过数据源收集表.xlsx'

def clean_num(v):
    if pd.isna(v): return 0
    s = str(v).replace(',','').replace('¥','').strip()
    try: return float(s)
    except:
        import re
        m = re.search(r'[\d.]+', s)
        return float(m.group()) if m else 0

def serial_to_date(v):
    if isinstance(v, (int, float)) and v > 40000:
        return datetime(1899, 12, 30) + timedelta(days=int(v))
    if isinstance(v, datetime): return v
    if pd.isna(v): return None
    s = str(v).strip()
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y-%m-%d %H:%M:%S']:
        try: return datetime.strptime(s, fmt)
        except: pass
    return None

def norm_date(v):
    d = serial_to_date(v)
    return d.strftime('%Y-%m-%d') if d else ''

def norm_month(v):
    d = serial_to_date(v)
    return d.strftime('%Y-%m') if d else ''

def safe_str(v):
    if pd.isna(v): return ''
    try: return str(int(v))
    except: return str(v).strip()

def map_brand(name):
    name = str(name).upper()
    if 'CASIO' in name or '卡西欧' in name: return '卡西欧'
    if 'COACH' in name or '蔻驰' in name: return '蔻驰'
    return 'OTHER'

print("=== 喜过数据提取 v2 ===\n")

# ===== 1. 商详访客 → DAILY_BRAND + DAILY_GOODS =====
print("[1/6] 商详访客数据...")
df_uv = pd.read_excel(PATH_XIGUO, sheet_name='商详访客数据', engine='calamine',
    usecols=['日期','商品名称','货号','品牌名称','支付订单金额','支付订单量','商详访问指数（UV）'])

daily_brand = defaultdict(lambda: {'GMV':0,'UV':0,'orders':0})
daily_goods = defaultdict(lambda: {'GMV':0,'UV':0,'orders':0})
dates_set = set()
goods_set = set()

for _, r in df_uv.iterrows():
    d = norm_date(r['日期'])
    if not d: continue
    brand = map_brand(r['品牌名称'])
    if brand == 'OTHER': continue
    goods = str(r['货号']).strip() if pd.notna(r['货号']) else str(r['商品名称']).strip()
    if not goods or goods == 'nan': continue
    gmv = clean_num(r['支付订单金额'])
    uv = clean_num(r['商详访问指数（UV）'])
    orders = int(clean_num(r['支付订单量']))
    
    dates_set.add(d)
    goods_set.add(goods)
    
    key_b = (d, brand)
    daily_brand[key_b]['GMV'] += gmv
    daily_brand[key_b]['UV'] += uv
    daily_brand[key_b]['orders'] += orders
    
    key_g = (d, brand, goods)
    daily_goods[key_g]['GMV'] += gmv
    daily_goods[key_g]['UV'] += uv
    daily_goods[key_g]['orders'] += orders

DAILY_BRAND = []
for (d, brand), v in sorted(daily_brand.items()):
    DAILY_BRAND.append({'date_str': d, 'brand': brand, 'GMV': round(v['GMV'],2), 'UV': v['UV'], 'orders': v['orders']})

DAILY_GOODS = []
for (d, brand, goods), v in sorted(daily_goods.items()):
    DAILY_GOODS.append({'date_str': d, 'brand': brand, '商品货号': goods, 'GMV': round(v['GMV'],2), 'UV': v['UV'], 'orders': v['orders']})

ALL_DATES = sorted(dates_set)
ALL_GOODS = sorted(goods_set)
ALL_BRANDS = ['卡西欧', '蔻驰']
ALL_MONTHS = sorted(set(d[:7] for d in ALL_DATES))

print(f"  DAILY_BRAND: {len(DAILY_BRAND)}, ALL_DATES: {len(ALL_DATES)}")

# ===== 2. 大盘指数 =====
print("[2/6] 大盘指数...")
df_mkt = pd.read_excel(PATH_XIGUO, sheet_name='大盘指数', engine='calamine',
    usecols=['日期','日韩表大盘成交GMV指数','单肩包大盘成交GMV指数'])

MARKET_MAP = {}
MARKET_BAG_MAP = {}
for _, r in df_mkt.iterrows():
    d = norm_date(r['日期'])
    if not d: continue
    v = clean_num(r['日韩表大盘成交GMV指数'])
    vb = clean_num(r['单肩包大盘成交GMV指数'])
    if v > 0: MARKET_MAP[d] = round(v, 2)
    if vb > 0: MARKET_BAG_MAP[d] = round(vb, 2)

print(f"  MARKET: {len(MARKET_MAP)} days, BAG: {len(MARKET_BAG_MAP)} days")

# ===== 3. 退货率 + 真实异动计算 =====
print("[3/6] 退货率+异动...")
df_orders = pd.read_excel(PATH_XIGUO, sheet_name='交易订单', engine='calamine',
    usecols=['订单号','订单状态','货号','品牌','买家支付时间'])
df_returns = pd.read_excel(PATH_XIGUO, sheet_name='售后订单', engine='calamine',
    usecols=['订单号'])

returned_set = set(str(r['订单号']).strip() for _, r in df_returns.iterrows() if pd.notna(r['订单号']))

# 收集成功订单的日期+货号信息
order_records = []  # (date_str, goods, is_returned)
for _, r in df_orders.iterrows():
    status = str(r['订单状态']).strip() if pd.notna(r['订单状态']) else ''
    if not ('成功' in status or '完成' in status or '收货' in status): continue
    goods = str(r['货号']).strip() if pd.notna(r['货号']) else ''
    if not goods or goods == 'nan': continue
    d = norm_date(r['买家支付时间'])
    if not d: continue
    oid = str(r['订单号']).strip()
    order_records.append((d, goods, oid in returned_set))

# 按货号统计总体退货率
goods_stats = defaultdict(lambda: {'total':0, 'returned':0})
for d, goods, is_ret in order_records:
    goods_stats[goods]['total'] += 1
    if is_ret: goods_stats[goods]['returned'] += 1

GOODS_RATE = []
for goods, stats in goods_stats.items():
    if stats['total'] >= 3:
        rate = round(stats['returned']/stats['total']*100, 2)
        GOODS_RATE.append({'货号': goods, 'total': stats['total'], 'returned': stats['returned'], 'return_rate': rate})
GOODS_RATE.sort(key=lambda x: x['return_rate'], reverse=True)

# 周度异动计算
if ALL_DATES:
    latest_date = ALL_DATES[-1]
    week7_ago = (datetime.strptime(latest_date, '%Y-%m-%d') - timedelta(days=7)).strftime('%Y-%m-%d')
    week14_ago = (datetime.strptime(latest_date, '%Y-%m-%d') - timedelta(days=14)).strftime('%Y-%m-%d')
    month_ago = (datetime.strptime(latest_date, '%Y-%m-%d') - timedelta(days=30)).strftime('%Y-%m-%d')
    month60_ago = (datetime.strptime(latest_date, '%Y-%m-%d') - timedelta(days=60)).strftime('%Y-%m-%d')

    def compute_anomaly(all_orders):
        goods_period = defaultdict(lambda: {'hist_total':0,'hist_ret':0,'recent_total':0,'recent_ret':0})
        for d, goods, is_ret in all_orders:
            if d >= week7_ago:
                goods_period[goods]['recent_total'] += 1
                if is_ret: goods_period[goods]['recent_ret'] += 1
            elif d >= week14_ago:
                goods_period[goods]['hist_total'] += 1
                if is_ret: goods_period[goods]['hist_ret'] += 1
        
        results = []
        for goods, v in goods_period.items():
            if v['recent_total'] < 2 or v['hist_total'] < 2: continue
            hist_rate = v['hist_ret']/v['hist_total']*100
            recent_rate = v['recent_ret']/v['recent_total']*100
            delta = recent_rate - hist_rate
            is_alert = (recent_rate >= 10 and hist_rate < 5) or (hist_rate > 0 and recent_rate >= hist_rate*1.5 and recent_rate >= 8)
            results.append({
                'goods': goods, 'hist_rate': round(hist_rate,1), 'recent_rate': round(recent_rate,1),
                'delta': round(delta,1), 'total_orders': v['hist_total']+v['recent_total'],
                'recent_orders': v['recent_total'], 'is_alert': is_alert
            })
        results.sort(key=lambda x: abs(x['delta']), reverse=True)
        return results[:50]

    # 30天异常
    def compute_anomaly_30d(all_orders):
        goods_period = defaultdict(lambda: {'hist_total':0,'hist_ret':0,'recent_total':0,'recent_ret':0})
        for d, goods, is_ret in all_orders:
            if d >= month_ago:
                goods_period[goods]['recent_total'] += 1
                if is_ret: goods_period[goods]['recent_ret'] += 1
            elif d >= month60_ago:
                goods_period[goods]['hist_total'] += 1
                if is_ret: goods_period[goods]['hist_ret'] += 1
        
        results = []
        for goods, v in goods_period.items():
            if v['recent_total'] < 3 or v['hist_total'] < 3: continue
            hist_rate = v['hist_ret']/v['hist_total']*100
            recent_rate = v['recent_ret']/v['recent_total']*100
            delta = recent_rate - hist_rate
            is_alert = (recent_rate >= 10 and hist_rate < 5) or (hist_rate > 0 and recent_rate >= hist_rate*1.5 and recent_rate >= 8)
            results.append({
                'goods': goods, 'hist_rate': round(hist_rate,1), 'recent_rate': round(recent_rate,1),
                'delta': round(delta,1), 'total_orders': v['hist_total']+v['recent_total'],
                'recent_orders': v['recent_total'], 'is_alert': is_alert
            })
        results.sort(key=lambda x: abs(x['delta']), reverse=True)
        return results[:50]

    ANOMALY_7D = compute_anomaly(order_records)
    ANOMALY_30D = compute_anomaly_30d(order_records)
else:
    ANOMALY_7D = []
    ANOMALY_30D = []

print(f"  GOODS_RATE: {len(GOODS_RATE)}, ANOMALY_7D: {len(ANOMALY_7D)}, ANOMALY_30D: {len(ANOMALY_30D)}")

# ===== 4. 社区投放任务 =====
print("[4/6] 社区投放任务...")
df_comm = pd.read_excel(PATH_XIGUO, sheet_name='社区投放任务', engine='calamine',
    usecols=list(range(30)))

TASK_RAW = []
task_months_set = set()
pub_months_set = set()

for _, r in df_comm.iterrows():
    task_month = norm_month(r.get('任务月份'))
    pub_date = norm_date(r.get('动态发布时间'))
    
    goods = str(r.iloc[28]).strip() if pd.notna(r.iloc[28]) and len(df_comm.columns) > 28 else ''
    goods2 = str(r.iloc[29]).strip() if pd.notna(r.iloc[29]) and len(df_comm.columns) > 29 else ''
    
    all_goods = []
    for g in [goods, goods2]:
        if g and g != 'nan':
            for part in g.split(','):
                part = part.strip()
                if part and part != '暂无':
                    all_goods.append(part)
    if not all_goods: continue
    
    amount = clean_num(r.get('实际任务金额')) or clean_num(r.get('任务金额'))
    exposure = clean_num(r.get('曝光')) if '曝光' in df_comm.columns else 0
    reads = clean_num(r.get('阅读数')) if '阅读数' in df_comm.columns else 0
    visits = clean_num(r.get('商详访问')) if '商详访问' in df_comm.columns else 0
    interactions = clean_num(r.get('互动数')) if '互动数' in df_comm.columns else 0
    
    per_goods = amount / len(all_goods) if all_goods else amount
    pub_month = pub_date[:7] if pub_date else ''
    if task_month: task_months_set.add(task_month)
    if pub_month: pub_months_set.add(pub_month)
    
    for g in all_goods:
        TASK_RAW.append({
            '匹配货号': g,
            'task_month': task_month,
            'pub_month': pub_month,
            '动态发布数量': 1,
            '实际任务金额': round(per_goods, 2),
            '曝光': int(exposure / len(all_goods)) if all_goods else int(exposure),
            '阅读数': int(reads / len(all_goods)) if all_goods else int(reads),
            '商详访问': int(visits / len(all_goods)) if all_goods else int(visits),
            '互动数': int(interactions / len(all_goods)) if all_goods else int(interactions),
        })

TASK_MONTHS = sorted(task_months_set)
PUB_MONTHS = sorted(pub_months_set)
print(f"  TASK_RAW: {len(TASK_RAW)}, MONTHS: {len(TASK_MONTHS)}")

# ===== 5. 得物推 =====
print("[5/6] 得物推数据...")

id2goods = {}
df_hp = pd.read_excel(PATH_XIGUO, sheet_name='货盘表', engine='calamine')
for _, r in df_hp.iterrows():
    spu_id = safe_str(r['SPU_ID'])
    goods = str(r['货号']).strip() if pd.notna(r['货号']) else ''
    if spu_id and goods and spu_id.isdigit():
        id2goods[spu_id] = goods

df_orders_fb = pd.read_excel(PATH_XIGUO, sheet_name='交易订单', engine='calamine', usecols=['spuID','货号'])
for _, r in df_orders_fb.iterrows():
    spu_id = safe_str(r['spuID'])
    goods = str(r['货号']).strip() if pd.notna(r['货号']) else ''
    if spu_id and goods and spu_id.isdigit() and spu_id not in id2goods:
        id2goods[spu_id] = goods

df_push = pd.read_excel(PATH_XIGUO, sheet_name='得物推数据', engine='calamine')

# Raw push records with dates for M5 date filter
PUSH_RAW = []
push_agg = defaultdict(lambda: {'消耗':0,'直接支付单量':0,'直接支付金额':0,'引导支付单量':0,'引导支付金额':0,'曝光':0,'点击':0})

for _, r in df_push.iterrows():
    spu_id = safe_str(r['商品ID2'])
    goods = id2goods.get(spu_id, '')
    if not goods: continue
    
    d = norm_date(r['时间'])
    cost = clean_num(r.get('消耗(元)'))
    direct_orders = int(clean_num(r.get('直接支付单量(单)')))
    direct_gmv = clean_num(r.get('直接支付金额(元)'))
    indirect_orders = int(clean_num(r.get('引导支付单量(单)')))
    indirect_gmv = clean_num(r.get('引导支付金额(元)'))
    exposure = int(clean_num(r.get('曝光(次)')))
    clicks = int(clean_num(r.get('点击(次)')))
    
    PUSH_RAW.append({'date_str': d, '货号': goods, '消耗': cost, '直接支付单量': direct_orders,
        '直接支付金额': direct_gmv, '引导支付单量': indirect_orders, '引导支付金额': indirect_gmv,
        '曝光': exposure, '点击': clicks})
    
    push_agg[goods]['消耗'] += cost
    push_agg[goods]['直接支付单量'] += direct_orders
    push_agg[goods]['直接支付金额'] += direct_gmv
    push_agg[goods]['引导支付单量'] += indirect_orders
    push_agg[goods]['引导支付金额'] += indirect_gmv
    push_agg[goods]['曝光'] += exposure
    push_agg[goods]['点击'] += clicks

DETUI_AGG = []
for goods, v in push_agg.items():
    total_gmv = v['直接支付金额'] + v['引导支付金额']
    total_orders = v['直接支付单量'] + v['引导支付单量']
    roi = round(total_gmv / v['消耗'], 2) if v['消耗'] > 0 else 0
    order_cost = round(v['消耗'] / total_orders, 2) if total_orders > 0 else 0
    DETUI_AGG.append({
        '货号': goods, '总消耗': round(v['消耗'],2), '直接支付单量': v['直接支付单量'],
        '直接支付金额': round(v['直接支付金额'],2), '引导支付单量': v['引导支付单量'],
        '引导支付金额': round(v['引导支付金额'],2), '总曝光': v['曝光'], '总点击': v['点击'],
        '总支付金额': round(total_gmv,2), '总支付单量': total_orders,
        '综合ROI': roi, '订单成本': order_cost,
    })
DETUI_AGG.sort(key=lambda x: x['综合ROI'], reverse=True)

print(f"  DETUI_AGG: {len(DETUI_AGG)}, PUSH_RAW: {len(PUSH_RAW)}")

# ===== 6. 输出 =====
print("[6/6] 生成 data.js...")
OUTPUT = '/home/Vic/dewu-reports/XiguoAnalysis/2026/data.js'

parts = []
parts.append(f"const DAILY_BRAND = {json.dumps(DAILY_BRAND, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const DAILY_GOODS = {json.dumps(DAILY_GOODS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ALL_DATES = {json.dumps(ALL_DATES, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ALL_BRANDS = {json.dumps(ALL_BRANDS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ALL_GOODS = {json.dumps(ALL_GOODS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ALL_MONTHS = {json.dumps(ALL_MONTHS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const MARKET_MAP = {json.dumps(MARKET_MAP, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const MARKET_BAG_MAP = {json.dumps(MARKET_BAG_MAP, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ANOMALY_7D = {json.dumps(ANOMALY_7D, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ANOMALY_30D = {json.dumps(ANOMALY_30D, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const GOODS_RATE = {json.dumps(GOODS_RATE, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const TASK_RAW = {json.dumps(TASK_RAW, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const TASK_MONTHS = {json.dumps(TASK_MONTHS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const PUB_MONTHS = {json.dumps(PUB_MONTHS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const DETUI_AGG = {json.dumps(DETUI_AGG, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const PUSH_RAW = {json.dumps(PUSH_RAW, ensure_ascii=False, separators=(',',':'))};")

content = '\n'.join(parts)
import os
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(content)

size_kb = os.path.getsize(OUTPUT) / 1024
print(f"\n✅ data.js: {OUTPUT} ({size_kb:.0f} KB)")
print(f"   ALL_DATES: {ALL_DATES[0]} ~ {ALL_DATES[-1]}")
