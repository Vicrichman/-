"""
喜过数据提取 - 适配蓄势模板（5模块）
输出 data.js 用于 HTML 看板
"""
import pandas as pd
import json
from datetime import datetime, timedelta
from collections import defaultdict

PATH_XIGUO = '/mnt/e/Obsidian本地仓库/09-数据源/喜过数据源收集表.xlsx'

def clean_num(v):
    """提取纯数字"""
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

def map_brand(name):
    name = str(name).upper()
    if 'CASIO' in name or '卡西欧' in name: return '卡西欧'
    if 'COACH' in name or '蔻驰' in name: return '蔻驰'
    return 'OTHER'

print("=== 喜过数据提取（蓄势模板） ===\n")

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

months_set = set(d[:7] for d in ALL_DATES)
ALL_MONTHS = sorted(months_set)

print(f"  DAILY_BRAND: {len(DAILY_BRAND)} rows, ALL_DATES: {len(ALL_DATES)}, ALL_GOODS: {len(ALL_GOODS)}")

# ===== 2. 大盘指数 → MARKET_MAP =====
print("[2/6] 大盘指数...")
df_mkt = pd.read_excel(PATH_XIGUO, sheet_name='大盘指数', engine='calamine',
    usecols=['日期','日韩表大盘成交GMV指数','单肩包大盘成交GMV指数'])

MARKET_MAP = {}
for _, r in df_mkt.iterrows():
    d = norm_date(r['日期'])
    if not d: continue
    v = clean_num(r['日韩表大盘成交GMV指数'])
    if v > 0: MARKET_MAP[d] = round(v, 2)

print(f"  MARKET_MAP: {len(MARKET_MAP)} days")

# ===== 3. 交易订单 + 售后订单 → 退货率 =====
print("[3/6] 退货率计算...")
df_orders = pd.read_excel(PATH_XIGUO, sheet_name='交易订单', engine='calamine',
    usecols=['订单号','订单状态','货号','品牌','买家支付时间'])
df_returns = pd.read_excel(PATH_XIGUO, sheet_name='售后订单', engine='calamine',
    usecols=['订单号','商品货号','品牌','价格（元）'])

# 成功订单
success_orders = set()
for _, r in df_orders.iterrows():
    status = str(r['订单状态']).strip() if pd.notna(r['订单状态']) else ''
    if '成功' in status or '完成' in status or '收货' in status:
        oid = str(r['订单号']).strip()
        success_orders.add(oid)

# 售后关联
returned_set = set()
for _, r in df_returns.iterrows():
    oid = str(r['订单号']).strip()
    returned_set.add(oid)

# 按货号统计
goods_stats = defaultdict(lambda: {'total':0, 'returned':0})
for _, r in df_orders.iterrows():
    goods = str(r['货号']).strip() if pd.notna(r['货号']) else ''
    if not goods or goods == 'nan': continue
    oid = str(r['订单号']).strip()
    goods_stats[goods]['total'] += 1
    if oid in returned_set:
        goods_stats[goods]['returned'] += 1

GOODS_RATE = []
for goods, stats in goods_stats.items():
    if stats['total'] >= 5:
        rate = round(stats['returned']/stats['total']*100, 2)
        GOODS_RATE.append({'货号': goods, 'total': stats['total'], 'returned': stats['returned'], 'return_rate': rate})

GOODS_RATE.sort(key=lambda x: x['return_rate'], reverse=True)
print(f"  GOODS_RATE: {len(GOODS_RATE)} items")

# 退货异动（简化版：取最近30天对比）
ANOMALY_7D = GOODS_RATE[:30]  # 简化
ANOMALY_30D = GOODS_RATE[:50]

# ===== 4. 社区投放任务 → TASK_RAW =====
print("[4/6] 社区投放任务...")
df_comm = pd.read_excel(PATH_XIGUO, sheet_name='社区投放任务', engine='calamine',
    usecols=list(range(30)))  # 读前30列

TASK_RAW = []
task_months_set = set()
pub_months_set = set()

for _, r in df_comm.iterrows():
    task_month = norm_month(r.get('任务月份'))
    pub_date = norm_date(r.get('动态发布时间'))
    
    # 货号在 col[28] 和 col[29]
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

print(f"  TASK_RAW: {len(TASK_RAW)} rows, TASK_MONTHS: {TASK_MONTHS}")

# ===== 5. 得物推 → DETUI_AGG =====
print("[5/6] 得物推数据...")

# 构建 SPU_ID → 货号 映射（货盘表 + 交易订单 fallback）
def safe_str(v):
    if pd.isna(v): return ''
    try: return str(int(v))
    except: return str(v).strip()

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
print(f"  SPU→货号映射: {len(id2goods)} 条")

df_push = pd.read_excel(PATH_XIGUO, sheet_name='得物推数据', engine='calamine')

push_agg = defaultdict(lambda: {'消耗':0,'直接支付单量':0,'直接支付金额':0,'引导支付单量':0,'引导支付金额':0,'曝光':0,'点击':0})

for _, r in df_push.iterrows():
    spu_id = safe_str(r['商品ID'])
    goods = id2goods.get(spu_id, '')
    if not goods: continue
    
    cost = clean_num(r.get('消耗(元)'))
    direct_orders = int(clean_num(r.get('直接支付单量(单)')))
    direct_gmv = clean_num(r.get('直接支付金额(元)'))
    indirect_orders = int(clean_num(r.get('引导支付单量(单)')))
    indirect_gmv = clean_num(r.get('引导支付金额(元)'))
    exposure = int(clean_num(r.get('曝光(次)')))
    clicks = int(clean_num(r.get('点击(次)')))
    
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
        '货号': goods,
        '总消耗': round(v['消耗'], 2),
        '直接支付单量': v['直接支付单量'],
        '直接支付金额': round(v['直接支付金额'], 2),
        '引导支付单量': v['引导支付单量'],
        '引导支付金额': round(v['引导支付金额'], 2),
        '总曝光': v['曝光'],
        '总点击': v['点击'],
        '总支付金额': round(total_gmv, 2),
        '总支付单量': total_orders,
        '综合ROI': roi,
        '订单成本': order_cost,
    })

# 按ROI降序排列
DETUI_AGG.sort(key=lambda x: x['综合ROI'], reverse=True)

print(f"  DETUI_AGG: {len(DETUI_AGG)} items")

# ===== 6. 输出 data.js =====
print("[6/6] 生成 data.js...")

OUTPUT = '/home/Vic/dewu-reports/XiguoAnalysis/2026/data.js'

# 用 json.dumps 避免 JS 变量污染
parts = []
parts.append(f"const DAILY_BRAND = {json.dumps(DAILY_BRAND, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const DAILY_GOODS = {json.dumps(DAILY_GOODS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ALL_DATES = {json.dumps(ALL_DATES, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ALL_BRANDS = {json.dumps(ALL_BRANDS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ALL_GOODS = {json.dumps(ALL_GOODS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ALL_MONTHS = {json.dumps(ALL_MONTHS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const MARKET_MAP = {json.dumps(MARKET_MAP, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ANOMALY_7D = {json.dumps(ANOMALY_7D, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const ANOMALY_30D = {json.dumps(ANOMALY_30D, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const GOODS_RATE = {json.dumps(GOODS_RATE, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const TASK_RAW = {json.dumps(TASK_RAW, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const TASK_MONTHS = {json.dumps(TASK_MONTHS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const PUB_MONTHS = {json.dumps(PUB_MONTHS, ensure_ascii=False, separators=(',',':'))};")
parts.append(f"const DETUI_AGG = {json.dumps(DETUI_AGG, ensure_ascii=False, separators=(',',':'))};")

content = '\n'.join(parts)

import os
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(content)

size_kb = os.path.getsize(OUTPUT) / 1024
print(f"\n✅ data.js 已生成: {OUTPUT} ({size_kb:.1f} KB)")
print(f"   ALL_DATES: {ALL_DATES[0]} ~ {ALL_DATES[-1]}")
print(f"   ALL_MONTHS: {ALL_MONTHS}")
