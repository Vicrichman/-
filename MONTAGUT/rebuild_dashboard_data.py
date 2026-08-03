#!/usr/bin/env python3
"""梦特娇 rebuild_dashboard_data.py — M1 LINEAGE FIX v3 (corrected)"""
import pandas as pd, json, os, argparse, hashlib
from datetime import datetime

ap = argparse.ArgumentParser()
ap.add_argument('--output-dir', required=True)
args = ap.parse_args()

SRC = '/home/Vic/.hermes/tmp_data/M1_transaction_lineage_fix_20260802_011247/梦特娇/downloaded.xlsx'
OUT = os.path.join(args.output_dir, 'data')
os.makedirs(OUT, exist_ok=True)

START = datetime.now()
SCRIPT_SHA = hashlib.sha256(open(__file__, 'rb').read()).hexdigest()[:16]
INPUT_SHA = hashlib.sha256(open(SRC, 'rb').read()).hexdigest()[:16]

CAT_MAP = {
    '服装-上衣-Polo衫': 'Polo衫', '服装-上衣-T恤': 'T恤', '服装-上衣-卫衣': '卫衣',
    '服装-上衣-毛衣': '毛衣', '服装-上衣-衬衫': '衬衫', '服装-上衣-针织衫': '针织衫',
    '服装-上衣-羊绒衫': '毛衣', '服装-外套-夹克': '夹克', '服装-外套-棉服': '棉服',
    '服装-外套-羽绒服': '羽绒服', '服装-长裤-休闲裤': '休闲裤', '服装-长裤-牛仔裤': '牛仔裤',
    '服装-短裤-休闲短裤': '短裤', '服装-短裤-牛仔短裤': '短裤',
}

# ============================================================
# Step 1: Build unique OID fact table from 交易订单
# ============================================================
print("Step 1: Reading 交易订单 ...")
orders = pd.read_excel(SRC, sheet_name='交易订单', engine='calamine', dtype=str)
raw_rows = len(orders)

# Column mapping (verified for 梦特娇)
orders['oid'] = orders.iloc[:, 0].astype(str).str.strip()    # 订单号
orders['goods'] = orders.iloc[:, 5].astype(str).str.strip()   # 货号
orders['amount'] = pd.to_numeric(orders.iloc[:, 10], errors='coerce').fillna(0)  # 出价金额（元）
orders['status'] = orders.iloc[:, 20].astype(str).str.strip() # 订单状态
orders['cat_raw'] = orders.iloc[:, 61].astype(str).str.strip() # 类目
orders['date'] = pd.to_datetime(orders.iloc[:, 62], errors='coerce')  # 下单日期

# Filter summary rows
is_summary = orders['oid'].str.contains('汇总|合计|总计|小计', na=False)
orders = orders[~is_summary].copy()
orders = orders[orders['oid'] != '']
orders = orders[orders['oid'] != 'nan']
orders = orders.dropna(subset=['date'])
after_filter = len(orders)

# OID dedup (keep first)
before_dedup = len(orders)
orders = orders.drop_duplicates(subset=['oid'], keep='first')
after_dedup = len(orders)
dup_removed = before_dedup - after_dedup

orders['date_str'] = orders['date'].dt.strftime('%Y-%m-%d')
orders['cat'] = orders['cat_raw'].apply(lambda x: CAT_MAP.get(x, '其他'))

print(f"  Raw: {raw_rows} → filter: {after_filter} → dedup: {after_dedup} (removed {dup_removed} dups)")

# ============================================================
# Step 2: Complement logic
# ============================================================
is_failed = orders['status'].str.contains('交易失败', na=False)
is_closed = orders['status'].str.contains('交易关闭成功', na=False)
is_effective = ~is_failed
is_gsv = ~is_failed & ~is_closed

eff = orders[is_effective]
gsv_only = orders[is_gsv]
closed = orders[is_closed]

n_total = after_dedup
n_fail = int(is_failed.sum())
n_eff = int(is_effective.sum())
n_close = int(is_closed.sum())
n_gsv = int(is_gsv.sum())

gmv_eff = float(round(eff['amount'].sum(), 2))
gmv_gsv = float(round(gsv_only['amount'].sum(), 2))
gmv_close = float(round(closed['amount'].sum(), 2))
gmv_all = float(round(orders['amount'].sum(), 2))

partition_gap = n_total - n_fail - n_close - n_gsv

print(f"  Total OIDs: {n_total}")
print(f"  Failed:     {n_fail}")
print(f"  Effective:  {n_eff} OIDs, ¥{gmv_eff:,.0f}")
print(f"  Closed:     {n_close} OIDs, ¥{gmv_close:,.0f}")
print(f"  GSV:        {n_gsv} OIDs, ¥{gmv_gsv:,.0f}")
print(f"  Partition gap: {partition_gap}")
print(f"  Identity all:  ¥{gmv_all - (orders[is_failed]['amount'].sum() + gmv_close + gmv_gsv):,.0f}")
print(f"  Identity gmv:  ¥{gmv_eff - gmv_close - gmv_gsv:,.0f}")

# Verify targets
targets = [(n_total, 9416), (n_fail, 1119), (n_eff, 8297), (n_close, 3280), (n_gsv, 5017),
           (gmv_eff, 2566442), (gmv_gsv, 1468155), (gmv_close, 1098287)]
all_match = True
for actual, expected in targets:
    ok = abs(actual - expected) < 2
    if not ok:
        print(f"  ❌ Target mismatch: actual={actual}, expected={expected}")
        all_match = False
if all_match: print("  ✅ All targets match")

# ============================================================
# Step 3: Read UV from 商详访客数据源
# ============================================================
print("\nStep 2: Reading 商详访客数据源 (UV) ...")
uv = pd.read_excel(SRC, sheet_name='商详访客数据源', engine='calamine', dtype=str)

# Column mapping (verified)
uv['date'] = pd.to_datetime(uv.iloc[:, 1], errors='coerce')     # 日期
uv['goods'] = uv.iloc[:, 4].astype(str).str.strip()              # 商品货号
uv['uv_val'] = pd.to_numeric(uv.iloc[:, 11], errors='coerce').fillna(0)  # 商详访问人数

uv = uv.dropna(subset=['date', 'goods'])
uv = uv[uv['goods'] != '']
uv = uv[uv['goods'] != 'nan']
uv['date_str'] = uv['date'].dt.strftime('%Y-%m-%d')

uv_raw = len(uv)
uv = uv[uv['uv_val'] > 0]
uv_valid = len(uv)
uv_total = int(uv['uv_val'].sum())

uv_agg = uv.groupby(['date_str', 'goods'])['uv_val'].sum()

print(f"  Sheet: 商详访客数据源")
print(f"  Date col: [1] 日期, Goods col: [4] 商品货号, UV col: [11] 商详访问人数")
print(f"  Raw: {uv_raw} → Valid (UV>0): {uv_valid}")
print(f"  UV total: {uv_total:,}")
print(f"  Date range: {uv['date_str'].min()} ~ {uv['date_str'].max()}")

if uv_valid == 0:
    print("  ❌ FAIL: UV rows = 0")
    raise SystemExit(1)

# ============================================================
# Step 4: Build DAILY_CAT_v2 (date × cat aggregation)
# ============================================================
print("\nStep 3: Building DAILY_CAT_v2 ...")

# GMV/GSV by date+cat
gmv_by_cat = eff.groupby(['date_str', 'cat'])['amount'].sum()
gsv_by_cat = gsv_only.groupby(['date_str', 'cat'])['amount'].sum()
ord_by_cat = eff.groupby(['date_str', 'cat'])['oid'].count()

# Convert groupby results to dict for safe scalar access
gmv_dict = {k: float(v) for k, v in gmv_by_cat.items()}
gsv_dict = {k: float(v) for k, v in gsv_by_cat.items()}
ord_dict = {k: int(v) for k, v in ord_by_cat.items()}

# Collect all keys
all_keys = set(gmv_dict.keys()) | set(uv_agg.index)

daily_cat_v2 = []
dates_set, cats_set, months_set = set(), set(), set()

for key in sorted(all_keys):
    d, c = key if isinstance(key, tuple) else (key, 'UNKNOWN')
    g = gmv_dict.get(key, 0)
    u = int(uv_agg.get(key, 0))
    gsv = gsv_dict.get(key, 0)
    o = ord_dict.get(key, 0)
    
    daily_cat_v2.append([d, c, round(g, 2), u, o, round(gsv, 2)])
    dates_set.add(d)
    cats_set.add(c)
    months_set.add(d[:7])

# ============================================================
# Step 5: Build DAILY_GOODS_v2 (date × cat × goods)
# ============================================================
gmv_by_goods = eff.groupby(['date_str', 'cat', 'goods'])['amount'].sum()
gsv_by_goods = gsv_only.groupby(['date_str', 'cat', 'goods'])['amount'].sum()
ord_by_goods = eff.groupby(['date_str', 'cat', 'goods'])['oid'].count()

gmv_g_dict = {k: float(v) for k, v in gmv_by_goods.items()}
gsv_g_dict = {k: float(v) for k, v in gsv_by_goods.items()}
ord_g_dict = {k: int(v) for k, v in ord_by_goods.items()}

# Merge UV at goods level
uv_goods = uv.groupby(['date_str', 'goods'])['uv_val'].sum()

daily_goods_v2 = []
goods_set = set()

all_gkeys = set(gmv_g_dict.keys()) | set(uv_goods.index)

for key in sorted(all_gkeys):
    d, c, g_name = key if isinstance(key, tuple) and len(key) == 3 else (key[0] if hasattr(key, '__getitem__') else '?', '?', str(key))
    g = gmv_g_dict.get(key, 0)
    u = int(uv_goods.get((d, g_name), 0))
    gsv = gsv_g_dict.get(key, 0)
    o = ord_g_dict.get(key, 0)
    
    daily_goods_v2.append([d, c, g_name, round(g, 2), u, o, round(gsv, 2)])
    goods_set.add(g_name)

# ============================================================
# Step 6: Save all outputs
# ============================================================
print(f"\nStep 4: Saving outputs ...")
print(f"  DAILY_CAT_v2: {len(daily_cat_v2)} records")
print(f"  DAILY_GOODS_v2: {len(daily_goods_v2)} records")
print(f"  Dates: {len(dates_set)}, Cats: {len(cats_set)}, Goods: {len(goods_set)}")

with open(f'{OUT}/DAILY_CAT_v2.json', 'w') as f:
    json.dump(daily_cat_v2, f, ensure_ascii=False, separators=(',', ':'))
with open(f'{OUT}/DAILY_GOODS_v2.json', 'w') as f:
    json.dump(daily_goods_v2, f, ensure_ascii=False, separators=(',', ':'))

# Legacy copies
for name in ['DAILY_CAT_v2', 'DAILY_GOODS_v2']:
    with open(f'{OUT}/{name}.json') as f:
        data = json.load(f)
    with open(f'{OUT}/{name.replace("_v2", "")}.json', 'w') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))

# ALL_ exports
for name, data in [('ALL_DATES', sorted(dates_set)), ('ALL_CATS', sorted(cats_set)),
                    ('ALL_GOODS', sorted(goods_set)), ('ALL_MONTHS', sorted(months_set))]:
    with open(f'{OUT}/{name}.json', 'w') as f:
        json.dump(data, f, ensure_ascii=False)

# ============================================================
# Step 7: 得物推 (unchanged logic)
# ============================================================
print("\nStep 5: Processing 得物推 ...")
push_p = pd.read_excel(SRC, sheet_name='得物推数据源-商品', engine='calamine', dtype=str)
push_raw = []
date_col = [c for c in push_p.columns if '日期' in str(c) or '时间' in str(c)][0]
for _, row in push_p.iterrows():
    d = pd.to_datetime(row[date_col], errors='coerce')
    if pd.isna(d): continue
    push_raw.append({
        'date_str': d.strftime('%Y-%m-%d'),
        '货号': str(row.get('货号', '')).strip(),
        '消耗': round(float(row.get('消耗(元)', 0) or 0), 2),
        '直接支付单量': int(float(row.get('直接支付单量(单)', 0) or 0)),
        '直接支付金额': round(float(row.get('直接支付金额(元)', 0) or 0), 2),
        '引导支付单量': int(float(row.get('引导支付单量(单)', 0) or 0)),
        '引导支付金额': round(float(row.get('引导支付金额(元)', 0) or 0), 2),
        '曝光': int(float(row.get('曝光', 0) or 0)),
        '点击': int(float(row.get('点击', 0) or 0)),
    })

with open(f'{OUT}/PUSH_RAW.json', 'w') as f:
    json.dump(push_raw, f, ensure_ascii=False, separators=(',', ':'))

push_df = pd.DataFrame(push_raw)
if len(push_df) > 0:
    agg = push_df.groupby('货号').agg(
        总消耗=('消耗', 'sum'), 直接支付单量=('直接支付单量', 'sum'),
        直接支付金额=('直接支付金额', 'sum'), 引导支付单量=('引导支付单量', 'sum'),
        引导支付金额=('引导支付金额', 'sum'), 总曝光=('曝光', 'sum'), 总点击=('点击', 'sum'),
    ).reset_index()
    agg['总支付金额'] = agg['直接支付金额'] + agg['引导支付金额']
    agg['总支付单量'] = agg['直接支付单量'] + agg['引导支付单量']
    agg['综合ROI'] = (agg['总支付金额'] / agg['总消耗']).round(2).fillna(0)
    detui_agg = []
    for _, row in agg.sort_values('综合ROI', ascending=False).iterrows():
        detui_agg.append({
            '货号': str(row['货号']), '总消耗': round(float(row['总消耗']), 2),
            '直接支付单量': int(row['直接支付单量']), '直接支付金额': round(float(row['直接支付金额']), 2),
            '引导支付单量': int(row['引导支付单量']), '引导支付金额': round(float(row['引导支付金额']), 2),
            '总曝光': int(row['总曝光']), '总点击': int(row['总点击']),
            '总支付金额': round(float(row['总支付金额']), 2), '总支付单量': int(row['总支付单量']),
            '综合ROI': float(row['综合ROI']),
        })
    with open(f'{OUT}/DETUI_AGG.json', 'w') as f:
        json.dump(detui_agg, f, ensure_ascii=False, separators=(',', ':'))
    print(f"  PUSH_RAW: {len(push_raw)}, DETUI_AGG: {len(detui_agg)}")

# ============================================================
# Step 8: Verification
# ============================================================
total_cat_gmv = sum(x[2] for x in daily_cat_v2)
total_cat_gsv = sum(x[5] for x in daily_cat_v2)

print(f"\n{'='*50}")
print(f"VERIFICATION")
print(f"{'='*50}")
print(f"DAILY_CAT GMV: ¥{total_cat_gmv:,.0f} (target: ¥2,566,442)")
print(f"DAILY_CAT GSV: ¥{total_cat_gsv:,.0f} (target: ¥1,468,155)")

gmv_ok = abs(total_cat_gmv - 2566442) < 2
gsv_ok = abs(total_cat_gsv - 1468155) < 2

print(f"GMV: {'✅' if gmv_ok else '❌'}")
print(f"GSV: {'✅' if gsv_ok else '❌'}")
print(f"UV rows: {uv_valid} {'✅' if uv_valid > 0 else '❌'}")

elapsed = (datetime.now() - START).total_seconds()
print(f"\nElapsed: {elapsed:.0f}s")
print(f"Script SHA: {SCRIPT_SHA}")
print(f"Input SHA: {INPUT_SHA}")

if gmv_ok and gsv_ok and uv_valid > 0:
    print("\n✅ PASS")
else:
    print("\n❌ FAIL")
    raise SystemExit(1)
