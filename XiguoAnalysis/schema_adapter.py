"""
喜过Schema适配器 v2 — 输出完整数据层
1. var D = {...}  — extract+supplement的完整D对象(供app.js M1-M5使用)
2. const 变量     — 16个页面控件变量(品牌选择器/月份/日期等)
"""

import json, os, sys, hashlib
from collections import defaultdict
from datetime import datetime, timedelta

DATA_JS_IN = sys.argv[1]
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else DATA_JS_IN.replace('.js', '_full.js')

# 1. 读取完整D对象
with open(DATA_JS_IN, 'r', encoding='utf-8') as f:
    content = f.read()
json_str = content[content.index('{'):content.rindex('};')+1]
D = json.loads(json_str)
months = D.get('months', [])
orders_monthly = D.get('orders_monthly', {})

# 2. 品牌标准化(动态，保留全部10品牌)
def std_brand(b):
    b = str(b).strip()
    if not b or b in ('-','/','nan','None'): return '未分类品牌'
    u = b.upper()
    if 'CASIO' in u or '卡西欧' in u: return '卡西欧'
    if 'COACH' in u or '蔻驰' in u: return '蔻驰'
    if 'SWATCH' in u or '斯沃琪' in u: return '斯沃琪'
    if 'GIVENCHY' in u or '纪梵希' in u: return '纪梵希'
    return b

# 收集所有品牌(从orders_monthly + uv_monthly)
all_brands = set()
for b in list(orders_monthly.keys()) + list(D.get('uv_monthly',{}).keys()):
    all_brands.add(std_brand(b))
ALL_BRANDS = sorted(all_brands, key=lambda x: (
    0 if x=='卡西欧' else 1 if x=='蔻驰' else 2 if x=='斯沃琪' else 
    3 if x=='未分类品牌' else 4
))

# 3. 从D.uv_daily_by_brand提取DAILY_BRAND (品牌日聚合)
DAILY_BRAND = []
all_dates_set = set()
for brand, daily_list in D.get('uv_daily_by_brand', {}).items():
    sb = std_brand(brand)
    for d in daily_list:
        date_str = d.get('date', '')
        if date_str:
            all_dates_set.add(date_str)
            DAILY_BRAND.append({
                'date_str': date_str,
                'brand': sb,
                'GMV': round(d.get('gmv', 0), 2),
                'UV': round(d.get('uv', 0), 1),
                'orders': d.get('orders', 0)
            })

# 4. ALL_DATES, ALL_MONTHS
ALL_DATES = sorted(all_dates_set)
ALL_MONTHS = sorted(set(d[:7] for d in ALL_DATES))

# 5. MARKET_MAP — 从D.market_monthly_avg
MARKET_MAP = {}
if 'market_monthly_avg' in D:
    for m, v in D['market_monthly_avg'].items():
        MARKET_MAP[m] = round(v, 2)

# 6. ANOMALY_7D/30D — 从D
ANOMALY_7D = D.get('anomaly_7d', [])
ANOMALY_30D = D.get('anomaly_30d', [])

# 7. TASK数据
TASK_RAW = D.get('comm_tasks', [])
TASK_MONTHS = sorted(set(D.get('comm_monthly', {}).keys()))

# 8. PUSH数据
PUSH_RAW = []
for r in D.get('push_daily', []):
    PUSH_RAW.append({
        'date_str': r.get('date', ''),
        '货号': r.get('huohao', ''),
        '消耗': r.get('cost', 0),
        '直接支付单量': r.get('direct_orders', 0),
        '直接支付金额': r.get('direct_gmv', 0),
        '引导支付单量': r.get('indirect_orders', 0),
        '引导支付金额': r.get('indirect_gmv', 0)
    })

# 9. DETUI_AGG — 从D
DETUI_AGG = []
if 'push_by_spu' in D:
    for brand, spus in D['push_by_spu'].items():
        for spu_data in spus:
            cost = spu_data.get('total_cost', 0) or sum(d.get('cost', 0) for d in spu_data.get('daily', []))
            gmv = spu_data.get('total_gmv', 0) or sum(d.get('gmv', 0) for d in spu_data.get('daily', []))
            orders = spu_data.get('total_orders', 0) or sum(d.get('orders', 0) for d in spu_data.get('daily', []))
            roi = round(gmv / cost, 1) if cost > 0 else 0
            DETUI_AGG.append({
                'spu': spu_data.get('spu', ''),
                '货号': spu_data.get('huohao', ''),
                '消耗': round(cost, 2),
                '支付金额': round(gmv, 2),
                '支付单量': orders,
                '综合ROI': roi
            })

# 10. ALL_GOODS, DAILY_GOODS
ALL_GOODS = []
DAILY_GOODS = []
for brand, spus in D.get('uv_spu_data', {}).items():
    sb = std_brand(brand)
    for s in spus:
        goods_name = s.get('spu', '')
        huohao = s.get('huohao', '')
        if goods_name and goods_name not in ALL_GOODS:
            ALL_GOODS.append(goods_name)
        for d in s.get('daily', []):
            date_str = d.get('date', '')
            if date_str:
                DAILY_GOODS.append({
                    'date_str': date_str,
                    'brand': sb,
                    '商品货号': huohao or goods_name,
                    'GMV': round(d.get('gmv', 0), 2),
                    'UV': round(d.get('uv', 0), 1),
                    'orders': d.get('orders', 0)
                })

# 11. GOODS_RATE, PUB_MONTHS
GOODS_RATE = D.get('goods_rate', [])
PUB_MONTHS = sorted(set(
    r.get('pub_date', '')[:7] 
    for r in D.get('comm_tasks', []) 
    if r.get('pub_date', '') and len(str(r.get('pub_date', ''))) >= 7
))

# MARKET_BAG_MAP
MARKET_BAG_MAP = {}
if 'market_bag_monthly_avg' in D:
    for m, v in D['market_bag_monthly_avg'].items():
        MARKET_BAG_MAP[m] = round(v, 2)

# 12. 构建输出: var D = {...} \n const XXXXX = [...];
parts = []

# 先输出 var D
d_json = json.dumps(D, ensure_ascii=False, separators=(',', ':'))
parts.append(f"var D={d_json};\n")

# 再输出所有const变量
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

content_out = '\n'.join(parts)

# 安全原子写入
tmp = OUTPUT + ".tmp"
with open(tmp, 'w', encoding='utf-8') as f:
    f.write(content_out)
    f.flush()
    os.fsync(f.fileno())

with open(tmp, 'r') as f:
    written = f.read()
assert len(written) > 0, "FATAL: empty"
assert written.startswith("var D={"), f"FATAL: bad format: {written[:30]}"
_bo = written.count('{')
_bc = written.count('}')
assert _bo == _bc, f"FATAL: bracket mismatch {{{_bo} vs {_bc}}}"
sha = hashlib.sha256(written.encode()).hexdigest()
os.rename(tmp, OUTPUT)

print(f"✅ {OUTPUT}: {len(written)} bytes, SHA256={sha[:16]}...", file=sys.stderr)
print(f"   DAILY_BRAND: {len(DAILY_BRAND)}, DAILY_GOODS: {len(DAILY_GOODS)}", file=sys.stderr)
print(f"   ALL_DATES: {len(ALL_DATES)}, ALL_BRANDS: {len(ALL_BRANDS)}", file=sys.stderr)
print(f"   ANOMALY_7D: {len(ANOMALY_7D)}, PUSH_RAW: {len(PUSH_RAW)}, TASK_RAW: {len(TASK_RAW)}", file=sys.stderr)
print(f"   DETUI_AGG: {len(DETUI_AGG)}", file=sys.stderr)
