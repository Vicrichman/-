#!/usr/bin/env python3
"""
C7 v2: schema_adapter — raw_decode parse + passthrough original D + append const variables
Key difference from v1: D object is preserved as-is (no re-serialization), only const vars are appended
"""
import json, os, sys, hashlib, math, re
from collections import defaultdict

# ── B patch: --output-dir + positional arg support ──
_OUTPUT_DIR = None
_POS_ARGS = []
for _a in sys.argv[1:]:
    if _a == '--output-dir':
        _OUTPUT_DIR = None  # flag, value follows
    elif _OUTPUT_DIR is None and _a != '--output-dir':
        if _a.startswith('--'):
            _OUTPUT_DIR = None
        elif _POS_ARGS or not _a.startswith('--'):
            # Check if previous arg was --output-dir
            _OUTPUT_DIR = _a if len(_POS_ARGS) == 0 and any(sys.argv[_i] == '--output-dir' for _i in range(len(sys.argv))) else _OUTPUT_DIR
    else:
        _POS_ARGS.append(_a)

# Re-parse properly
_OUTPUT_DIR = None
_data_js_in = None
_i = 1
while _i < len(sys.argv):
    if sys.argv[_i] == '--output-dir' and _i + 1 < len(sys.argv):
        _OUTPUT_DIR = sys.argv[_i + 1]
        _i += 2
    elif not sys.argv[_i].startswith('--'):
        if _data_js_in is None:
            _data_js_in = sys.argv[_i]
        _i += 1
    else:
        _i += 1

DATA_JS_IN = _data_js_in or sys.argv[1]
if _OUTPUT_DIR:
    assert os.path.isdir(_OUTPUT_DIR), f"output-dir not found: {_OUTPUT_DIR}"
    OUTPUT = os.path.join(_OUTPUT_DIR, "data_full.js")
else:
    OUTPUT = sys.argv[2] if len(sys.argv) > 2 else DATA_JS_IN.replace('.js', '_full.js')

def clean_float(v):
    if v is None: return None
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return None
    return v

def recursive_clean(obj):
    if isinstance(obj, dict):
        return {k: recursive_clean(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [recursive_clean(v) for v in obj]
    elif isinstance(obj, float):
        return clean_float(obj)
    return obj

def strict_json_dumps(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(',', ':'), allow_nan=False)

# ── Read original file ──
with open(DATA_JS_IN, 'r', encoding='utf-8') as f:
    raw_content = f.read()

# ── Parse D using raw_decode ──
if 'var D=' in raw_content:
    d_var_start = raw_content.index('var D=')
else:
    d_var_start = raw_content.index('var D =')
brace_start = raw_content.index('{', d_var_start)

decoder = json.JSONDecoder()
try:
    D, end_pos = decoder.raw_decode(raw_content, brace_start)
except json.JSONDecodeError as e:
    print(f"FATAL: Cannot parse D object: {e}", file=sys.stderr)
    sys.exit(1)

# Verify termination
rest = raw_content[end_pos:].lstrip()
if not rest.startswith(';'):
    print(f"FATAL: D object not terminated with ; got: {rest[:20]}", file=sys.stderr)
    sys.exit(1)

# Clean only for variable extraction, not for output
D_clean = recursive_clean(D)
print(f"✅ D parsed: {len(D_clean)} keys", file=sys.stderr)

months = D_clean.get('months', [])
orders_monthly = D_clean.get('orders_monthly', {})

# ── Brand standardization ──
def std_brand(b):
    b = str(b).strip()
    if not b or b in ('-','/','nan','None'): return '未分类品牌'
    u = b.upper()
    if 'CASIO' in u or '卡西欧' in u: return '卡西欧'
    if 'COACH' in u or '蔻驰' in u: return '蔻驰'
    if 'SWATCH' in u or '斯沃琪' in u: return '斯沃琪'
    if 'GIVENCHY' in u or '纪梵希' in u: return '纪梵希'
    return b

all_brands = set()
for b in list(orders_monthly.keys()) + list(D_clean.get('uv_monthly',{}).keys()):
    all_brands.add(std_brand(b))
ALL_BRANDS = sorted(all_brands, key=lambda x: (
    0 if x=='卡西欧' else 1 if x=='蔻驰' else 2 if x=='斯沃琪' else 
    3 if x=='未分类品牌' else 4
))

# ── Extract all variables from D ──
DAILY_BRAND = []
all_dates_set = set()
for brand, daily_list in D_clean.get('uv_daily_by_brand', {}).items():
    sb = std_brand(brand)
    for d in daily_list:
        ds = d.get('date', '')
        if ds:
            all_dates_set.add(ds)
            DAILY_BRAND.append({'date_str': ds, 'brand': sb,
                'GMV': round(d.get('gmv', 0), 2),
                'UV': round(d.get('uv', 0), 1), 'orders': d.get('orders', 0)})

ALL_DATES = sorted(all_dates_set)
ALL_MONTHS = sorted(set(d[:7] for d in ALL_DATES))

MARKET_MAP = {}
if 'market_monthly_avg' in D_clean:
    for m, v in D_clean['market_monthly_avg'].items():
        MARKET_MAP[m] = round(v, 2)

ANOMALY_7D = D_clean.get('anomaly_7d', [])
ANOMALY_30D = D_clean.get('anomaly_30d', [])

TASK_RAW = D_clean.get('comm_tasks', [])
if not TASK_RAW: TASK_RAW = []
TASK_MONTHS = sorted(set(D_clean.get('comm_monthly', {}).keys()))

PUSH_RAW = D_clean.get('push_daily', [])
if not PUSH_RAW: PUSH_RAW = []

DETUI_AGG = []
if 'push_by_spu' in D_clean:
    for brand, spus in D_clean['push_by_spu'].items():
        for spu_data in spus:
            cost = spu_data.get('total_cost', 0) or sum(d.get('cost', 0) for d in spu_data.get('daily', []))
            gmv = spu_data.get('total_gmv', 0) or sum(d.get('gmv', 0) for d in spu_data.get('daily', []))
            orders = spu_data.get('total_orders', 0) or sum(d.get('orders', 0) for d in spu_data.get('daily', []))
            roi = round(gmv / cost, 1) if cost > 0 else 0
            DETUI_AGG.append({'spu': spu_data.get('spu', ''), '货号': spu_data.get('huohao', ''),
                '消耗': round(cost, 2), '支付金额': round(gmv, 2), '支付单量': orders, '综合ROI': roi})

ALL_GOODS = []
DAILY_GOODS = []
for brand, spus in D_clean.get('uv_spu_data', {}).items():
    sb = std_brand(brand)
    for s in spus:
        gn = s.get('spu', '')
        hh = s.get('huohao', '')
        if gn and gn not in ALL_GOODS: ALL_GOODS.append(gn)
        for d in s.get('daily', []):
            ds = d.get('date', '')
            if ds:
                DAILY_GOODS.append({'date_str': ds, 'brand': sb, '货号': hh or gn,
                    'GMV': round(d.get('gmv', 0), 2), 'UV': round(d.get('uv', 0), 1), 'orders': d.get('orders', 0)})

GOODS_RATE = D_clean.get('goods_rate', [])
PUB_MONTHS = sorted(set(
    r.get('pub_date', '')[:7] for r in D_clean.get('comm_tasks', [])
    if r.get('pub_date') and len(str(r.get('pub_date', ''))) >= 7))

MARKET_BAG_MAP = {}
if 'market_bag_monthly_avg' in D_clean:
    for m, v in D_clean['market_bag_monthly_avg'].items():
        MARKET_BAG_MAP[m] = round(v, 2)

# ── Build output: preserve original D text, append const variables ──
# Use original text from start to end_pos+1 (includes closing }), then add ;
d_text = raw_content[d_var_start:end_pos+1] + ';'

consts = [
    ("DAILY_BRAND", DAILY_BRAND), ("DAILY_GOODS", DAILY_GOODS),
    ("ALL_DATES", ALL_DATES), ("ALL_BRANDS", ALL_BRANDS),
    ("ALL_GOODS", ALL_GOODS), ("ALL_MONTHS", ALL_MONTHS),
    ("MARKET_MAP", MARKET_MAP), ("MARKET_BAG_MAP", MARKET_BAG_MAP),
    ("ANOMALY_7D", ANOMALY_7D), ("ANOMALY_30D", ANOMALY_30D),
    ("GOODS_RATE", GOODS_RATE), ("TASK_RAW", TASK_RAW),
    ("TASK_MONTHS", TASK_MONTHS), ("PUB_MONTHS", PUB_MONTHS),
    ("DETUI_AGG", DETUI_AGG), ("PUSH_RAW", PUSH_RAW),
]

names = [c[0] for c in consts]
dupes = [n for n in names if names.count(n) > 1]
if dupes:
    print(f"FATAL: duplicate const names: {dupes}", file=sys.stderr)
    sys.exit(1)

parts = [d_text]
for name, data in consts:
    parts.append(f"const {name} = {strict_json_dumps(data)};")

content_out = '\n'.join(parts)

# ── Validate ──
nan_count = len(re.findall(r'(?<!")(?<!\w)NaN(?!\w)(?!")', content_out))
inf_count = len(re.findall(r'(?<!")(?<!\w)Infinity(?!\w)(?!")', content_out))
if nan_count > 0 or inf_count > 0:
    print(f"FATAL: output NaN={nan_count} Inf={inf_count}", file=sys.stderr)
    sys.exit(1)

# ── Atomic write ──
tmp = OUTPUT + ".tmp"
with open(tmp, 'w', encoding='utf-8') as f:
    f.write(content_out)
    f.flush()
    os.fsync(f.fileno())

with open(tmp, 'r') as f:
    written = f.read()
assert len(written) > 0, "FATAL: empty"
assert "var D={" in written, "FATAL: missing var D"
sha = hashlib.sha256(written.encode()).hexdigest()
os.rename(tmp, OUTPUT)

print(f"✅ {OUTPUT}: {len(written)} bytes, SHA256={sha[:16]}...", file=sys.stderr)
print(f"   DAILY_BRAND: {len(DAILY_BRAND)}, ALL_BRANDS: {len(ALL_BRANDS)}", file=sys.stderr)
print(f"   PUSH_RAW: {len(PUSH_RAW)}, TASK_RAW: {len(TASK_RAW)}", file=sys.stderr)
print(f"   NaN=0 Inf=0", file=sys.stderr)
