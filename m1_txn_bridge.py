#!/usr/bin/env python3
"""
M1 Transaction Bridge v3 — 五店统一交易数据提取模块
单一OID事实表驱动所有M1消费者：
- DAILY_BRAND, DAILY_GOODS, ALL_DATES, ALL_GOODS, ALL_MONTHS
- FIVECAT (五分类OID/GMV from OID table, not from extract_fivecat.py)
- GSV = complement(交易失败 ∪ 交易关闭成功)

v3 changes:
- Contract_lib loaded via DEWU_CONTRACT_LIB env var (no relative file-path guess)
- 柏治廷 cat mapping dynamically built from 商详访客 '类目名称' column
- No static baizhiting_cat_map.json dependency
- New goods → 未分类类目 with full OID/GMV preservation
"""
import pandas as pd
import json, os, sys
from pathlib import Path
from collections import defaultdict

# ── contract_lib: loaded via env var or known default ──
def _resolve_contract_lib():
    """Resolve contract_lib path. Priority: DEWU_CONTRACT_LIB env → default."""
    env_path = os.environ.get('DEWU_CONTRACT_LIB', '')
    if env_path and os.path.isdir(env_path):
        return env_path
    # Default: known skill scripts directory
    default = str(Path.home() / '.hermes' / 'skills' / 'dewu' / 'dewu-operations-analysis' / 'scripts')
    if os.path.isdir(default):
        return default
    raise RuntimeError(
        f"contract_lib not found. Set DEWU_CONTRACT_LIB env var. Tried: {default}"
    )

_CONTRACT_PATH = _resolve_contract_lib()
sys.path.insert(0, _CONTRACT_PATH)
import contract_lib as cl

# Log actual contract_lib source for auditability
_cl_file = getattr(cl, '__file__', os.path.join(_CONTRACT_PATH, 'contract_lib.py'))
if os.path.exists(_cl_file):
    import hashlib as _hl
    _cl_sha = _hl.sha256(open(_cl_file, 'rb').read()).hexdigest()[:16]
    print(f"  contract_lib: {_cl_file} (SHA {_cl_sha})")
else:
    print(f"  contract_lib: {_CONTRACT_PATH} (SHA unavailable)")

# 五分类标签
CAT_NORMAL = '正常留存'
CAT_RETURN = '明确实质退货'
CAT_CANCEL = '支付后取消'
CAT_UNCLEAR = '原因待确认关闭'
CAT_OTHER = '其他异常'

UNCLASSIFIED_CAT = '未分类类目'


def load_oid_fact_table(xlsx_path, store_name):
    """
    从交易订单Sheet构建唯一OID事实表。
    返回 DataFrame with all needed columns.
    """
    cfg = cl.load_store_config(store_name)

    df = pd.read_excel(xlsx_path, sheet_name='交易订单', engine='calamine', dtype=str)

    # Dynamic column detection
    col_map = {}
    for i, c in enumerate(df.columns):
        cs = str(c).strip()
        if '订单号' in cs and 'oid' not in col_map:
            col_map['oid'] = i
        if '订单状态' in cs and 'status' not in col_map:
            col_map['status'] = i
        if '出价金额' in cs and 'amount' not in col_map:
            col_map['amount'] = i
        if cs == '品牌' and 'brand' not in col_map:
            col_map['brand'] = i
        if ('货号' in cs or '商品货号' in cs) and 'goods' not in col_map:
            col_map['goods'] = i
        if ('spuID' in cs or 'spuid' in cs or 'SPUID' in cs) and 'spuid' not in col_map:
            col_map['spuid'] = i
        if '下单日期' in cs and 'date' not in col_map:
            col_map['date'] = i

    # Validate required columns
    required = ['oid', 'status', 'amount']
    missing = [k for k in required if k not in col_map]
    if missing:
        raise ValueError(
            f"{store_name}: missing columns: {missing}. Found: {list(df.columns[:10])}"
        )

    df['_oid'] = df.iloc[:, col_map['oid']].astype(str).str.strip()
    df['_status'] = df.iloc[:, col_map['status']].astype(str).str.strip()
    df['_amount'] = pd.to_numeric(df.iloc[:, col_map['amount']], errors='coerce').fillna(0)
    df['_brand'] = (
        df.iloc[:, col_map.get('brand', 0)].astype(str).str.strip()
        if 'brand' in col_map
        else store_name
    )
    df['_goods'] = (
        df.iloc[:, col_map.get('goods', 0)].astype(str).str.strip()
        if 'goods' in col_map
        else ''
    )
    df['_spuid'] = (
        df.iloc[:, col_map.get('spuid', 0)].astype(str).str.strip()
        if 'spuid' in col_map
        else ''
    )
    df['_date'] = (
        pd.to_datetime(df.iloc[:, col_map['date']], errors='coerce')
        if 'date' in col_map
        else pd.NaT
    )

    # Filter: remove summary rows
    is_summary = df['_oid'].str.contains('汇总|合计|总计|小计', na=False)
    df = df[~is_summary].copy()
    df = df[df['_oid'] != '']
    df = df[df['_oid'] != 'nan']
    df = df.dropna(subset=['_date'])

    # OID dedup
    before = len(df)
    df = df.drop_duplicates(subset=['_oid'], keep='first')
    after = len(df)
    print(f"  OID: {before}→{after} (removed {before - after} dups)")

    # Date string
    df['_date_str'] = df['_date'].dt.strftime('%Y-%m-%d')

    # Brand standardization
    df['_brand_std'] = df['_brand'].apply(
        lambda x: '未分类品牌' if x in ('', '-', '/', 'nan', 'None', 'null') else x
    )

    # === COMPLEMENT LOGIC ===
    fail_m = df['_status'].str.contains('交易失败', na=False)
    close_m = df['_status'].str.contains('交易关闭成功', na=False)

    df['_is_fail'] = fail_m
    df['_is_close'] = close_m
    df['_is_gsv'] = ~fail_m & ~close_m
    df['_is_effective'] = ~fail_m  # all non-fail = close + gsv

    return df


def classify_order(status, has_return_evidence=False):
    """五分类：基于订单状态"""
    if '交易失败' in str(status):
        return CAT_OTHER  # excluded from effective
    if '交易关闭成功' in str(status):
        if '售后' in str(status) or has_return_evidence:
            return CAT_RETURN
        return CAT_UNCLEAR
    if (
        '交易成功' in str(status)
        or '待买家收货' in str(status)
        or '待卖家发货' in str(status)
    ):
        return CAT_NORMAL
    # Other GSV statuses (待平台发货, 待平台收货, etc.)
    return CAT_NORMAL


def _extract_uv_raw(xlsx_path):
    """Extract raw UV DataFrame from 商详访客 sheet."""
    for sheet in ['商详访客数据源', '商详访客数据', '商详访客']:
        try:
            raw = pd.read_excel(xlsx_path, sheet_name=sheet, engine='calamine', dtype=str)
            return raw, sheet
        except Exception:
            continue
    return pd.DataFrame(), ''


def extract_uv(xlsx_path, store_name):
    """从商详访客提取UV数据（date × goods）"""
    raw, _ = _extract_uv_raw(xlsx_path)
    if raw.empty:
        return pd.DataFrame(columns=['_date_str', '_goods', '_uv'])

    col_map = {}
    for i, c in enumerate(raw.columns):
        cs = str(c).strip()
        if '日期' in cs and 'date' not in col_map:
            col_map['date'] = i
        if ('货号' in cs or '商品货号' in cs) and 'goods' not in col_map:
            col_map['goods'] = i
        if ('商详访问' in cs or 'UV' in cs.upper()) and 'uv' not in col_map:
            col_map['uv'] = i

    if not all(k in col_map for k in ['date', 'goods', 'uv']):
        return pd.DataFrame(columns=['_date_str', '_goods', '_uv'])

    raw['_date'] = pd.to_datetime(raw.iloc[:, col_map['date']], errors='coerce')
    raw['_goods'] = raw.iloc[:, col_map['goods']].astype(str).str.strip()
    raw['_uv'] = pd.to_numeric(raw.iloc[:, col_map['uv']], errors='coerce').fillna(0)
    raw = raw.dropna(subset=['_date'])
    raw['_date_str'] = raw['_date'].dt.strftime('%Y-%m-%d')
    raw = raw[raw['_goods'] != '']
    raw = raw[raw['_goods'] != 'nan']

    uv_agg = raw.groupby(['_date_str', '_goods'])['_uv'].sum().reset_index()
    return uv_agg


def build_dynamic_cat_map(xlsx_path):
    """
    v3: Dynamically build goods→cat mapping from 商详访客 sheet.
    Reads '商品货号' → '类目名称' from the 商详访客 data.
    Falls back to keyword matching for unmapped goods.
    Returns dict: {goods_name: category}
    """
    raw, sheet_name = _extract_uv_raw(xlsx_path)
    if raw.empty:
        return {}

    # Find goods column and category column
    goods_col = None
    cat_col = None
    for i, c in enumerate(raw.columns):
        cs = str(c).strip()
        if ('货号' in cs or '商品货号' in cs) and goods_col is None:
            goods_col = i
        # Target '类目名称' specifically (not '商品类型' or other derivative columns)
        if cs == '类目名称' and cat_col is None:
            cat_col = i

    if goods_col is None or cat_col is None:
        print(f"  [cat_map] Cannot find goods/cat columns in {sheet_name}. Cols: {list(raw.columns[:15])}")
        return {}

    raw['_map_goods'] = raw.iloc[:, goods_col].astype(str).str.strip()
    raw['_map_cat'] = raw.iloc[:, cat_col].astype(str).str.strip()

    # Filter empty/invalid
    valid = raw[(raw['_map_goods'] != '') & (raw['_map_goods'] != 'nan') &
                (raw['_map_cat'] != '') & (raw['_map_cat'] != 'nan')]

    # Build unique mapping (first cat wins per goods)
    cat_map = {}
    for _, r in valid.iterrows():
        g = r['_map_goods']
        c = r['_map_cat']
        if g not in cat_map:
            cat_map[g] = c

    print(f"  [cat_map] Built {len(cat_map)} goods→cat entries from {sheet_name}")
    return cat_map


def baizhiting_classify_goods(goods_name, cat_map):
    """
    v3: Classify goods to category using ONLY authoritative mapping from 商详访客.
    No keyword guessing. Unmapped goods → '未分类类目' (preserves OID/GMV).
    """
    if not goods_name or goods_name in ('nan', ''):
        return UNCLASSIFIED_CAT

    name = str(goods_name).strip()

    # Authoritative map only (from 商详访客 '类目名称' column)
    if name in cat_map:
        mapped = cat_map[name]
        # Normalize to our 4-category system
        if '永生花' in mapped:
            return '永生花'
        if '盲盒' in mapped:
            return '盲盒'
        if any(kw in mapped for kw in ('香薰', '蜡烛', '烛台', '香氛')):
            return '香薰礼盒'
        # If the mapped category doesn't match known categories, keep as 未分类类目
        return UNCLASSIFIED_CAT

    # No mapping → 未分类类目 (完整保留OID和GMV)
    return UNCLASSIFIED_CAT


def build_m1_data(xlsx_path, store_name, brand_label=None):
    """
    构建所有M1数据，全部来自同一OID事实表。

    返回 dict with:
    - DAILY_BRAND, DAILY_GOODS, ALL_DATES, ALL_GOODS, ALL_MONTHS
    - FIVECAT (五分类 from OID table)
    - summary (KPI)
    """
    brand_label = brand_label or store_name

    # Load OID fact table
    txn = load_oid_fact_table(xlsx_path, store_name)
    uv = extract_uv(xlsx_path, store_name)

    effective = txn[txn['_is_effective']]
    effective_orig = effective  # Save before any cat filtering
    gsv = txn[txn['_is_gsv']]

    # === FIVECAT from OID table ===
    _five_map = defaultdict(lambda: {'oids': 0, 'gmv': 0.0})
    for _, row in effective.iterrows():
        cat = classify_order(row['_status'])
        _five_map[cat]['oids'] += 1
        _five_map[cat]['gmv'] += float(row['_amount'])

    fivecat = {
        'n_normal': _five_map[CAT_NORMAL]['oids'],
        'gmv_normal': round(_five_map[CAT_NORMAL]['gmv'], 2),
        'n_ret': _five_map[CAT_RETURN]['oids'],
        'gmv_ret': round(_five_map[CAT_RETURN]['gmv'], 2),
        'n_cancel': _five_map[CAT_CANCEL]['oids'],
        'gmv_cancel': round(_five_map[CAT_CANCEL]['gmv'], 2),
        'n_unclear': _five_map[CAT_UNCLEAR]['oids'],
        'gmv_unclear': round(_five_map[CAT_UNCLEAR]['gmv'], 2),
        'n_other': _five_map[CAT_OTHER]['oids'],
        'gmv_other': round(_five_map[CAT_OTHER]['gmv'], 2),
        'n_pay': len(effective),
        'gmv_pay': round(float(effective['_amount'].sum()), 2),
        'mature_return_rate': round(
            (_five_map[CAT_RETURN]['oids'] + _five_map[CAT_UNCLEAR]['oids'])
            / len(effective) * 100
            if len(effective) > 0
            else 0,
            2,
        ),
        'as_start': '2025-01-01',
        'refund_total': round(float(effective[effective['_is_close']]['_amount'].sum()), 2),
        'n_reliable_mature': len(effective),
        'n_reliable_ret': _five_map[CAT_RETURN]['oids'],
        'daily_gmv': round(
            float(effective['_amount'].sum()) / max(1, effective['_date_str'].nunique()), 2
        ),
    }

    # === DAILY_BRAND / DAILY_GOODS ===
    is_baizhiting = store_name == '柏治廷'

    if is_baizhiting:
        # v3: Dynamic cat map from 商详访客
        cat_map = build_dynamic_cat_map(xlsx_path)
        effective['_cat'] = effective['_goods'].apply(
            lambda g: baizhiting_classify_goods(g, cat_map)
        )

        # Date × cat aggregation
        db_cat = (
            effective.groupby(['_date_str', '_cat'])
            .agg(GMV=('_amount', 'sum'), orders=('_oid', 'count'))
            .reset_index()
        )
        db_cat.columns = ['date_str', 'cat', 'GMV', 'orders']

        # Merge UV at date+cat level
        uv_with_cat = effective[['_date_str', '_goods', '_cat']].drop_duplicates()
        uv_cat_merged = uv.merge(
            uv_with_cat,
            left_on=['_date_str', '_goods'],
            right_on=['_date_str', '_goods'],
            how='left',
        )
        uv_cat = (
            uv_cat_merged.groupby(['_date_str', '_cat'])['_uv'].sum().reset_index()
        )
        uv_cat.columns = ['date_str', 'cat', 'UV']

        db_cat = db_cat.merge(uv_cat, on=['date_str', 'cat'], how='outer').fillna(0)

        DAILY_BRAND = []
        for _, r in db_cat.iterrows():
            DAILY_BRAND.append({
                'date_str': r['date_str'],
                'cat': r['cat'],
                'GMV': float(round(r['GMV'], 2)),
                'UV': float(round(r['UV'], 2)),
                'orders': int(r['orders']),
            })

        ALL_CATS = sorted(set(r['cat'] for r in DAILY_BRAND))
        MARKET_CATS = ALL_CATS.copy()

        # DAILY_GOODS with cat
        dg = (
            effective.groupby(['_date_str', '_cat', '_goods'])
            .agg(GMV=('_amount', 'sum'), orders=('_oid', 'count'))
            .reset_index()
        )
        dg.columns = ['date_str', 'cat', 'goods', 'GMV', 'orders']

        uv_dg = (
            uv_cat_merged.groupby(['_date_str', '_cat', '_goods'])['_uv']
            .sum()
            .reset_index()
        )
        uv_dg.columns = ['date_str', 'cat', 'goods', 'UV']
        dg = dg.merge(uv_dg, on=['date_str', 'cat', 'goods'], how='outer').fillna(0)

        DAILY_GOODS = []
        for _, r in dg.iterrows():
            DAILY_GOODS.append({
                'date_str': r['date_str'],
                'cat': r['cat'],
                'goods': r['goods'],
                'GMV': float(round(r['GMV'], 2)),
                'UV': float(round(r['UV'], 2)),
                'orders': int(r['orders']),
            })

    else:
        # Standard: date level
        db = (
            effective.groupby('_date_str')
            .agg(GMV=('_amount', 'sum'), orders=('_oid', 'count'))
            .reset_index()
        )
        db.columns = ['date_str', 'GMV', 'orders']
        db['brand'] = brand_label

        uv_date = uv.groupby('_date_str')['_uv'].sum().reset_index()
        uv_date.columns = ['date_str', 'UV']
        db = db.merge(uv_date, on='date_str', how='outer').fillna(0)

        DAILY_BRAND = []
        for _, r in db.iterrows():
            DAILY_BRAND.append({
                'date_str': r['date_str'],
                'brand': brand_label,
                'GMV': float(round(r['GMV'], 2)),
                'UV': float(round(r['UV'], 2)),
                'orders': int(r['orders']),
            })

        ALL_CATS = []
        MARKET_CATS = []

        # DAILY_GOODS
        dg = (
            effective.groupby(['_date_str', '_goods'])
            .agg(GMV=('_amount', 'sum'), orders=('_oid', 'count'))
            .reset_index()
        )
        dg.columns = ['date_str', 'goods', 'GMV', 'orders']

        uv_dg = uv.groupby(['_date_str', '_goods'])['_uv'].sum().reset_index()
        uv_dg.columns = ['date_str', 'goods', 'UV']
        dg = dg.merge(uv_dg, on=['date_str', 'goods'], how='outer').fillna(0)

        DAILY_GOODS = []
        for _, r in dg.iterrows():
            DAILY_GOODS.append({
                'date_str': r['date_str'],
                'brand': brand_label,
                '商品货号': r['goods'],
                'GMV': float(round(r['GMV'], 2)),
                'UV': float(round(r['UV'], 2)),
                'orders': int(r['orders']),
            })

    ALL_DATES = sorted(set(r['date_str'] for r in DAILY_BRAND))
    ALL_GOODS = sorted(
        set(
            r.get('goods', r.get('商品货号', ''))
            for r in DAILY_GOODS
        )
    )
    ALL_MONTHS = sorted(set(d[:7] for d in ALL_DATES))

    # Verify identities
    eff_gmv = float(round(effective_orig['_amount'].sum(), 2))
    gsv_gmv = float(round(gsv['_amount'].sum(), 2))
    if not is_baizhiting:
        total_gmv_daily = sum(r['GMV'] for r in DAILY_BRAND)
        assert abs(total_gmv_daily - eff_gmv) < 0.02, (
            f"DAILY_BRAND GMV {total_gmv_daily} ≠ effective GMV {eff_gmv}"
        )

    return {
        'DAILY_BRAND': DAILY_BRAND,
        'DAILY_GOODS': DAILY_GOODS,
        'ALL_DATES': ALL_DATES,
        'ALL_GOODS': ALL_GOODS,
        'ALL_MONTHS': ALL_MONTHS,
        'ALL_CATS': ALL_CATS,
        'MARKET_CATS': MARKET_CATS,
        'FIVECAT': fivecat,
        'summary': {
            'total_oids': len(txn),
            'fail_oids': int(txn['_is_fail'].sum()),
            'effective_oids': int(txn['_is_effective'].sum()),
            'close_oids': int(txn['_is_close'].sum()),
            'gsv_oids': int(txn['_is_gsv'].sum()),
            'effective_gmv': eff_gmv,
            'gsv_gmv': gsv_gmv,
        },
    }


if __name__ == '__main__':
    xlsx = sys.argv[1]
    store = sys.argv[2]
    brand = sys.argv[3] if len(sys.argv) > 3 else store
    result = build_m1_data(xlsx, store, brand)
    print(json.dumps(result['summary'], ensure_ascii=False))
    print(f"DAILY_BRAND: {len(result['DAILY_BRAND'])}")
    print(
        f"FIVECAT: pay={result['FIVECAT']['n_pay']}, "
        f"gmv={result['FIVECAT']['gmv_pay']:.0f}"
    )
