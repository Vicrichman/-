#!/usr/bin/env python3
"""喜过五分类全品牌数据提取 — 输出到 fivedata_full.js"""
import pandas as pd, numpy as np, json, sys
sys.path.insert(0, '/home/Vic/dewu-reports')
from extract_fivecat import classify_store, STORE_CONFIG
from datetime import timedelta

MATURE_DAYS = 45
SRC = '/home/Vic/.hermes/tmp_data/喜过数据.xlsx'
OUT = '/home/Vic/dewu-reports/XiguoAnalysis/2026/fivedata.js'

def norm_oid(x):
    if pd.isna(x): return ''
    try: return str(int(float(x)))
    except: return str(x).strip()

# ── Load transaction orders ──
tx = pd.read_excel(SRC, sheet_name='交易订单')
tx['oid'] = tx['订单号'].apply(norm_oid)
tx['支付时间'] = pd.to_datetime(tx['买家支付时间'], errors='coerce')
tx['gmv'] = pd.to_numeric(tx['出价金额（元）'], errors='coerce').fillna(0)
tx['bid'] = tx['gmv']
tx['order_status'] = tx['订单状态'].fillna('未知')
tx['spuid'] = tx['spuID'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
tx['goods'] = tx['货号'].fillna('').astype(str)
tx['brand'] = tx['品牌'].fillna('未分类品牌').apply(lambda x: '未分类品牌' if str(x).strip() in ['', '-', '/'] else str(x).strip())

GMV_STATES = ['待平台发货','待平台收货','待卖家发货','待买家收货','交易成功','交易关闭成功']
tx['is_payment'] = tx['order_status'].isin(GMV_STATES)
pay_all = tx[tx['is_payment']].copy()

# ── Load after-sale ──
as_all = pd.read_excel(SRC, sheet_name='售后订单')
as_all['oid'] = as_all['订单号'].apply(norm_oid)
as_all['as_type'] = as_all['售后类型'].fillna('')
as_all['as_status'] = as_all['售后状态'].fillna('')
as_all['as_refund'] = pd.to_numeric(as_all['退款金额（元）'], errors='coerce').fillna(0)
as_all['as_time'] = pd.to_datetime(as_all['买家申请售后时间'], errors='coerce')

as_valid = as_all[(as_all['as_status']=='售后成功') & (~as_all['as_type'].isin(['取消','补寄']))].copy()
as_cancel = as_all[(as_all['as_status']=='售后成功') & (as_all['as_type']=='取消')].copy()

as_grp = as_valid.groupby('oid').agg(total_refund=('as_refund','sum'), as_types=('as_type',lambda x: '|'.join(sorted(set(x))))).reset_index()
cancel_oids = set(as_cancel['oid'].unique())

as_start = as_all[as_all['as_time'].notna() & (as_all['as_type']!='')]['as_time'].min()

# ── Load 商详访客 for UV ──
uv = pd.read_excel(SRC, sheet_name='商详访客数据')
uv['SPUID'] = uv['SPUID'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
uv['日期'] = pd.to_datetime(uv['日期'], errors='coerce')

# ── Merge & Classify ──
pay_all = pay_all.merge(as_grp, on='oid', how='left')
pay_all['total_refund'] = pay_all['total_refund'].fillna(0)
pay_all['as_types'] = pay_all['as_types'].fillna('')
pay_all['is_closed'] = pay_all['order_status'] == '交易关闭成功'
pay_all['cat_ret_as'] = pay_all['is_closed'] & pay_all['as_types'].str.contains('退货', na=False)
pay_all['cat_ret_50'] = (pay_all['order_status']=='交易成功') & (pay_all['total_refund']>0) & (pay_all['bid']>0) & (pay_all['total_refund']/pay_all['bid']>0.5)
pay_all['cat_cancel'] = pay_all['is_closed'] & pay_all['oid'].isin(cancel_oids)
pay_all['cat_unclear'] = pay_all['is_closed'] & (~pay_all['cat_ret_as']) & (~pay_all['cat_cancel'])
pay_all['cat_ret'] = pay_all['cat_ret_as'] | pay_all['cat_ret_50']
pay_all['cat_normal'] = (~pay_all['cat_ret']) & (~pay_all['cat_cancel']) & (~pay_all['cat_unclear'])

# Pre-compute GMV columns for aggregation
pay_all['gmv_ret_d'] = pay_all['gmv'].where(pay_all['cat_ret'], 0)
pay_all['gmv_cancel_d'] = pay_all['gmv'].where(pay_all['cat_cancel'], 0)
pay_all['gmv_unclear_d'] = pay_all['gmv'].where(pay_all['cat_unclear'], 0)
pay_all['gmv_normal_d'] = pay_all['gmv'].where(pay_all['cat_normal'], 0)

# ── Mature rate ──
latest_pay = pay_all['支付时间'].max()
mature_cutoff = latest_pay - timedelta(days=MATURE_DAYS)
reliable = pay_all[(pay_all['支付时间'] >= as_start) & (pay_all['支付时间'] <= mature_cutoff)]
mature_rate = reliable['cat_ret'].sum() / len(reliable) * 100 if len(reliable) > 0 else 0

# ── Aggregate by brand ──
def agg_brand(df, prefix=''):
    result = {}
    for brand, grp in df.groupby('brand'):
        result[brand] = {
            f'n_pay': len(grp),
            f'gmv_pay': grp['gmv'].sum(),
            f'gmv_ret': grp['gmv_ret_d'].sum(),
            f'gmv_cancel': grp['gmv_cancel_d'].sum(),
            f'gmv_unclear': grp['gmv_unclear_d'].sum(),
            f'gmv_normal': grp['gmv_normal_d'].sum(),
            f'n_ret': int(grp['cat_ret'].sum()),
            f'n_cancel': int(grp['cat_cancel'].sum()),
            f'n_unclear': int(grp['cat_unclear'].sum()),
        }
    return result

brand_data = agg_brand(pay_all)

# ── Aggregate by spuid ──
def agg_spuid(df):
    result = {}
    for spuid, grp in df.groupby('spuid'):
        if not spuid: continue
        brand = grp['brand'].mode().iloc[0] if len(grp) > 0 else '未知'
        goods = grp['goods'].mode().iloc[0] if len(grp) > 0 else ''
        result[spuid] = {
            'spuid': spuid,
            'brand': brand,
            'goods': goods,
            'n_pay': len(grp),
            'gmv_pay': grp['gmv'].sum(),
            'gmv_ret': grp['gmv_ret_d'].sum(),
            'gmv_cancel': grp['gmv_cancel_d'].sum(),
            'gmv_unclear': grp['gmv_unclear_d'].sum(),
            'gmv_normal': grp['gmv_normal_d'].sum(),
        }
    return result

spuid_data = agg_spuid(pay_all)

# ── Aggregate by date ──
pay_all['pay_date'] = pay_all['支付时间'].dt.date
daily = pay_all.groupby(['pay_date','brand']).agg(
    gmv_pay=('gmv','sum'), gmv_ret=('gmv_ret_d','sum'),
    gmv_cancel=('gmv_cancel_d','sum'), gmv_unclear=('gmv_unclear_d','sum'),
    gmv_normal=('gmv_normal_d','sum'), n_pay=('oid','count'),
    n_ret=('cat_ret','sum'), n_cancel=('cat_cancel','sum'),
).reset_index()
daily['pay_date'] = daily['pay_date'].astype(str)

# ── Summary ──
n_pay = len(pay_all)
summary = {
    'store': '喜过',
    'cutoff_date': str(latest_pay.date()),
    'as_start': str(as_start.date()),
    'mature_cutoff': str(mature_cutoff.date()),
    'n_pay': n_pay,
    'gmv_pay': pay_all['gmv'].sum(),
    'gmv_ret': pay_all['gmv_ret_d'].sum(),
    'gmv_cancel': pay_all['gmv_cancel_d'].sum(),
    'gmv_unclear': pay_all['gmv_unclear_d'].sum(),
    'gmv_normal': pay_all['gmv_normal_d'].sum(),
    'n_ret': int(pay_all['cat_ret'].sum()),
    'n_cancel': int(pay_all['cat_cancel'].sum()),
    'n_unclear': int(pay_all['cat_unclear'].sum()),
    'mature_return_rate': round(mature_rate, 2),
    'n_reliable_mature': len(reliable),
    'n_reliable_ret': int(reliable['cat_ret'].sum()),
    'all_brands': sorted(pay_all['brand'].unique().tolist()),
    'brand_data': brand_data,
    'spuid_data': list(spuid_data.values()),
    'daily_brand': daily.to_dict('records'),
}

# ── Write JS ──
js = 'const FIVECAT = ' + json.dumps(summary, ensure_ascii=False, default=str) + ';'
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(js)

print(f'✅ Written {OUT} ({len(js)} chars)')
print(f'   pay={n_pay}, GMV=¥{summary["gmv_pay"]:,.0f}, mature_rate={mature_rate:.2f}%')
print(f'   Brands: {summary["all_brands"]}')
print(f'   SPUIDs: {len(spuid_data)}')
print(f'   Daily-brand records: {len(daily)}')

# Verify identity
total = summary['gmv_ret']+summary['gmv_cancel']+summary['gmv_unclear']+summary['gmv_normal']
ok = abs(summary['gmv_pay'] - total) < 0.01
print(f'   GMV恒等式: {"✅" if ok else "❌"}')

# Brand identity check
brand_gmv = sum(b['gmv_pay'] for b in brand_data.values())
print(f'   品牌GMV合计=¥{brand_gmv:,.0f} vs 店铺GMV=¥{summary["gmv_pay"]:,.0f} → {"✅" if abs(brand_gmv-summary["gmv_pay"])<0.01 else "❌"}')
