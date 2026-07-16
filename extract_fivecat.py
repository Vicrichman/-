#!/usr/bin/env python3
"""
六店统一五分类提取模块 v2
产出每个店铺的五分类GMV数据、可靠队列成熟退货率、KPI卡片JSON
供各店 extract_*.py 调用或独立运行
"""
import pandas as pd, numpy as np, json, re
from datetime import datetime, timedelta
from collections import defaultdict

MATURE_DAYS = 45

# ── 店铺配置 ──
STORE_CONFIG = {
    '醒狮': {
        'src': '/home/Vic/.hermes/tmp_data/醒狮数据.xlsx',
        'tx_sheet': '交易订单', 'as_sheet': '售后订单',
        'gmv_col': '出价金额（元）',
        'gmv_states': ['待平台发货','待平台收货','待卖家发货','待买家收货','交易成功','交易关闭成功'],
    },
    '美兰': {
        'src': '/home/Vic/.hermes/tmp_data/美兰数据.xlsx',
        'tx_sheet': '交易订单', 'as_sheet': '售后订单',
        'gmv_col': '出价金额（元）',
        'gmv_states': ['待平台发货','待平台收货','待卖家发货','待买家收货','交易成功','交易关闭成功'],
    },
    '喜过': {
        'src': '/home/Vic/.hermes/tmp_data/喜过数据.xlsx',
        'tx_sheet': '交易订单', 'as_sheet': '售后订单',
        'gmv_col': '出价金额（元）',
        'gmv_states': ['待平台发货','待平台收货','待卖家发货','待买家收货','交易成功','交易关闭成功'],
    },
    '蓄势': {
        'src': '/home/Vic/.hermes/tmp_data/蓄势数据.xlsx',
        'tx_sheet': '交易订单', 'as_sheet': '售后订单',
        'gmv_col': '出价金额（元）',
        'gmv_states': ['待平台发货','待平台收货','待卖家发货','待买家收货','交易成功','交易关闭成功'],
    },
    '梦特娇': {
        'src': '/home/Vic/.hermes/tmp_data/梦特娇数据.xlsx',
        'tx_sheet': '交易订单', 'as_sheet': '售后订单',
        'gmv_col': '出价金额（元）',
        'gmv_states': ['已发货','交易成功','交易关闭成功'],
        'no_as_fields': True,
    },
    '柏治廷': {
        'src': '/home/Vic/.hermes/tmp_data/baizhiting_latest.xlsx',
        'tx_sheet': '交易订单', 'as_sheet': '售后订单',
        'gmv_col': '出价金额（）',
        'gmv_states': ['待平台发货','待平台收货','待卖家发货','待买家收货','交易成功','交易关闭成功','待物流揽收'],
    },
}

def norm_oid(x):
    if pd.isna(x): return ''
    try: return str(int(float(x)))
    except: return str(x).strip()

def classify_store(store_name, cfg):
    """Run full five-category classification for one store. Returns dict."""
    tx = pd.read_excel(cfg['src'], sheet_name=cfg['tx_sheet'])
    as_all = pd.read_excel(cfg['src'], sheet_name=cfg['as_sheet'])
    
    tx['oid'] = tx['订单号'].apply(norm_oid)
    tx['支付时间'] = pd.to_datetime(tx['买家支付时间'], errors='coerce')
    tx['gmv'] = pd.to_numeric(tx[cfg['gmv_col']], errors='coerce').fillna(0)
    tx['bid'] = tx['gmv']
    tx['order_status'] = tx['订单状态'].fillna('未知')
    tx['is_payment'] = tx['order_status'].isin(cfg['gmv_states'])
    
    pay_all = tx[tx['is_payment']].copy()
    latest_pay = pay_all['支付时间'].max()
    
    # After-sale processing
    no_as = cfg.get('no_as_fields', False)
    has_as_fields = ('售后类型' in as_all.columns and '售后状态' in as_all.columns and 
                     any('退款金额' in str(c) for c in as_all.columns))
    
    if has_as_fields and not no_as:
        as_all['oid'] = as_all['订单号'].apply(norm_oid)
        as_all['as_type'] = as_all['售后类型'].fillna('')
        as_all['as_status'] = as_all['售后状态'].fillna('')
        as_refund_col = next(c for c in as_all.columns if '退款金额' in str(c))
        as_all['as_refund'] = pd.to_numeric(as_all[as_refund_col], errors='coerce').fillna(0)
        as_all['as_time'] = pd.to_datetime(as_all['买家申请售后时间'], errors='coerce')
        
        as_valid = as_all[(as_all['as_status']=='售后成功') & (~as_all['as_type'].isin(['取消','补寄']))].copy()
        as_cancel = as_all[(as_all['as_status']=='售后成功') & (as_all['as_type']=='取消')].copy()
        
        as_grp = as_valid.groupby('oid').agg(
            total_refund=('as_refund','sum'), as_types=('as_type',lambda x: '|'.join(sorted(set(x))))
        ).reset_index()
        cancel_oids = set(as_cancel['oid'].unique())
        
        as_start = as_all[as_all['as_time'].notna() & (as_all['as_type']!='')]['as_time'].min()
        as_available = True
    else:
        as_grp = pd.DataFrame(columns=['oid','total_refund','as_types'])
        cancel_oids = set()
        as_start = pd.NaT
        as_available = False
    
    pay_all = pay_all.merge(as_grp, on='oid', how='left')
    pay_all['total_refund'] = pay_all['total_refund'].fillna(0)
    pay_all['as_types'] = pay_all['as_types'].fillna('')
    
    # ── Classification ──
    if as_available:
        pay_all['is_closed'] = pay_all['order_status'] == '交易关闭成功'
        pay_all['cat_ret_as'] = pay_all['is_closed'] & pay_all['as_types'].str.contains('退货', na=False)
        pay_all['cat_ret_50'] = (pay_all['order_status']=='交易成功') & (pay_all['total_refund']>0) & (pay_all['bid']>0) & (pay_all['total_refund']/pay_all['bid']>0.5)
        pay_all['cat_cancel'] = pay_all['is_closed'] & pay_all['oid'].isin(cancel_oids)
        pay_all['cat_unclear'] = pay_all['is_closed'] & (~pay_all['cat_ret_as']) & (~pay_all['cat_cancel'])
        pay_all['cat_ret'] = pay_all['cat_ret_as'] | pay_all['cat_ret_50']
    else:
        pay_all['cat_ret'] = False
        pay_all['cat_ret_as'] = False
        pay_all['cat_ret_50'] = False
        pay_all['cat_cancel'] = False
        pay_all['cat_unclear'] = pay_all['order_status'] == '交易关闭成功'
    
    pay_all['cat_normal'] = (~pay_all['cat_ret']) & (~pay_all['cat_cancel']) & (~pay_all['cat_unclear'])
    
    # Mature data cutoff = MIN(transaction cutoff, after-sale cutoff)
    if as_available and pd.notna(as_start):
        as_max_dt = as_all['as_time'].max()
        data_cutoff = min(latest_pay, as_max_dt)
    else:
        data_cutoff = latest_pay
    mature_cutoff = data_cutoff - timedelta(days=MATURE_DAYS)
    
    # ── Metrics ──
    n = len(pay_all)
    gmv_pay = pay_all['gmv'].sum()
    
    def sum_gmv(df, col): return df[df[col]]['gmv'].sum()
    
    result = {
        'store': store_name,
        'cutoff_date': str(latest_pay.date()),
        'mature_cutoff': str(mature_cutoff.date()),
        'n_pay': int(n),
        'gmv_pay': float(gmv_pay),
        'n_ret': int(pay_all['cat_ret'].sum()),
        'gmv_ret': float(sum_gmv(pay_all, 'cat_ret')),
        'n_cancel': int(pay_all['cat_cancel'].sum()),
        'gmv_cancel': float(sum_gmv(pay_all, 'cat_cancel')),
        'n_unclear': int(pay_all['cat_unclear'].sum()),
        'gmv_unclear': float(sum_gmv(pay_all, 'cat_unclear')),
        'n_normal': int(pay_all['cat_normal'].sum()),
        'gmv_normal': float(sum_gmv(pay_all, 'cat_normal')),
        'gmv_other': float(gmv_pay - sum_gmv(pay_all, 'cat_ret') - sum_gmv(pay_all, 'cat_cancel') - sum_gmv(pay_all, 'cat_unclear') - sum_gmv(pay_all, 'cat_normal')),
        'refund_total': float(pay_all['total_refund'].sum()),
        'as_available': as_available,
    }
    
    # ── Mature rate with reliable cohort ──
    if as_available and pd.notna(as_start):
        as_start_dt = as_start
        reliable = pay_all[(pay_all['支付时间'] >= as_start_dt) & (pay_all['支付时间'] <= mature_cutoff)]
        n_rel = len(reliable)
        n_rel_ret = int(reliable['cat_ret'].sum())
        mature_rate = n_rel_ret / n_rel * 100 if n_rel > 0 else 0
        
        result['as_start'] = str(as_start_dt.date())
        result['n_reliable_mature'] = n_rel
        result['n_reliable_ret'] = n_rel_ret
        result['mature_return_rate'] = round(mature_rate, 2)
        result['reliable_cancel_rate'] = round(reliable['cat_cancel'].sum() / n_rel * 100, 2) if n_rel else 0
        result['reliable_unclear_rate'] = round(reliable['cat_unclear'].sum() / n_rel * 100, 2) if n_rel else 0
        result['n_excluded_pre_as'] = int((pay_all['支付时间'] < as_start_dt).sum())
    else:
        result['as_start'] = None
        result['n_reliable_mature'] = 0
        result['n_reliable_ret'] = 0
        result['mature_return_rate'] = None
        result['reliable_cancel_rate'] = None
        result['reliable_unclear_rate'] = None
        result['n_excluded_pre_as'] = 0
    
    # ── All-history mature rate (for comparison) ──
    all_mature = pay_all[pay_all['支付时间'] <= mature_cutoff]
    n_am = len(all_mature)
    if n_am > 0:
        result['all_hist_mature_rate'] = round(all_mature['cat_ret'].sum() / n_am * 100, 2)
    else:
        result['all_hist_mature_rate'] = None
    
    # ── Daily aggregation ──
    pay_all['pay_date'] = pay_all['支付时间'].dt.date
    pay_all['gmv_ret_d'] = pay_all['gmv'].where(pay_all['cat_ret'], 0)
    pay_all['gmv_cancel_d'] = pay_all['gmv'].where(pay_all['cat_cancel'], 0)
    pay_all['gmv_unclear_d'] = pay_all['gmv'].where(pay_all['cat_unclear'], 0)
    pay_all['gmv_normal_d'] = pay_all['gmv'].where(pay_all['cat_normal'], 0)
    daily = pay_all.groupby('pay_date').agg(
        gmv_pay=('gmv','sum'),
        gmv_ret=('gmv_ret_d','sum'),
        gmv_cancel=('gmv_cancel_d','sum'),
        gmv_unclear=('gmv_unclear_d','sum'),
        gmv_normal=('gmv_normal_d','sum'),
        n_pay=('oid','count'),
        n_ret=('cat_ret','sum'),
        n_cancel=('cat_cancel','sum'),
    ).reset_index()
    daily['pay_date'] = daily['pay_date'].astype(str)
    daily.columns = ['date','gmv_pay','gmv_ret','gmv_cancel','gmv_unclear','gmv_normal','n_pay','n_ret','n_cancel']
    result['daily_gmv'] = daily.to_dict('records')
    
    return result

if __name__ == '__main__':
    import sys
    stores = sys.argv[1:] if len(sys.argv) > 1 else list(STORE_CONFIG.keys())
    
    all_results = {}
    for store in stores:
        cfg = STORE_CONFIG[store]
        print(f"Processing {store}...", file=sys.stderr)
        r = classify_store(store, cfg)
        all_results[store] = r
        print(f"  {store}: pay={r['n_pay']}, ret={r['n_ret']}, cancel={r['n_cancel']}, unclear={r['n_unclear']}, mature_rate={r['mature_return_rate']}", file=sys.stderr)
    
    print(json.dumps(all_results, ensure_ascii=False, default=str, indent=2))
