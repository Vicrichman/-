#!/usr/bin/env python3
"""
喜过分析 - pandas 数据补充脚本
从 Excel 补充 extract_data.py 缺失的三个字段到 data.js:
  - uv_daily_by_brand  (模块二依赖)
  - uv_spu_data        (模块三/四依赖)
  - huohao_daily       (模块五依赖)

Usage: python3 supplement_data.py
Requirements: pip install pandas python-calamine
"""
import pandas as pd
import json
import re
import os
from collections import defaultdict

EXCEL = "/mnt/e/Obsidian本地仓库/09-数据源/喜过数据源收集表.xlsx"
DATA_JS = "/home/Vic/dewu-reports/XiguoAnalysis/2026/data.js"
BRANDS_LIST = ["CASIO/卡西欧", "COACH/蔻驰"]
BRAND_SHORT = {"CASIO/卡西欧": "卡西欧", "COACH/蔻驰": "蔻驰"}
START_DATE = "2025-05-01"
END_DATE = "2026-06-28"

def in_range(d):
    return d and START_DATE <= str(d)[:10] <= END_DATE

def clean_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    m = re.match(r'[\d.]+', s)
    return float(m.group()) if m else 0

# 1. Read existing data.js
print("1. Reading existing data.js...")
with open(DATA_JS, 'r', encoding='utf-8') as f:
    content = f.read()
json_str = content[content.index('{'):content.rindex('};')+1]
D = json.loads(json_str)

# 2. Build SPU->brand mapping
print("2. Building SPU→brand mapping...")
df_huopan = pd.read_excel(EXCEL, sheet_name="货盘表", engine="calamine")
spu_brand, spu_name = {}, {}
for _, row in df_huopan.iterrows():
    spuid, name, brand = row.iloc[0], row.iloc[1], row.iloc[4]
    try:
        if brand in BRANDS_LIST:
            spu_brand[int(spuid)] = BRAND_SHORT[brand]
            spu_name[int(spuid)] = str(name)
    except (ValueError, TypeError):
        pass
print(f"   {len(spu_brand)} SPU->brand mappings")

# 3. huohao_daily from 交易订单
print("3. Building huohao_daily from 交易订单...")
df_orders = pd.read_excel(EXCEL, sheet_name="交易订单", engine="calamine")

col_order_status = col_pay_time = col_order_time = None
for i, c in enumerate(df_orders.columns):
    cn = str(c).strip()
    if "订单状态" in cn: col_order_status = i
    elif "支付时间" in cn or "付款时间" in cn: col_pay_time = i
    elif "订单创建时间" in cn or "下单时间" in cn: col_order_time = i

VALID_STATUSES = ["交易成功", "待发货", "待收货", "待买家收货", "待平台发货", "待卖家发货", "待平台收货"]

huohao_daily = defaultdict(list)
processed = 0
for _, row in df_orders.iterrows():
    spuid, brand, status = row.iloc[2], row.iloc[7], row.iloc[col_order_status] if col_order_status is not None else None
    date_str = None
    if col_pay_time is not None:
        dt = row.iloc[col_pay_time]
        if pd.notna(dt): date_str = str(dt)[:10]
    if not date_str and col_order_time is not None:
        dt = row.iloc[col_order_time]
        if pd.notna(dt): date_str = str(dt)[:10]
    if not (pd.notna(spuid) and pd.notna(brand) and brand in BRANDS_LIST and 
            str(status).strip() in VALID_STATUSES and date_str and in_range(date_str)):
        continue
    b = BRAND_SHORT[brand]
    amount = float(row.iloc[10]) if pd.notna(row.iloc[10]) else 0
    qty = int(row.iloc[9]) if pd.notna(row.iloc[9]) else 1
    gmv = amount * qty
    hh = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else "—"
    huohao_daily[b].append({"date": date_str[:10], "huohao": hh, "gmv": round(gmv, 2), "orders": 1})
    processed += 1
print(f"   Processed {processed} orders ({', '.join(f'{b}:{len(huohao_daily[b])}' for b in BRAND_SHORT.values())})")

# 4. uv_daily_by_brand + uv_spu_data from 商详访客数据
print("4. Building UV data from 商详访客数据...")
df_uv = pd.read_excel(EXCEL, sheet_name="商详访客数据", engine="calamine")

uv_daily_raw = defaultdict(list)
uv_spu_raw = defaultdict(lambda: defaultdict(lambda: {"daily": [], "huohao": ""}))
uv_processed = 0
for _, row in df_uv.iterrows():
    date_str = str(row.iloc[1])[:10] if pd.notna(row.iloc[1]) else None
    spuid, brand = row.iloc[2], row.iloc[5]
    huohao = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
    if not (pd.notna(spuid) and brand in BRANDS_LIST and date_str and in_range(date_str)):
        continue
    b = BRAND_SHORT[brand]
    uv_val = int(clean_num(row.iloc[11]))
    pay_amt = clean_num(row.iloc[9])
    pay_ord = int(clean_num(row.iloc[10]))
    uv_daily_raw[b].append({"date": date_str[:10], "uv": uv_val, "gmv": round(pay_amt, 2), "orders": pay_ord})
    spu_key = spu_name.get(int(spuid), str(spuid))
    uv_spu_raw[b][spu_key]["daily"].append({"date": date_str[:10], "uv": uv_val, "gmv": round(pay_amt, 2), "orders": pay_ord})
    if not uv_spu_raw[b][spu_key]["huohao"]:
        uv_spu_raw[b][spu_key]["huohao"] = huohao
    uv_processed += 1
print(f"   UV processed {uv_processed} records")

# Aggregate uv_daily_by_brand
uv_daily_out = {}
for brand in BRAND_SHORT.values():
    dm = defaultdict(lambda: {"uv": 0, "gmv": 0.0, "orders": 0})
    for d in uv_daily_raw[brand]:
        dm[d["date"]]["uv"] += d["uv"]
        dm[d["date"]]["gmv"] += d["gmv"]
        dm[d["date"]]["orders"] += d["orders"]
    uv_daily_out[brand] = [{"date": k, "uv": v["uv"], "gmv": round(v["gmv"], 2), "orders": v["orders"]}
                           for k, v in sorted(dm.items())]

# Build uv_spu_data
uv_spu_out = {}
for brand in BRAND_SHORT.values():
    spus = []
    for spu_key, data in uv_spu_raw[brand].items():
        total_uv = sum(d["uv"] for d in data["daily"])
        total_gmv = sum(d["gmv"] for d in data["daily"])
        total_orders = sum(d["orders"] for d in data["daily"])
        spus.append({"spu": spu_key, "huohao": data["huohao"],
                      "total_uv": total_uv, "total_gmv": round(total_gmv, 2),
                      "total_orders": total_orders,
                      "daily": sorted(data["daily"], key=lambda x: x["date"])})
    uv_spu_out[brand] = spus

# 5. Merge and write
print("5. Merging and writing data.js...")
D["huohao_daily"] = {b: huohao_daily[b] for b in BRAND_SHORT.values()}
D["uv_daily_by_brand"] = uv_daily_out
D["uv_spu_data"] = uv_spu_out

with open(DATA_JS, 'w', encoding='utf-8') as f:
    f.write("var D=")
    json.dump(D, f, ensure_ascii=False, separators=(",", ":"))
    f.write(";\n")

size_mb = os.path.getsize(DATA_JS) / (1024*1024)
print(f"   Written: {size_mb:.1f}MB, keys: {list(D.keys())}")
print("✅ Supplement complete!")
