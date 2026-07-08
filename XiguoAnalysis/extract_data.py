"""
喜过分析模板 - 数据提取脚本
从 喜过数据源收集表.xlsx 提取卡西欧 vs 蔻驰的全量运营数据
输出: data.js (内联到看板HTML)

核心修复:
1. 大盘指数提取覆盖全时间范围（不只是May）
2. 社区投放/得物推按月汇总
3. UV月度数据
4. SPU级别的UV、支付订单、支付金额详情
"""
import openpyxl
from datetime import datetime, timedelta
from collections import defaultdict
import json
import sys

EXCEL_PATH = "/mnt/e/Obsidian本地仓库/09-数据源/喜过数据源收集表.xlsx"
START_DATE = "2025-05-01"
END_DATE = "2026-06-28"
BRANDS = ["CASIO/卡西欧", "COACH/蔻驰"]
BRAND_SHORT = {"CASIO/卡西欧": "卡西欧", "COACH/蔻驰": "蔻驰"}

def excel_to_date(serial):
    """Convert Excel date serial or string to YYYY-MM-DD string"""
    if isinstance(serial, datetime):
        return serial.strftime("%Y-%m-%d")
    if isinstance(serial, str) and serial.strip():
        # Try to parse string dates like "2026-05-07 14:21:57"
        s = serial.strip()
        for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d", "%Y/%m/%d %H:%M:%S"]:
            try:
                return datetime.strptime(s[:19] if len(s)>=19 else s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        # If it starts with YYYY-MM-DD, just take the first 10 chars
        if len(s) >= 10 and s[4] == '-' and s[7] == '-':
            return s[:10]
    if isinstance(serial, (int, float)) and serial > 40000:
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=int(serial))).strftime("%Y-%m-%d")
    return None

def in_range(date_str):
    """Check if date string is within range. date_str is YYYY-MM-DD format."""
    if not date_str:
        return False
    return START_DATE <= date_str <= END_DATE

def date_to_month(date_str):
    return date_str[:7]

print("Loading workbook...", file=sys.stderr)
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

# ============================================================
# 1. 大盘指数 (日韩表 + 单肩包)
# ============================================================
print("1. Extracting 大盘指数...", file=sys.stderr)
ws = wb["大盘指数"]
market_data = []          # 日韩表大盘
market_monthly = defaultdict(list)
market_data_bag = []      # 单肩包大盘
market_monthly_bag = defaultdict(list)
for r in range(2, ws.max_row + 1):
    d = excel_to_date(ws.cell(r, 2).value)
    v = ws.cell(r, 3).value  # 日韩表大盘成交GMV指数
    v_bag = ws.cell(r, 4).value  # 单肩包大盘成交GMV指数
    if d and in_range(d):
        if v:
            market_data.append({"date": d, "value": round(float(v), 2)})
            market_monthly[date_to_month(d)].append(float(v))
        if v_bag:
            market_data_bag.append({"date": d, "value": round(float(v_bag), 2)})
            market_monthly_bag[date_to_month(d)].append(float(v_bag))

market_monthly_avg = {m: round(sum(vals)/len(vals), 2) for m, vals in sorted(market_monthly.items())}
market_monthly_bag_avg = {m: round(sum(vals)/len(vals), 2) for m, vals in sorted(market_monthly_bag.items())}
print(f"  Market(日韩表): {len(market_data)} daily records, {len(market_monthly_avg)} months", file=sys.stderr)
print(f"  Market(单肩包): {len(market_data_bag)} daily records, {len(market_monthly_bag_avg)} months", file=sys.stderr)
for m in sorted(set(list(market_monthly_avg.keys()) + list(market_monthly_bag_avg.keys()))):
    a1 = market_monthly_avg.get(m, "—")
    a2 = market_monthly_bag_avg.get(m, "—")
    print(f"    {m}: 日韩表={a1}, 单肩包={a2}", file=sys.stderr)

# ============================================================
# 2. 货盘表 - build SPU -> brand mapping
# ============================================================
print("2. Building SPU->brand mapping...", file=sys.stderr)
ws = wb["货盘表"]
spu_brand = {}
spu_name = {}
for r in range(2, ws.max_row + 1):
    spuid = ws.cell(r, 1).value
    name = ws.cell(r, 2).value
    brand = ws.cell(r, 5).value
    if spuid and brand in BRANDS:
        try:
            spu_brand[int(spuid)] = BRAND_SHORT[brand]
            spu_name[int(spuid)] = str(name)
        except (ValueError, TypeError):
            pass

print(f"  SPU mapping: {len(spu_brand)} SPUs", file=sys.stderr)

# ============================================================
# 3. 交易订单 - 按月/按SPU汇总
# ============================================================
print("3. Extracting 交易订单...", file=sys.stderr)
ws = wb["交易订单"]
# Headers: 订单号, 订单类型, spuID, skuID, 商品名称, 货号, SKU货号, 品牌, 规格, 数量, 出价金额(元), ...
# Find column indices
header = {}
for c in range(1, ws.max_column + 1):
    h = ws.cell(1, c).value
    if h: header[str(h).strip()] = c

col_order_status = None
col_pay_time = None
col_order_time = None
for c in range(1, ws.max_column + 1):
    h = str(ws.cell(1, c).value or "").strip()
    if "订单状态" in h:
        col_order_status = c
    elif "支付时间" in h or "付款时间" in h:
        col_pay_time = c
    elif "订单创建时间" in h or "下单时间" in h:
        col_order_time = c

print(f"  order_status col={col_order_status}, pay_time={col_pay_time}, order_time={col_order_time}", file=sys.stderr)

spu_id_col = 3  # spuID
sku_id_col = 4  
name_col = 5
brand_col = 8
amount_col = 11
qty_col = 10
huohao_col = 6  # 货号

VALID_STATUSES = ["交易成功", "待发货", "待收货", "待买家收货", "待平台发货", "待卖家发货", "待平台收货"]

orders_monthly = defaultdict(lambda: defaultdict(lambda: {"gmv": 0.0, "orders": 0}))
orders_by_spu = defaultdict(lambda: defaultdict(lambda: {"gmv": 0.0, "orders": 0}))
orders_by_huohao = defaultdict(lambda: defaultdict(lambda: {"gmv": 0.0, "orders": 0}))  # 按货号汇总

processed = 0
for r in range(2, ws.max_row + 1):
    spuid = ws.cell(r, spu_id_col).value
    amount = ws.cell(r, amount_col).value
    brand = ws.cell(r, brand_col).value
    status = ws.cell(r, col_order_status).value if col_order_status else None
    
    # Get date
    date_str = None
    if col_pay_time:
        dt = ws.cell(r, col_pay_time).value
        date_str = excel_to_date(dt)
    if not date_str and col_order_time:
        dt = ws.cell(r, col_order_time).value
        date_str = excel_to_date(dt)
    
    if not (spuid and amount and brand in BRANDS and status in VALID_STATUSES and date_str and in_range(date_str)):
        continue
    
    b = BRAND_SHORT[brand]
    month = date_to_month(date_str)
    gmv = float(amount) * (ws.cell(r, qty_col).value or 1)
    
    orders_monthly[b][month]["gmv"] += gmv
    orders_monthly[b][month]["orders"] += 1
    
    spu_key = spu_name.get(int(spuid), str(spuid))
    orders_by_spu[b][spu_key]["gmv"] += gmv
    orders_by_spu[b][spu_key]["orders"] += 1
    orders_by_spu[b][spu_key]["spuid"] = int(spuid)
    
    # 按货号汇总
    hh = str(ws.cell(r, huohao_col).value).strip() if ws.cell(r, huohao_col).value else spu_key
    orders_by_huohao[b][hh]["gmv"] += gmv
    orders_by_huohao[b][hh]["orders"] += 1
    
    processed += 1

print(f"  Processed {processed} valid orders", file=sys.stderr)

# Convert to sorted lists
orders_monthly_out = {}
for brand in BRAND_SHORT.values():
    orders_monthly_out[brand] = {}
    for m in sorted(orders_monthly[brand].keys()):
        orders_monthly_out[brand][m] = {
            "gmv": round(orders_monthly[brand][m]["gmv"], 2),
            "orders": orders_monthly[brand][m]["orders"]
        }

orders_by_spu_out = {}
for brand in BRAND_SHORT.values():
    spus = []
    for spu_name_key, data in sorted(orders_by_spu[brand].items(), key=lambda x: -x[1]["gmv"]):
        spus.append({
            "spu": spu_name_key,
            "spuid": data["spuid"],
            "gmv": round(data["gmv"], 2),
            "orders": data["orders"]
        })
    orders_by_spu_out[brand] = spus

# ============================================================
# 4. 商详访客数据 - UV按月/按SPU
# ============================================================
print("4. Extracting 商详访客数据...", file=sys.stderr)
ws = wb["商详访客数据"]
# Columns: 月份, 日期, SPUID, 商品名称, 货号, 品牌名称, 类目名称, 商品类型, 出价渠道, 支付订单金额, 支付订单量, ...
# Find the key columns
uv_header = {}
for c in range(1, min(50, ws.max_column + 1)):
    h = ws.cell(1, c).value
    if h: uv_header[str(h).strip()] = c

# Check what columns exist
print(f"  商详 headers (first 20 cols): {[ws.cell(1,c).value for c in range(1,21)]}", file=sys.stderr)

# Find indices: SPUID=col3, 品牌=col6, 商详访客人数 column
spuid_col_uv = 3
brand_col_uv = 6
# Look for 商详访客人数 column
uv_col = None
pay_amount_col = None
pay_orders_col = None
for c in range(1, min(50, ws.max_column + 1)):
    h = str(ws.cell(1, c).value or "").strip()
    # Specifically look for 商详访问指数（UV）or similar
    if "商详访问" in h and ("uv" in h.lower() or "指数" in h):
        if uv_col is None:
            uv_col = c
    elif "支付订单金额" in h:
        pay_amount_col = c
    elif "支付订单量" in h:
        pay_orders_col = c

print(f"  UV col={uv_col}, pay_amount={pay_amount_col}, pay_orders={pay_orders_col}", file=sys.stderr)

uv_monthly = defaultdict(lambda: defaultdict(lambda: {"uv": 0, "gmv": 0.0, "orders": 0}))
uv_by_spu = defaultdict(lambda: defaultdict(lambda: {"daily": [], "monthly": defaultdict(lambda: {"uv": 0, "gmv": 0.0, "orders": 0})}))

uv_processed = 0
for r in range(2, ws.max_row + 1):
    spuid = ws.cell(r, spuid_col_uv).value
    brand = ws.cell(r, brand_col_uv).value
    
    # Get date - column 2 is 日期
    date_cell = ws.cell(r, 2).value
    date_str = excel_to_date(date_cell)
    
    if not (spuid and brand in BRANDS and date_str and in_range(date_str)):
        continue
    
    b = BRAND_SHORT[brand]
    month = date_to_month(date_str)
    
    uv = ws.cell(r, uv_col).value if uv_col else 0
    pay_amt = ws.cell(r, pay_amount_col).value if pay_amount_col else 0
    pay_ord = ws.cell(r, pay_orders_col).value if pay_orders_col else 0
    
    def clean_num(v):
        """Extract number from values like '122 (指数)' or '0 (指数)'"""
        if v is None:
            return 0
        if isinstance(v, (int, float)):
            return v
        if isinstance(v, str):
            # Take only the part before space/parenthesis
            import re
            m = re.match(r'[\d.]+', v.strip())
            return float(m.group()) if m else 0
        return 0
    
    uv_val = int(clean_num(uv))
    pay_amt_val = float(clean_num(pay_amt))
    pay_ord_val = int(clean_num(pay_ord))
    
    uv_monthly[b][month]["uv"] += uv_val
    uv_monthly[b][month]["gmv"] += pay_amt_val
    uv_monthly[b][month]["orders"] += pay_ord_val
    
    spu_key = spu_name.get(int(spuid), str(spuid))
    uv_by_spu[b][spu_key]["daily"].append({
        "date": date_str,
        "uv": uv_val,
        "gmv": round(pay_amt_val, 2),
        "orders": pay_ord_val
    })
    uv_by_spu[b][spu_key]["monthly"][month]["uv"] += uv_val
    uv_by_spu[b][spu_key]["monthly"][month]["gmv"] += pay_amt_val
    uv_by_spu[b][spu_key]["monthly"][month]["orders"] += pay_ord_val
    
    uv_processed += 1

print(f"  UV processed {uv_processed} records", file=sys.stderr)

uv_monthly_out = {}
for brand in BRAND_SHORT.values():
    uv_monthly_out[brand] = {}
    for m in sorted(uv_monthly[brand].keys()):
        uv_monthly_out[brand][m] = {
            "uv": uv_monthly[brand][m]["uv"],
            "gmv": round(uv_monthly[brand][m]["gmv"], 2),
            "orders": uv_monthly[brand][m]["orders"]
        }

uv_by_spu_out = {}
for brand in BRAND_SHORT.values():
    uv_by_spu_out[brand] = {}
    for spu_key, data in uv_by_spu[brand].items():
        uv_by_spu_out[brand][spu_key] = {
            "daily": sorted(data["daily"], key=lambda x: x["date"]),
            "monthly": {m: {"uv": d["uv"], "gmv": round(d["gmv"], 2), "orders": d["orders"]} 
                       for m, d in sorted(data["monthly"].items())}
        }

# ============================================================
# 5. 得物推数据 - 按月汇总（双源品牌匹配：货盘表 + 交易订单）
# ============================================================
print("5. Extracting 得物推数据...", file=sys.stderr)
ws = wb["得物推数据"]
# Headers (0-indexed): 时间, 用户ID, 计划名称, 计划ID, 计划类型, 优化目标, 商品ID2, 消耗(元), ...
# openpyxl 1-indexed: col 1=时间, col 7=商品ID2, col 8=消耗(元), col 15=直接支付单量, col 16=直接支付金额, col 17=引导支付单量, col 18=引导支付金额
push_header = {}
for c in range(1, min(30, ws.max_column + 1)):
    h = ws.cell(1, c).value
    if h: push_header[str(h).strip()] = c

time_col = 1
goods_id_col = 7
cost_col = 8
direct_orders_col = 15
direct_gmv_col = 16
indirect_orders_col = 17
indirect_gmv_col = 18
# Note: this sheet has NO brand/huohao columns — use SPU mapping from 货盘表
print(f"  Push: time_col={time_col}, goods_id_col={goods_id_col}, cost_col={cost_col}", file=sys.stderr)

# Build SPU->brand mapping from 交易订单 too (supplement 货盘表)
spu_brand_from_orders = {}
ws_orders = wb["交易订单"]
spu_id_col_ord = 3
brand_col_ord = 8
for r in range(2, ws_orders.max_row + 1):
    spuid = ws_orders.cell(r, spu_id_col_ord).value
    brand = ws_orders.cell(r, brand_col_ord).value
    if spuid and brand in BRANDS:
        try:
            spu_brand_from_orders[int(spuid)] = BRAND_SHORT[brand]
        except (ValueError, TypeError):
            pass
print(f"  Orders SPU→brand mapping: {len(spu_brand_from_orders)} SPUs", file=sys.stderr)

push_monthly = defaultdict(lambda: defaultdict(float))
push_daily = []  # Per-day per-goods push records for Module 5
push_processed = 0
for r in range(2, ws.max_row + 1):
    dt = ws.cell(r, time_col).value
    date_str = excel_to_date(dt)
    good_id = ws.cell(r, goods_id_col).value
    cost = ws.cell(r, cost_col).value
    
    if not (date_str and good_id and cost and in_range(date_str)):
        continue
    
    # Match brand: use SPU mapping from 货盘表 / 交易订单
    brand = None
    huohao = None
    try:
        gid = int(good_id)
        brand = spu_brand.get(gid) or spu_brand_from_orders.get(gid)
        huohao = spu_name.get(gid)
    except (ValueError, TypeError):
        pass
    
    if not brand:
        continue
    
    if not huohao:
        huohao = str(good_id)
    
    month = date_to_month(date_str)
    cost_val = float(cost) if cost else 0
    
    push_monthly[brand][month] += cost_val
    
    # Build daily record for Module 5
    def safe_float(v):
        try: return float(v) if v else 0.0
        except: return 0.0
    def safe_int(v):
        try: return int(float(v)) if v else 0
        except: return 0
    
    push_daily.append({
        "date": date_str,
        "brand": brand,
        "huohao": huohao,
        "cost": round(cost_val, 2),
        "direct_orders": safe_int(ws.cell(r, direct_orders_col).value),
        "direct_gmv": round(safe_float(ws.cell(r, direct_gmv_col).value), 2),
        "indirect_orders": safe_int(ws.cell(r, indirect_orders_col).value),
        "indirect_gmv": round(safe_float(ws.cell(r, indirect_gmv_col).value), 2)
    })
    push_processed += 1

print(f"  Push processed {push_processed} records (daily: {len(push_daily)})", file=sys.stderr)

push_monthly_out = {}
for brand in BRAND_SHORT.values():
    push_monthly_out[brand] = {m: round(v, 2) for m, v in sorted(push_monthly[brand].items())}

# ============================================================
# 6. 社区投放任务 - 按月汇总
# ============================================================
print("6. Extracting 社区投放任务...", file=sys.stderr)
ws = wb["社区投放任务"]
# Headers: 任务月份, 父任务ID, 子任务ID, 任务名称, 任务发布时间, 任务推广形式, 任务模式, 任务状态,
#   任务完成时间, 任务金额, 实际任务金额, 合作达人, 达人uid, ... 匹配货号(29)
comm_header = {}
for c in range(1, min(35, ws.max_column + 1)):
    h = ws.cell(1, c).value
    if h: comm_header[str(h).strip()] = c

# Find columns
task_month_col = comm_header.get("任务月份", 1)
match_goods_col = comm_header.get("匹配货号", 29)
amount_col_comm = comm_header.get("实际任务金额", 11)
time_col_comm = comm_header.get("动态发布时间", 19)
status_col = comm_header.get("任务状态", 8)

print(f"  Comm cols: month={task_month_col}, match_goods={match_goods_col}, amount={amount_col_comm}, time={time_col_comm}, status={status_col}", file=sys.stderr)

comm_monthly = defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "tasks": 0}))
comm_processed = 0
for r in range(2, ws.max_row + 1):
    match_goods = ws.cell(r, match_goods_col).value
    amount_str = ws.cell(r, amount_col_comm).value
    dt = ws.cell(r, time_col_comm).value
    date_str = excel_to_date(dt)
    
    if not (match_goods and amount_str and date_str and in_range(date_str)):
        continue
    
    # Parse amount
    try:
        if isinstance(amount_str, str):
            amount_str = amount_str.replace("¥", "").replace("元", "").replace(",", "").strip()
        amount = float(amount_str)
    except:
        continue
    
    # Match brand from goods name keywords
    goods_list = str(match_goods).split(",")
    brands_for_task = set()
    for g in goods_list:
        g = g.strip()
        # CASIO keywords
        if any(kw in g for kw in ['CASIO', '卡西欧', 'GA-', 'W-', 'F-', 'A1', 'AE-', 'WS-', 'DW-', 'LTP']):
            brands_for_task.add('卡西欧')
        # COACH keywords
        elif any(kw in g for kw in ['COACH', '蔻驰', 'Coach']):
            brands_for_task.add('蔻驰')
    
    if not brands_for_task:
        continue
    
    # Split cost evenly among brands
    month = date_to_month(date_str)
    per_brand = amount / len(brands_for_task)
    for b in brands_for_task:
        comm_monthly[b][month]["cost"] += per_brand
        comm_monthly[b][month]["tasks"] += 1
    
    comm_processed += 1

print(f"  Comm processed {comm_processed} records", file=sys.stderr)

comm_monthly_out = {}
for brand in BRAND_SHORT.values():
    comm_monthly_out[brand] = {}
    for m in sorted(comm_monthly[brand].keys()):
        comm_monthly_out[brand][m] = {
            "cost": round(comm_monthly[brand][m]["cost"], 2),
            "tasks": comm_monthly[brand][m]["tasks"]
        }

# ============================================================
# 6b. 社区投放任务 - 按货号明细（新增模块七数据）
# ============================================================
print("6b. Extracting 社区投放任务-货号明细...", file=sys.stderr)
comm_tasks = []

for r in range(2, ws.max_row + 1):
    # 任务月份 (col 1) - may be Excel serial or datetime
    task_month_val = ws.cell(r, 1).value
    task_month = excel_to_date(task_month_val)
    if task_month:
        task_month = task_month[:7]  # YYYY-MM
    
    # 动态发布时间 (col 19)
    pub_date_val = ws.cell(r, 19).value
    pub_date = excel_to_date(pub_date_val)
    
    # 实际任务金额 (col 11)
    amount_val = ws.cell(r, 11).value
    if amount_val is None:
        continue
    try:
        if isinstance(amount_val, str):
            amount_val = amount_val.replace("¥", "").replace("元", "").replace(",", "").strip()
        amount = float(amount_val)
    except:
        continue
    
    # 曝光 (col 21), 阅读数 (col 22), 商详访问 (col 24)
    def safe_int_comm(v):
        if v is None:
            return 0
        if isinstance(v, (int, float)):
            return int(v)
        s = str(v).strip().replace(",", "")
        if s == "暂无" or s == "":
            return 0
        try:
            return int(float(s))
        except:
            return 0
    
    exposure = safe_int_comm(ws.cell(r, 21).value)
    reads = safe_int_comm(ws.cell(r, 22).value)
    visits = safe_int_comm(ws.cell(r, 24).value)
    
    # 匹配货号1 (col 29) and 匹配货号2 (col 30)
    match_goods_1 = ws.cell(r, 29).value
    match_goods_2 = ws.cell(r, 30).value
    
    all_goods = []
    if match_goods_1:
        for g in str(match_goods_1).split(","):
            g = g.strip()
            if g and g != "暂无" and g != "无":
                all_goods.append(g)
    if match_goods_2:
        for g in str(match_goods_2).split(","):
            g = g.strip()
            if g and g != "暂无" and g != "无":
                all_goods.append(g)
    
    if not all_goods:
        continue
    
    # Determine brand per goods item
    for goods in all_goods:
        brand = None
        if any(kw in goods for kw in ['CASIO', '卡西欧', 'GA-', 'W-', 'F-', 'A1', 'AE-', 'WS-', 'DW-', 'LTP']):
            brand = '卡西欧'
        elif any(kw in goods for kw in ['COACH', '蔻驰', 'Coach']):
            brand = '蔻驰'
        
        if not brand:
            continue
        
        # Split amount evenly among goods in this task
        per_amount = round(amount / len(all_goods), 2)
        
        comm_tasks.append({
            "brand": brand,
            "huohao": goods,
            "month": task_month or "",
            "pub_date": pub_date or "",
            "amount": per_amount,
            "exposure": exposure,
            "reads": reads,
            "visits": visits
        })

print(f"  Comm task records (by huohao): {len(comm_tasks)}", file=sys.stderr)
# Show sample
sample_brands = {}
for t in comm_tasks:
    b = t["brand"]
    if b not in sample_brands:
        sample_brands[b] = 0
    sample_brands[b] += 1
for b, c in sample_brands.items():
    print(f"    {b}: {c} records", file=sys.stderr)

# ============================================================
# 7. Build output (supplemental fields added separately)
# ============================================================
months_list = sorted(set(
    list(orders_monthly_out.get("卡西欧", {}).keys()) +
    list(orders_monthly_out.get("蔻驰", {}).keys()) +
    list(market_monthly_avg.keys())
))

output = {
    "dateRange": {"start": START_DATE, "end": END_DATE},
    "months": months_list,
    "orders_monthly": orders_monthly_out,
    "orders_by_spu": orders_by_spu_out,
    "uv_monthly": uv_monthly_out,
    "uv_by_spu": uv_by_spu_out,
    "market": market_data,
    "market_monthly_avg": market_monthly_avg,
    "market_bag": market_data_bag,
    "market_monthly_bag_avg": market_monthly_bag_avg,
    "push_monthly": push_monthly_out,
    "push_daily": push_daily,
    "comm_monthly": comm_monthly_out,
    "comm_tasks": comm_tasks
}

# Write as data.js
OUTPUT_PATH = "/home/Vic/dewu-reports/XiguoAnalysis/2026/data.js"
with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    f.write("var D=")
    json.dump(output, f, ensure_ascii=False, separators=(",", ":"))
    f.write(";\n")

print(f"\n✅ Written to {OUTPUT_PATH}", file=sys.stderr)
print(f"  Size: {len(json.dumps(output, ensure_ascii=False))} chars", file=sys.stderr)
print(f"  Market records: {len(market_data)}", file=sys.stderr)
print(f"  Months: {months_list}", file=sys.stderr)

# ⚠️ IMPORTANT: data.js generated by openpyxl above is MISSING three fields
# needed by app.js 模块二/三/四/五:
#   - uv_daily_by_brand
#   - uv_spu_data
#   - huohao_daily
# Run the pandas supplement script to inject these:
#   python3 /home/Vic/dewu-reports/XiguoAnalysis/supplement_data.py
