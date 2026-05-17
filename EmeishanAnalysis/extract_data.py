#!/usr/bin/env python3
"""峨眉山数据提取 - 优化版（仅读必要列）"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from collections import defaultdict
import json, re, sys, os, time

EXCEL = "/mnt/e/Obsidian本地仓库/09-数据源/峨眉山数据源收集表.xlsx"
START = "2025-05-01"; END = "2026-05-15"; BRAND = "峨眉山"
OUT = "/home/Vic/dewu-reports/EmeishanAnalysis/2026/data.js"
os.makedirs(os.path.dirname(OUT), exist_ok=True)

def td(v):
    if pd.isna(v): return None
    if isinstance(v, datetime): return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        s=v.strip()
        for f in ["%Y-%m-%d %H:%M:%S","%Y-%m-%d","%Y/%m/%d"]:
            try: return datetime.strptime(s[:19] if len(s)>=19 else s,f).strftime("%Y-%m-%d")
            except: continue
        if len(s)>=10 and s[4]=='-' and s[7]=='-': return s[:10]
    if isinstance(v,(int,float)) and v>40000:
        return (datetime(1899,12,30)+timedelta(days=int(v))).strftime("%Y-%m-%d")
    return None

def tm(v):
    d=td(v); return d[:7] if d else None

def ok(d): return bool(d and START<=d<=END)
def dm(d): return d[:7]

def cn(v):
    if pd.isna(v): return 0.0
    if isinstance(v,(int,float)): return float(v) if v==v else 0.0
    if isinstance(v,str):
        m=re.match(r'[\d.]+',v.strip()); return float(m.group()) if m else 0.0
    return 0.0

def ci(v): return int(cn(v))

t0=time.time()
print("=== 峨眉山数据提取 ===", flush=True)

# ── 1. 大盘指数 ──
print("1. 大盘指数...", flush=True)
dfm=pd.read_excel(EXCEL, sheet_name='大盘指数')
mkt_cats={}
for c in dfm.columns[1:]:
    n=re.sub(r'大盘成交指数[（(]万元[）)]','',str(c)).strip()
    mkt_cats[c]=n

mkt={n:[] for n in mkt_cats.values()}
mkt_ml=defaultdict(lambda: defaultdict(list))
for _,r in dfm.iterrows():
    d=td(r.iloc[0])
    if not d or not ok(d): continue
    m=dm(d)
    for i,c in enumerate(dfm.columns[1:]):
        v=r.iloc[i+1]
        if not pd.isna(v):
            val=round(float(v),2); n=mkt_cats[c]
            mkt[n].append({"date":d,"value":val})
            mkt_ml[n][m].append(val)
mkt_avg={n:{m:round(sum(v)/len(v),2) for m,v in sorted(mkt_ml[n].items())} for n in mkt_cats.values()}
print(f"  Done ({time.time()-t0:.1f}s)", flush=True)

# ── 2. 交易订单 ──
print("2. 交易订单...", flush=True)
df=pd.read_excel(EXCEL, sheet_name='交易订单',
    usecols=['下单日期','下单月份','买家支付时间','订单状态','spuID','货号','商品名称','品牌','数量','出价金额（元）','类目'])
print(f"  {len(df)} rows ({time.time()-t0:.1f}s)", flush=True)

VS={"交易成功","待发货","待收货","待买家收货","待平台发货","待卖家发货","待平台收货","待物流揽收","待卖家发货 (处理中)"}
RS={"交易关闭成功","交易失败"}

om=defaultdict(lambda: {"gmv":0,"o":0})
ocm=defaultdict(lambda: defaultdict(lambda: {"gmv":0,"o":0}))
oh=defaultdict(lambda: {"gmv":0,"o":0,"c":"","n":""})
hd=defaultdict(lambda: defaultdict(lambda: {"gmv":0,"o":0}))
rm=defaultdict(lambda: {"gmv":0,"o":0})
rcm=defaultdict(lambda: defaultdict(lambda: {"gmv":0,"o":0}))
ac=defaultdict(lambda: {"gmv":0,"o":0})
pv=rv=0

for _,r in df.iterrows():
    spuid=r['spuID']; hh=str(r['货号']).strip() if pd.notna(r['货号']) else ""
    amt=cn(r['出价金额（元）']); st=str(r['订单状态']).strip() if pd.notna(r['订单状态']) else ""
    ct=str(r['类目']).strip() if pd.notna(r['类目']) else ""
    nm=str(r['商品名称']).strip() if pd.notna(r['商品名称']) else ""
    qt=int(cn(r['数量']) or 1)
    d=td(r['下单日期'])
    if not d: d=td(r['买家支付时间'])
    if not d: d=tm(r['下单月份']); d=d+"-01" if d else None
    if not (pd.notna(spuid) and amt>0 and ok(d)): continue
    mo=dm(d); gmv=amt*qt
    if st in VS:
        om[mo]["gmv"]+=gmv; om[mo]["o"]+=1
        if ct:
            ocm[mo][ct]["gmv"]+=gmv; ocm[mo][ct]["o"]+=1
            ac[ct]["gmv"]+=gmv; ac[ct]["o"]+=1
        if hh:
            oh[hh]["gmv"]+=gmv; oh[hh]["o"]+=1
            oh[hh]["c"]=ct; oh[hh]["n"]=nm
            hd[d][hh]["gmv"]+=gmv; hd[d][hh]["o"]+=1
        pv+=1
    elif st in RS:
        rm[mo]["gmv"]+=gmv; rm[mo]["o"]+=1
        if ct: rcm[mo][ct]["gmv"]+=gmv; rcm[mo][ct]["o"]+=1
        rv+=1

tc5=sorted(ac.items(),key=lambda x:-x[1]["gmv"])[:5]
TC5=[c for c,_ in tc5]
mo_list=sorted(om.keys())
print(f"  Valid:{pv} Return:{rv} Months:{len(mo_list)} ({time.time()-t0:.1f}s)", flush=True)

# ── 3. 商详访客 ──
print("3. 商详访客...", flush=True)
du=pd.read_excel(EXCEL, sheet_name='商详访客数据',
    usecols=['日期','SPUID','商品标题','商品货号','品牌名称','类目名称','支付金额','支付订单数','商详访问人数'])
print(f"  {len(du)} rows ({time.time()-t0:.1f}s)", flush=True)

uvm=defaultdict(lambda: {"uv":0,"gmv":0,"o":0})
uvd=defaultdict(lambda: {"uv":0,"gmv":0,"o":0})
uvs=defaultdict(lambda: {"dl":[],"ml":defaultdict(lambda: {"uv":0,"gmv":0,"o":0}),"hh":"","nm":""})
uvcm=defaultdict(lambda: defaultdict(lambda: {"uv":0,"gmv":0,"o":0}))

for _,r in du.iterrows():
    spuid=r['SPUID']; brand=str(r['品牌名称']).strip() if pd.notna(r['品牌名称']) else ""
    ct=str(r['类目名称']).strip() if pd.notna(r['类目名称']) else ""
    nm=str(r['商品标题']).strip() if pd.notna(r['商品标题']) else ""
    hh=str(r['商品货号']).strip() if pd.notna(r['商品货号']) else ""
    d=td(r['日期'])
    if not d or not ok(d) or brand!=BRAND: continue
    mo=dm(d); uv=ci(r['商详访问人数']); pa=cn(r['支付金额']); po=ci(r['支付订单数'])
    uvm[mo]["uv"]+=uv; uvm[mo]["gmv"]+=pa; uvm[mo]["o"]+=po
    uvd[d]["uv"]+=uv; uvd[d]["gmv"]+=pa; uvd[d]["o"]+=po
    if ct: uvcm[mo][ct]["uv"]+=uv; uvcm[mo][ct]["gmv"]+=pa; uvcm[mo][ct]["o"]+=po
    if pd.notna(spuid) and nm:
        sk=str(nm); us=uvs[sk]
        us["dl"].append({"date":d,"uv":uv,"gmv":round(pa,2),"o":po})
        us["ml"][mo]["uv"]+=uv; us["ml"][mo]["gmv"]+=pa; us["ml"][mo]["o"]+=po
        us["hh"]=hh; us["nm"]=nm

print(f"  UV: {len(uvd)} days, {len(uvs)} SPUs ({time.time()-t0:.1f}s)", flush=True)

# ── 4. 得物推 ──
print("4. 得物推...", flush=True)
dp=pd.read_excel(EXCEL, sheet_name='得物推数据', usecols=['日期','消耗()'])
pm=defaultdict(float)
for _,r in dp.iterrows():
    d=td(r['日期']); cst=cn(r['消耗()'])
    if d and ok(d) and cst>0: pm[dm(d)]+=cst
print(f"  {len(pm)} months ({time.time()-t0:.1f}s)", flush=True)

# ── 5. 社区投放 ──
print("5. 社区投放...", flush=True)
dt=pd.read_excel(EXCEL, sheet_name='社区投放任务', usecols=['任务月份','子任务ID','任务状态','任务发布时间','任务金额'])
dc=pd.read_excel(EXCEL, sheet_name='社区投放', usecols=['子任务ID','任务实际金额'])
print(f"  Task:{len(dt)} Comm:{len(dc)} ({time.time()-t0:.1f}s)", flush=True)

EXCL={"关闭","待发货","待收货","待发布","待支付"}
ti={}
for _,r in dt.iterrows():
    sid=str(r['子任务ID']).strip() if pd.notna(r['子任务ID']) else ""
    sts=str(r['任务状态']).strip() if pd.notna(r['任务状态']) else ""
    if sts in EXCL: continue
    m=tm(r['任务月份'])
    if not m:
        d=td(r['任务发布时间'])
        if d: m=dm(d)
    amt_s=str(r['任务金额']).strip() if pd.notna(r['任务金额']) else "0"
    amt_s=amt_s.replace("元","").replace("¥","").replace(",","").strip()
    try: amt=float(amt_s)
    except: amt=0
    if sid and m: ti[sid]={"m":m,"amt":amt}

cm=defaultdict(lambda: {"c":0,"t":0})
for _,info in ti.items():
    m=info["m"]; cm[m]["c"]+=info["amt"]; cm[m]["t"]+=1
print(f"  Comm:{len(cm)} months ({time.time()-t0:.1f}s)", flush=True)

# ── 6. 输出 ──
print("6. Building output...", flush=True)
cs={}
for ct in ac:
    p=ct.split('-'); cs[ct]=p[-1] if len(p)>1 else ct

out={
    "brand":BRAND,
    "dr":{"s":mo_list[0]+"-01" if mo_list else START,"e":mo_list[-1]+"-01" if mo_list else END},
    "months":mo_list,
    "topCategories":TC5,
    "orders_monthly":{m:{"gmv":round(om[m]["gmv"],2),"o":om[m]["o"]} for m in mo_list},
    "orders_by_category":{m:{c:{"gmv":round(d["gmv"],2),"o":d["o"]} for c,d in sorted(ocm[m].items(),key=lambda x:-x[1]["gmv"])} for m in mo_list},
    "orders_by_huohao":[{"hh":h,"gmv":round(d["gmv"],2),"o":d["o"],"c":d["c"],"n":d["n"]} for h,d in sorted(oh.items(),key=lambda x:-x[1]["gmv"])],
    "huohao_daily":{d:{h:{"gmv":round(v["gmv"],2),"o":v["o"]} for h,v in sorted(hh.items(),key=lambda x:-x[1]["gmv"])} for d,hh in sorted(hd.items())},
    "returns_monthly":{m:{"gmv":round(rm[m]["gmv"],2),"o":rm[m]["o"]} for m in mo_list},
    "returns_by_category":{m:{c:{"gmv":round(d["gmv"],2),"o":d["o"]} for c,d in sorted(rcm[m].items(),key=lambda x:-x[1]["gmv"])} for m in mo_list if m in rcm},
    "uv_monthly":{m:{"uv":uvm[m]["uv"],"gmv":round(uvm[m]["gmv"],2),"o":uvm[m]["o"]} for m in mo_list},
    "uv_daily":[{"date":d,"uv":v["uv"],"gmv":round(v["gmv"],2),"o":v["o"]} for d,v in sorted(uvd.items())],
    "uv_by_spu":sorted([{"spu":k,"hh":d["hh"],"uv":sum(x["uv"] for x in d["dl"]),"gmv":round(sum(x["gmv"] for x in d["dl"]),2),"o":sum(x["o"] for x in d["dl"]),
        "dl":sorted(d["dl"],key=lambda x:x["date"]),
        "ml":{m:{"uv":t["uv"],"gmv":round(t["gmv"],2),"o":t["o"]} for m,t in sorted(d["ml"].items())}} for k,d in uvs.items()],key=lambda x:-x["uv"]),
    "uv_by_category":{m:{c:{"uv":d["uv"],"gmv":round(d["gmv"],2),"o":d["o"]} for c,d in sorted(uvcm[m].items(),key=lambda x:-x[1]["uv"])} for m in mo_list},
    "market":mkt,"market_monthly_avg":mkt_avg,
    "push_monthly":{m:round(v,2) for m,v in sorted(pm.items())},
    "comm_monthly":{m:{"cost":round(d["c"],2),"tasks":d["t"]} for m,d in sorted(cm.items())},
    "cat_short":cs,
    "all_categories":{c:{"gmv":round(d["gmv"],2),"o":d["o"]} for c,d in sorted(ac.items(),key=lambda x:-x[1]["gmv"])},
    "total_gmv":round(sum(v["gmv"] for v in om.values()),2),
    "total_orders":sum(v["o"] for v in om.values()),
}

with open(OUT,"w",encoding="utf-8") as f:
    f.write("var D=")
    json.dump(out,f,ensure_ascii=False,separators=(",",":"))
    f.write(";\n")

sz=os.path.getsize(OUT)
print(f"\n✅ {OUT} ({sz:,} bytes)", flush=True)
print(f"  Months: {mo_list}", flush=True)
print(f"  GMV: {out['total_gmv']:,.0f}, Orders: {out['total_orders']:,}", flush=True)
print(f"  Total time: {time.time()-t0:.0f}s", flush=True)
