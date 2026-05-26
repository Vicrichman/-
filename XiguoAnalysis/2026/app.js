// ==================== 喜过运营分析看板 v2 - 核心脚本 ====================
var MO=D.months,ML=MO.map(function(m){return m.replace("2026-","")+"月"});
var CC="#f59e0b",CHC="#a78bfa",TX="#94a3b8",GG="#1e2235",GRN="#22c55e",RED="#ef4444",BLU="#60a5fa";
var DR=D.dateRange;

// Utility: get monthly data
function gm(b,k){var m=D.orders_monthly[b]||{},r=[];MO.forEach(function(x){r.push(m[x]?m[x][k]||0:0)});return r;}
function uvm(b,k){var m=D.uv_monthly[b]||{},r=[];MO.forEach(function(x){r.push(m[x]?m[x][k]||0:0)});return r;}

// Global marketing variables
var pc=[],pch=[],cc2=[],mc=[],mch=[];
var cacC=[],cacCh=[],roiC=[],roiCh=[];

// Chart references for destroy/recreate
var c2Chart=null,c6Chart=null,c7aChart=null,c7bChart=null,spC=null;

window.addEventListener("load",function(){
// ===== KPI =====
var cg=gm("卡西欧","gmv"),co=gm("卡西欧","orders"),chg=gm("蔻驰","gmv"),cho=gm("蔻驰","orders");
document.getElementById("kv-cgmv").textContent="¥"+(cg.reduce(function(a,b){return a+b},0)/10000).toFixed(1)+"万";
document.getElementById("kv-co").textContent=co.reduce(function(a,b){return a+b},0);
document.getElementById("kv-chgmv").textContent="¥"+(chg.reduce(function(a,b){return a+b},0)/10000).toFixed(1)+"万";
document.getElementById("kv-cho").textContent=cho.reduce(function(a,b){return a+b},0);

// Init date inputs
var ds=DR.start,de=DR.end;
["m2","m3","m5"].forEach(function(p){
  document.getElementById(p+"-ds").value=ds;
  document.getElementById(p+"-de").value=de;
});

// ===== 模块一: UV曲线 + GMV柱状图 + 大盘指数 =====
renderM1();
// ===== 模块二: 商详访客日常趋势 =====
renderM2();
// ===== 模块三: 三栏TOP20 =====
renderM3();
// ===== 模块四: 近三日异动 =====
renderM4();
// ===== 模块五: 货号销售排行 =====
renderM5();
// ===== 模块六: 月度品牌占比 =====
renderM6();
// ===== 模块七: 社区投放任务明细 =====
initM7();
renderM7();
// ===== 模块八: 营销效率 =====
renderM8();
// ===== 模块九: 运营诊断 =====
renderM9();

}); // end window.load

// ================================================================
// 模块一: 品牌月度趋势 (UV曲线 + GMV柱状图 + 大盘指数)
// ================================================================
function renderM1(){
var mktAvg=MO.map(function(m){return D.market_monthly_avg?D.market_monthly_avg[m]||null:null});
var mktBagAvg=MO.map(function(m){return D.market_monthly_bag_avg?D.market_monthly_bag_avg[m]||null:null});
var gmvC=MO.map(function(m){var v=D.orders_monthly["卡西欧"]&&D.orders_monthly["卡西欧"][m]?D.orders_monthly["卡西欧"][m].gmv:0;return v/10000;});
var gmvCH=MO.map(function(m){var v=D.orders_monthly["蔻驰"]&&D.orders_monthly["蔻驰"][m]?D.orders_monthly["蔻驰"][m].gmv:0;return v/10000;});

new Chart(document.getElementById("c1"),{type:"bar",data:{labels:ML,datasets:[
{label:"卡西欧 UV",data:uvm("卡西欧","uv"),borderColor:CC,backgroundColor:"transparent",yAxisID:"y",tension:0.3,borderWidth:2.5,pointRadius:3,pointBackgroundColor:CC,type:"line",order:1},
{label:"蔻驰 UV",data:uvm("蔻驰","uv"),borderColor:CHC,backgroundColor:"transparent",yAxisID:"y",tension:0.3,borderWidth:2.5,borderDash:[5,5],pointRadius:3,pointBackgroundColor:CHC,type:"line",order:2},
{label:"卡西欧 GMV(万)",data:gmvC,backgroundColor:CC+"70",borderColor:CC,borderWidth:1,yAxisID:"y1",order:3},
{label:"蔻驰 GMV(万)",data:gmvCH,backgroundColor:CHC+"70",borderColor:CHC,borderWidth:1,yAxisID:"y1",order:4},
{label:"大盘·日韩表",data:mktAvg,borderColor:"#64748b",backgroundColor:"transparent",yAxisID:"y2",tension:0.3,borderWidth:1.5,pointRadius:4,pointBackgroundColor:"#64748b",type:"line",order:0},
{label:"大盘·单肩包",data:mktBagAvg,borderColor:"#f97316",backgroundColor:"transparent",yAxisID:"y2",tension:0.3,borderWidth:1.5,borderDash:[3,3],pointRadius:4,pointBackgroundColor:"#f97316",type:"line",order:0}
]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:"index",intersect:false},plugins:{legend:{labels:{color:TX,usePointStyle:true,padding:14,font:{size:11}},position:"top"},tooltip:{callbacks:{label:function(ctx){var v=ctx.raw;if(ctx.dataset.label.indexOf("UV")>=0)return ctx.dataset.label+": "+v.toLocaleString();if(ctx.dataset.label.indexOf("大盘")>=0)return ctx.dataset.label+": "+v;return ctx.dataset.label+": ¥"+v.toFixed(1)+"万"}}}},scales:{y:{position:"left",title:{display:true,text:"UV",color:TX},grid:{color:GG},ticks:{color:TX,callback:function(v){return v>=1000?(v/1000).toFixed(0)+"k":v}}},y1:{position:"right",title:{display:true,text:"GMV(万元)",color:TX},grid:{drawOnChartArea:false},ticks:{color:TX,callback:function(v){return"¥"+v.toFixed(1)}},beginAtZero:true},y2:{position:"left",title:{display:true,text:"大盘指数",color:"#64748b"},grid:{drawOnChartArea:false},ticks:{color:"#64748b"},offset:true},x:{ticks:{color:TX},grid:{color:GG}}}}});
}

// ================================================================
// 模块二: 商详访客日常趋势 (UV曲线 + GMV柱状图 + 订单数柱状图)
// ================================================================
function renderM2(){
if(c2Chart){c2Chart.destroy();c2Chart=null;}
var brand=document.getElementById("m2-brand").value;
var ds=document.getElementById("m2-ds").value||DR.start;
var de=document.getElementById("m2-de").value||DR.end;
var daily=D.uv_daily_by_brand[brand]||[];
daily=daily.filter(function(d){return d.date>=ds&&d.date<=de});
if(daily.length===0){
  var parent=document.getElementById("c2").parentElement;
  parent.innerHTML='<canvas id="c2"></canvas><div style="text-align:center;color:#94a3b8;padding:60px">暂无数据</div>';
  return;
}
var lbs=daily.map(function(d){return d.date.slice(5)});
var uv=daily.map(function(d){return d.uv});
var gmv=daily.map(function(d){return d.gmv/10000});
var ord=daily.map(function(d){return d.orders});
c2Chart=new Chart(document.getElementById("c2"),{type:"bar",data:{labels:lbs,datasets:[
{label:"UV",data:uv,borderColor:BLU,backgroundColor:"transparent",yAxisID:"y",tension:0.3,borderWidth:2,pointRadius:0,type:"line",order:1},
{label:"GMV(万)",data:gmv,backgroundColor:brand==="卡西欧"?CC+"80":CHC+"80",borderColor:brand==="卡西欧"?CC:CHC,borderWidth:1,yAxisID:"y1",order:2},
{label:"订单数",data:ord,backgroundColor:"#22c55e50",borderColor:"#22c55e",borderWidth:1,yAxisID:"y2",order:3}
]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:"index",intersect:false},plugins:{legend:{labels:{color:TX,usePointStyle:true},position:"top"},tooltip:{callbacks:{label:function(ctx){var v=ctx.raw;if(ctx.dataset.label==="UV")return"UV: "+v.toLocaleString();if(ctx.dataset.label==="订单数")return"订单: "+v+"单";return ctx.dataset.label+": ¥"+v.toFixed(2)+"万"}}}},scales:{y:{position:"left",title:{display:true,text:"UV",color:TX},grid:{color:GG},ticks:{color:TX,callback:function(v){return v>=1000?(v/1000).toFixed(0)+"k":v}}},y1:{position:"right",title:{display:true,text:"GMV(万元)",color:TX},grid:{drawOnChartArea:false},ticks:{color:TX,callback:function(v){return"¥"+v.toFixed(1)}},beginAtZero:true},y2:{position:"right",title:{display:true,text:"订单数",color:"#22c55e"},grid:{drawOnChartArea:false},ticks:{color:"#22c55e"},beginAtZero:true},x:{ticks:{color:TX,maxTicksLimit:20},grid:{color:GG}}}}});
}

// ================================================================
// 模块三: 三栏TOP20 (商详访客UV / 支付订单量 / 支付GMV)
// ================================================================
function renderM3(){
var brand=document.getElementById("m3-brand").value;
var ds=document.getElementById("m3-ds").value||DR.start;
var de=document.getElementById("m3-de").value||DR.end;
var spus=D.uv_spu_data[brand]||[];

// Aggregate by SPU within date range
var agg={};
spus.forEach(function(s){
  var u=0,g=0,o=0;
  s.daily.forEach(function(d){if(d.date>=ds&&d.date<=de){u+=d.uv;g+=d.gmv;o+=d.orders;}});
  if(u>0||g>0||o>0) agg[s.spu]={spu:s.spu,huohao:s.huohao,uv:u,gmv:g,orders:o};
});

var list=Object.values(agg);
var uvTop=list.slice().sort(function(a,b){return b.uv-a.uv}).slice(0,20);
var ordTop=list.slice().sort(function(a,b){return b.orders-a.orders}).slice(0,20);
var gmvTop=list.slice().sort(function(a,b){return b.gmv-a.gmv}).slice(0,20);

function tbl(arr,type,container){
  var el=document.getElementById(container);
  if(arr.length===0){el.innerHTML='<div style="color:#94a3b8;font-size:12px;text-align:center;padding:20px">暂无数据</div>';return;}
  var rows=arr.map(function(item,i){
    var full=item.spu.replace(/"/g,"&quot;");
    var displayName=brand==="蔻驰"?(item.huohao||item.spu):item.spu;
    var nm=displayName.length>25?displayName.slice(0,25)+"...":displayName;
    var v=type==="uv"?item.uv.toLocaleString():type==="gmv"?"¥"+item.gmv.toLocaleString():item.orders+"单";
    return'<tr><td style="color:#94a3b8;width:24px">'+(i+1)+'</td><td><span class="slnk" data-brand="'+brand+'" data-spu="'+full+'">'+nm+'</span></td><td style="text-align:right">'+v+'</td></tr>';
  });
  el.innerHTML='<table style="font-size:12px"><tbody>'+rows.join("")+'</tbody></table>';
}

tbl(uvTop,"uv","m3-uv-tbl");
tbl(ordTop,"ord","m3-ord-tbl");
tbl(gmvTop,"gmv","m3-gmv-tbl");
}

// ================================================================
// 模块四: 近三日异动款式
// ================================================================
function renderM4(){
var grid=document.getElementById("m4-grid");
var allSpus=(D.uv_spu_data["卡西欧"]||[]).concat(D.uv_spu_data["蔻驰"]||[]);

// Find all unique dates and get last 3
var allDates=new Set();
allSpus.forEach(function(s){s.daily.forEach(function(d){allDates.add(d.date);});});
var sortedDates=Array.from(allDates).sort();
if(sortedDates.length<6){grid.innerHTML='<div style="color:#94a3b8;padding:20px">数据不足，需要至少6天数据</div>';return;}
var last3=sortedDates.slice(-3);
var prev3=sortedDates.slice(-6,-3);

function sum(arr,field,dates){var s=0;arr.forEach(function(d){if(dates.indexOf(d.date)>=0)s+=d[field];});return s;}

var anomalies=[];
allSpus.forEach(function(s){
  var uvR=sum(s.daily,"uv",last3),uvP=sum(s.daily,"uv",prev3);
  var ordR=sum(s.daily,"orders",last3),ordP=sum(s.daily,"orders",prev3);
  var gmvR=sum(s.daily,"gmv",last3),gmvP=sum(s.daily,"gmv",prev3);

  var uvCh=uvP>0?((uvR-uvP)/uvP*100):(uvR>0?999:0);
  var ordCh=ordP>0?((ordR-ordP)/ordP*100):(ordR>0?999:0);

  if(Math.abs(uvCh)>=30||Math.abs(ordCh)>=30){
    var brand=s.spu.indexOf("Coach")>=0||s.huohao.indexOf("COACH")>=0?"蔻驰":"卡西欧";
    if(brand==="卡西欧"&&D.uv_spu_data["蔻驰"]){
      var found=D.uv_spu_data["蔻驰"].some(function(x){return x.spu===s.spu;});
      if(found)brand="蔻驰";
    }
    anomalies.push({spu:s.spu,huohao:s.huohao,brand:brand,uvR:uvR,uvP:uvP,uvCh:uvCh,ordR:ordR,ordP:ordP,ordCh:ordCh,gmvR:gmvR,gmvP:gmvP});
  }
});

// Sort by absolute change magnitude
anomalies.sort(function(a,b){return (Math.abs(b.uvCh)+Math.abs(b.ordCh))-(Math.abs(a.uvCh)+Math.abs(a.ordCh));});
anomalies=anomalies.slice(0,20);

if(anomalies.length===0){
  grid.innerHTML='<div style="color:#94a3b8;padding:20px;grid-column:1/-1;text-align:center">近期无异动款式</div>';
  return;
}

var bc=brandColor;
var html='';
anomalies.forEach(function(a){
  var uvArrow=a.uvCh>=0?"↑":"↓",uvCls=a.uvCh>=0?"up":"down";
  var ordArrow=a.ordCh>=0?"↑":"↓",ordCls=a.ordCh>=0?"up":"down";
  var full=a.spu.replace(/"/g,"&quot;");
  var nm=a.spu.length>35?a.spu.slice(0,35)+"...":a.spu;
  var bc=a.brand==="卡西欧"?CC:CHC;
  html+='<div class="anomaly-item">';
  html+='<div class="spu" style="color:'+bc+'"><span class="slnk" data-brand="'+a.brand+'" data-spu="'+full+'">'+nm+'</span></div>';
  html+='<div style="color:#94a3b8;font-size:11px">货号: '+a.huohao+'</div>';
  html+='<div class="stat">UV: '+a.uvP.toLocaleString()+' → '+a.uvR.toLocaleString()+' <span class="change '+uvCls+'">'+uvArrow+Math.abs(a.uvCh).toFixed(0)+'%</span></div>';
  html+='<div class="stat">订单: '+a.ordP+' → '+a.ordR+' <span class="change '+ordCls+'">'+ordArrow+Math.abs(a.ordCh).toFixed(0)+'%</span></div>';
  html+='<div class="stat" style="font-size:11px">近3日GMV: ¥'+a.gmvR.toLocaleString()+' | 前3日: ¥'+a.gmvP.toLocaleString()+'</div>';
  html+='</div>';
});
grid.innerHTML=html;
}

function brandColor(b){return b==="卡西欧"?CC:CHC;}

// ================================================================
// 模块五: 货号销售排行（使用交易订单货号 + 日期筛选）
// ================================================================
function renderM5(){
var brand=document.getElementById("m5-brand").value;
var ds=document.getElementById("m5-ds").value||DR.start;
var de=document.getElementById("m5-de").value||DR.end;
var daily=D.huohao_daily[brand]||[];

// Aggregate within date range
var agg={};
daily.forEach(function(d){if(d.date>=ds&&d.date<=de){
  var hh=d.huohao;
  if(!agg[hh])agg[hh]={huohao:hh,gmv:0,orders:0};
  agg[hh].gmv+=d.gmv;
  agg[hh].orders+=d.orders;
}});

var list=Object.values(agg).sort(function(a,b){return b.gmv-a.gmv});
var total=list.reduce(function(s,a){return s+a.gmv;},0);
var el=document.getElementById("m5-tbl");

if(list.length===0){el.innerHTML='<div style="color:#94a3b8;text-align:center;padding:20px">暂无数据</div>';return;}

var h3=brand==="卡西欧"?'<h3 style="color:#f59e0b;font-size:14px;margin-bottom:8px">卡西欧 · 货号排行</h3>':'<h3 style="color:#a78bfa;font-size:14px;margin-bottom:8px">蔻驰 · 货号排行</h3>';
var rows=list.map(function(item,i){
  var pct=total>0?(item.gmv/total*100).toFixed(1):"0.0";
  var nm=item.huohao.length>45?item.huohao.slice(0,45)+"...":item.huohao;
  return'<tr><td style="color:#94a3b8;width:28px">'+(i+1)+'</td><td>'+nm+'</td><td style="text-align:right">¥'+item.gmv.toLocaleString()+'</td><td style="text-align:right">'+item.orders+'单</td><td style="text-align:right">'+pct+'%</td></tr>';
});
el.innerHTML=h3+'<table><thead><tr><th style="width:28px">#</th><th>货号</th><th style="text-align:right">GMV</th><th style="text-align:right">订单</th><th style="text-align:right">占比</th></tr></thead><tbody>'+rows.join("")+'</tbody></table>';
}

// ================================================================
// 模块六: 月度品牌销售占比（堆叠柱状图）
// ================================================================
function renderM6(){
var sd=MO.map(function(m){
var c=0,ch=0;
if(D.orders_monthly["卡西欧"]&&D.orders_monthly["卡西欧"][m])c=D.orders_monthly["卡西欧"][m].gmv||0;
if(D.orders_monthly["蔻驰"]&&D.orders_monthly["蔻驰"][m])ch=D.orders_monthly["蔻驰"][m].gmv||0;
return {c:c,ch:ch,total:c+ch};
});
c6Chart=new Chart(document.getElementById("c6"),{type:"bar",data:{labels:ML,datasets:[
{label:"卡西欧",data:sd.map(function(d){return d.c}),backgroundColor:CC+"80",borderColor:CC,borderWidth:1},
{label:"蔻驰",data:sd.map(function(d){return d.ch}),backgroundColor:CHC+"80",borderColor:CHC,borderWidth:1}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:TX,usePointStyle:true}},tooltip:{callbacks:{label:function(ctx){var v=ctx.raw,t=sd[ctx.dataIndex].total,p=t>0?(v/t*100).toFixed(1):"0";return ctx.dataset.label+": ¥"+v.toLocaleString()+" ("+p+"%)";}}}},scales:{x:{stacked:true,ticks:{color:TX},grid:{color:GG}},y:{stacked:true,ticks:{color:TX,callback:function(v){return"¥"+(v/10000).toFixed(0)+"万"}},grid:{color:GG}}}}});
}

// ================================================================
// 模块七: 社区投放任务明细（按货号汇总，支持日期+月份双筛选）
// ================================================================
function initM7(){
  var tasks = D.comm_tasks || [];
  var months = [];
  tasks.forEach(function(t){ if(t.month && months.indexOf(t.month)<0) months.push(t.month); });
  months.sort();
  var sel = document.getElementById("m7-month");
  months.forEach(function(m){ var o=document.createElement("option"); o.value=m; o.textContent=m; sel.appendChild(o); });
}
function renderM7(){
  var brand = document.getElementById("m7-brand").value;
  var ds = document.getElementById("m7-ds").value;
  var de = document.getElementById("m7-de").value;
  var month = document.getElementById("m7-month").value;
  var tasks = (D.comm_tasks||[]).filter(function(t){return t.brand===brand;});
  if(month) tasks = tasks.filter(function(t){return t.month===month;});
  if(ds) tasks = tasks.filter(function(t){return t.pub_date&&t.pub_date>=ds;});
  if(de) tasks = tasks.filter(function(t){return t.pub_date&&t.pub_date<=de;});
  var agg={};
  tasks.forEach(function(t){
    var hh=t.huohao; if(!agg[hh]) agg[hh]={huohao:hh,tasks:0,amount:0,exposure:0,reads:0,visits:0};
    agg[hh].tasks+=1; agg[hh].amount+=t.amount; agg[hh].exposure+=t.exposure; agg[hh].reads+=t.reads; agg[hh].visits+=t.visits;
  });
  var list=Object.values(agg).sort(function(a,b){return b.amount-a.amount;});
  var el=document.getElementById("m7-tbl").getElementsByTagName("tbody")[0];
  if(list.length===0){el.innerHTML='<tr><td colspan="9" style="color:#94a3b8;text-align:center;padding:20px">暂无数据</td></tr>';return;}
  var rows=list.map(function(item,i){
    var cpe=item.exposure>0?(item.amount/item.exposure).toFixed(2):"—";
    var cpv=item.visits>0?(item.amount/item.visits).toFixed(2):"—";
    return'<tr><td style="color:#94a3b8">'+(i+1)+'</td><td>'+item.huohao+'</td><td style="text-align:right">'+item.tasks+'</td><td style="text-align:right">¥'+item.amount.toLocaleString()+'</td><td style="text-align:right">'+item.exposure.toLocaleString()+'</td><td style="text-align:right">'+item.reads.toLocaleString()+'</td><td style="text-align:right">'+item.visits.toLocaleString()+'</td><td style="text-align:right">'+cpe+'</td><td style="text-align:right">'+cpv+'</td></tr>';
  });
  el.innerHTML=rows.join("");
}
function resetM7(){
  document.getElementById("m7-ds").value=""; document.getElementById("m7-de").value="";
  document.getElementById("m7-month").value=""; renderM7();
}

// ================================================================
// 模块八: 营销效率
// ================================================================
function renderM8(){
var co=gm("卡西欧","orders"),cho=gm("蔻驰","orders");
MO.forEach(function(m,i){pc.push(D.push_monthly["卡西欧"]?D.push_monthly["卡西欧"][m]||0:0);pch.push(D.push_monthly["蔻驰"]?D.push_monthly["蔻驰"][m]||0:0);cc2.push(D.comm_monthly["卡西欧"]&&D.comm_monthly["卡西欧"][m]?D.comm_monthly["卡西欧"][m].cost||0:0);});
MO.forEach(function(m,i){mc.push(pc[i]+cc2[i]);});
mch=pch;
MO.forEach(function(m,i){cacC.push(co[i]>0?mc[i]/co[i]:0);cacCh.push(cho[i]>0?mch[i]/cho[i]:0);roiC.push(mc[i]>0?gm("卡西欧","gmv")[i]/mc[i]:0);roiCh.push(mch[i]>0?gm("蔻驰","gmv")[i]/mch[i]:0);});

c7aChart=new Chart(document.getElementById("c7a"),{type:"bar",data:{labels:ML,datasets:[
{label:"卡西欧",data:cacC,backgroundColor:CC+"60",borderColor:CC,borderWidth:1},
{label:"蔻驰",data:cacCh,backgroundColor:CHC+"60",borderColor:CHC,borderWidth:1}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:TX,usePointStyle:true,font:{size:10}}},title:{display:true,text:"获客成本(¥/单)",color:TX,font:{size:13}}},scales:{y:{ticks:{color:TX,callback:function(v){return"¥"+v.toFixed(0)}},grid:{color:GG}},x:{ticks:{color:TX},grid:{color:GG}}}}});

c7bChart=new Chart(document.getElementById("c7b"),{type:"bar",data:{labels:ML,datasets:[
{label:"卡西欧",data:roiC,backgroundColor:CC+"60",borderColor:CC,borderWidth:1},
{label:"蔻驰",data:roiCh,backgroundColor:CHC+"60",borderColor:CHC,borderWidth:1}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:TX,usePointStyle:true,font:{size:10}}},title:{display:true,text:"ROI (GMV/营销费)",color:TX,font:{size:13}}},scales:{y:{ticks:{color:TX,callback:function(v){return v.toFixed(1)+"x"}},grid:{color:GG}},x:{ticks:{color:TX},grid:{color:GG}}}}});

var mt=document.getElementById("tb-mkt");
MO.forEach(function(m,i){var tr=document.createElement("tr"),cr=roiC[i],chr=roiCh[i],t1=cr>3?"tg":cr>1?"ty":"tr",t2=chr>3?"tg":chr>1?"ty":"tr";
tr.innerHTML="<td>"+m.replace("2026-","")+"月</td><td>¥"+mc[i].toLocaleString()+"</td><td>"+(cacC[i]>0?"¥"+cacC[i].toFixed(0):"—")+"</td><td><span class='tag "+t1+"'>"+(cr>0?cr.toFixed(1)+"x":"—")+"</span></td><td>¥"+mch[i].toLocaleString()+"</td><td>"+(cacCh[i]>0?"¥"+cacCh[i].toFixed(0):"—")+"</td><td><span class='tag "+t2+"'>"+(chr>0?chr.toFixed(1)+"x":"—")+"</span></td>";
mt.appendChild(tr);});
}

// ================================================================
// 模块九: 动态运营诊断与建议（详细版）
// ================================================================
function renderM9(){
var el=document.getElementById("mod8");if(!el)return;

var totalGMV_C=gm("卡西欧","gmv").reduce(function(a,b){return a+b},0);
var totalGMV_CH=gm("蔻驰","gmv").reduce(function(a,b){return a+b},0);
var totalOrd_C=gm("卡西欧","orders").reduce(function(a,b){return a+b},0);
var totalOrd_CH=gm("蔻驰","orders").reduce(function(a,b){return a+b},0);
var asp_C=totalOrd_C>0?totalGMV_C/totalOrd_C:0;
var asp_CH=totalOrd_CH>0?totalGMV_CH/totalOrd_CH:0;
var totalGMV=totalGMV_C+totalGMV_CH;

var totalUV_C=uvm("卡西欧","uv").reduce(function(a,b){return a+b},0);
var totalUV_CH=uvm("蔻驰","uv").reduce(function(a,b){return a+b},0);
var uvOrd_C=uvm("卡西欧","orders").reduce(function(a,b){return a+b},0);
var uvOrd_CH=uvm("蔻驰","orders").reduce(function(a,b){return a+b},0);

var cvr_C=totalUV_C>0?(uvOrd_C/totalUV_C*100):0;
var cvr_CH=totalUV_CH>0?(uvOrd_CH/totalUV_CH*100):0;

var totalPushC=pc.reduce(function(a,b){return a+b},0);
var totalPushCH=pch.reduce(function(a,b){return a+b},0);
var totalCommC=cc2.reduce(function(a,b){return a+b},0);
var totalMktC=totalPushC+totalCommC;
var totalMktCH=totalPushCH;
var roiAvgC=totalMktC>0?totalGMV_C/totalMktC:0;
var roiAvgCH=totalMktCH>0?totalGMV_CH/totalMktCH:0;

var gmvArr_C=gm("卡西欧","gmv");
var trendC=gmvArr_C.length>=2?(gmvArr_C[gmvArr_C.length-1]/Math.max(1,gmvArr_C[gmvArr_C.length-2])-1)*100:0;
var gmvArr_CH=gm("蔻驰","gmv");
var trendCH=gmvArr_CH.length>=2?(gmvArr_CH[gmvArr_CH.length-1]/Math.max(1,gmvArr_CH[gmvArr_CH.length-2])-1)*100:0;

var spuCountC=(D.orders_by_spu["卡西欧"]||[]).length;
var spuCountCH=(D.orders_by_spu["蔻驰"]||[]).length;

var top3GMV_C=0,top3GMV_CH=0;
(D.orders_by_spu["卡西欧"]||[]).slice(0,3).forEach(function(s){top3GMV_C+=s.gmv});
(D.orders_by_spu["蔻驰"]||[]).slice(0,3).forEach(function(s){top3GMV_CH+=s.gmv});
var concC=totalGMV_C>0?(top3GMV_C/totalGMV_C*100):0;
var concCH=totalGMV_CH>0?(top3GMV_CH/totalGMV_CH*100):0;

var mktVals=D.market_monthly_avg?Object.values(D.market_monthly_avg):[];
var mktTrend=mktVals.length>=2?((mktVals[mktVals.length-1]/mktVals[0])-1)*100:0;

var chPct=totalGMV>0?(totalGMV_CH/totalGMV*100):0;

var lastM=MO[MO.length-1];
var prevM=MO.length>1?MO[MO.length-2]:lastM;
var lmC=D.orders_monthly["卡西欧"]&&D.orders_monthly["卡西欧"][lastM]?D.orders_monthly["卡西欧"][lastM].gmv:0;
var lmCH=D.orders_monthly["蔻驰"]&&D.orders_monthly["蔻驰"][lastM]?D.orders_monthly["蔻驰"][lastM].gmv:0;
var pmC=D.orders_monthly["卡西欧"]&&D.orders_monthly["卡西欧"][prevM]?D.orders_monthly["卡西欧"][prevM].gmv:0;
var pmCH=D.orders_monthly["蔻驰"]&&D.orders_monthly["蔻驰"][prevM]?D.orders_monthly["蔻驰"][prevM].gmv:0;
var lmTotal=lmC+lmCH;
var lmChPct=lmTotal>0?(lmCH/lmTotal*100):0;

// UV monthly trends
var uvArr_C=uvm("卡西欧","uv");
var uvTrendC=uvArr_C.length>=2?(uvArr_C[uvArr_C.length-1]/Math.max(1,uvArr_C[uvArr_C.length-2])-1)*100:0;
var uvArr_CH=uvm("蔻驰","uv");
var uvTrendCH=uvArr_CH.length>=2?(uvArr_CH[uvArr_CH.length-1]/Math.max(1,uvArr_CH[uvArr_CH.length-2])-1)*100:0;

// Conversion rate monthly
var cvrMo_C=[],cvrMo_CH=[];
MO.forEach(function(m){
  var uvc=D.uv_monthly["卡西欧"]&&D.uv_monthly["卡西欧"][m]?D.uv_monthly["卡西欧"][m].uv:0;
  var ordc=D.uv_monthly["卡西欧"]&&D.uv_monthly["卡西欧"][m]?D.uv_monthly["卡西欧"][m].orders:0;
  cvrMo_C.push(uvc>0?ordc/uvc*100:0);
  var uvch=D.uv_monthly["蔻驰"]&&D.uv_monthly["蔻驰"][m]?D.uv_monthly["蔻驰"][m].uv:0;
  var ordch=D.uv_monthly["蔻驰"]&&D.uv_monthly["蔻驰"][m]?D.uv_monthly["蔻驰"][m].orders:0;
  cvrMo_CH.push(uvch>0?ordch/uvch*100:0);
});

// 得物推 efficiency
var pushC_last=pc.length>0?pc[pc.length-1]:0;
var pushCH_last=pch.length>0?pch[pch.length-1]:0;
var pushC_prev=pc.length>1?pc[pc.length-2]:0;

// Community efficiency
var commC_last=cc2.length>0?cc2[cc2.length-1]:0;

var html='<div class="analysis">';

// ===== 卡西欧 =====
html+='<h3 style="color:#f59e0b;margin-top:0">🟠 卡西欧 — 基本盘诊断</h3>';
html+='<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px 12px;margin-bottom:12px;font-size:13px">';
html+='<div><b>累计交易GMV：</b>¥'+(totalGMV_C/10000).toFixed(1)+'万</div>';
html+='<div><b>累计UV：</b>'+(totalUV_C/10000).toFixed(1)+'万</div>';
html+='<div><b>有效订单：</b>'+totalOrd_C+'单</div>';
html+='<div><b>笔单价：</b>¥'+asp_C.toFixed(0)+'</div>';
html+='<div><b>在售SPU：</b>'+spuCountC+'个</div>';
html+='<div><b>商详转化率：</b>'+cvr_C.toFixed(2)+'%</div>';
html+='<div><b>TOP3集中度：</b>'+concC.toFixed(0)+'%</div>';
html+='<div><b>月环比GMV：</b><span style="color:'+(trendC>=0?GRN:RED)+'">'+(trendC>=0?'+':'')+trendC.toFixed(1)+'%</span></div>';
html+='<div><b>月环比UV：</b><span style="color:'+(uvTrendC>=0?GRN:RED)+'">'+(uvTrendC>=0?'+':'')+uvTrendC.toFixed(1)+'%</span></div>';
html+='<div><b>累计营销费：</b>¥'+(totalMktC/10000).toFixed(1)+'万</div>';
html+='<div><b>综合ROI：</b>'+roiAvgC.toFixed(1)+'x</div>';
html+='<div><b>当月得物推：</b>¥'+pushC_last.toLocaleString()+'</div>';
html+='</div>';

// ===== 蔻驰 =====
html+='<h3 style="color:#a78bfa">🟣 蔻驰 — 高潜增长诊断</h3>';
html+='<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px 12px;margin-bottom:12px;font-size:13px">';
html+='<div><b>累计交易GMV：</b>¥'+(totalGMV_CH/10000).toFixed(1)+'万</div>';
html+='<div><b>累计UV：</b>'+(totalUV_CH/10000).toFixed(1)+'万</div>';
html+='<div><b>有效订单：</b>'+totalOrd_CH+'单</div>';
html+='<div><b>笔单价：</b>¥'+asp_CH.toFixed(0)+'</div>';
html+='<div><b>在售SPU：</b>'+spuCountCH+'个</div>';
html+='<div><b>商详转化率：</b>'+cvr_CH.toFixed(2)+'%</div>';
html+='<div><b>TOP3集中度：</b>'+concCH.toFixed(0)+'%</div>';
html+='<div><b>月环比GMV：</b><span style="color:'+(trendCH>=0?GRN:RED)+'">'+(trendCH>=0?'+':'')+trendCH.toFixed(1)+'%</span></div>';
html+='<div><b>月环比UV：</b><span style="color:'+(uvTrendCH>=0?GRN:RED)+'">'+(uvTrendCH>=0?'+':'')+uvTrendCH.toFixed(1)+'%</span></div>';
html+='<div><b>总GMV占比：</b>'+chPct.toFixed(1)+'%</div>';
html+='<div><b>综合ROI：</b>'+(roiAvgCH>0?roiAvgCH.toFixed(1)+'x':'无投放')+'</div>';
html+='<div><b>当月占比：</b>'+lmChPct.toFixed(1)+'%</div>';
html+='</div>';

// ===== 市场环境 =====
html+='<h3 style="color:#64748b">🌐 市场环境分析</h3>';
html+='<div style="font-size:13px;margin-bottom:12px">';
html+='<div><b>大盘·日韩表：</b>年初 '+mktVals[0].toFixed(0)+' → '+lastM.replace("2026-","")+'月 '+mktVals[mktVals.length-1].toFixed(0)+'，累计变动 <span style="color:'+(mktTrend>=0?GRN:RED)+'">'+(mktTrend>=0?'+':'')+mktTrend.toFixed(1)+'%</span></div>';

var mktBagVals=D.market_monthly_bag_avg?Object.values(D.market_monthly_bag_avg):[];
if(mktBagVals.length>=2){
var mktBagTrend=((mktBagVals[mktBagVals.length-1]/mktBagVals[0])-1)*100;
html+='<div style="margin-top:4px"><b>大盘·单肩包：</b>年初 '+mktBagVals[0].toFixed(0)+' → '+lastM.replace("2026-","")+'月 '+mktBagVals[mktBagVals.length-1].toFixed(0)+'，累计变动 <span style="color:'+(mktBagTrend>=0?GRN:RED)+'">'+(mktBagTrend>=0?'+':'')+mktBagTrend.toFixed(1)+'%</span></div>';
}
html+='<div style="margin-top:4px"><b>品类竞争格局：</b>卡西欧占主导('+((100-chPct).toFixed(1))+'%)，蔻驰占比'+chPct.toFixed(1)+'%。'+lastM.replace("2026-","")+'月蔻驰占比'+lmChPct.toFixed(1)+'%</div>';

// Monthly conversion trend
html+='<div style="margin-top:6px"><b>月度转化率趋势：</b></div>';
html+='<div style="display:flex;gap:16px;margin-top:4px">';
html+='<div><span style="color:#f59e0b">卡西欧：</span>'+MO.map(function(m,i){return m.replace("2026-","")+'月 '+cvrMo_C[i].toFixed(2)+'%';}).join(" → ")+'</div>';
html+='<div><span style="color:#a78bfa">蔻驰：</span>'+MO.map(function(m,i){return m.replace("2026-","")+'月 '+cvrMo_CH[i].toFixed(2)+'%';}).join(" → ")+'</div>';
html+='</div>';
html+='</div>';

// ===== 详细建议 =====
html+='<h3 style="color:#22c55e">⚡ 运营建议</h3>';
html+='<table class="st" style="font-size:13px">';
html+='<thead><tr><th style="width:40px;text-align:center">级别</th><th style="width:80px">方向</th><th>建议详情</th></tr></thead>';

// P0 - 紧急
// 卡西欧 GMV 月环比大幅下滑
if(trendC<-10){
html+='<tr><td style="color:#ef4444;font-weight:600;text-align:center">P0</td><td>GMV下滑</td><td><b>卡西欧'+lastM.replace("2026-","")+'月GMV环比'+trendC.toFixed(0)+'%</b>，从¥'+(pmC/10000).toFixed(1)+'万降至¥'+(lmC/10000).toFixed(1)+'万。紧急排查：①检查TOP5款是否被竞品截流（对比价格、主图、评价）；②核查得物推送计划是否停投或降权；③确认有无售后/违规问题导致降权；④查看是否因大盘季节性回落（日韩表指数'+mktVals[mktVals.length-1].toFixed(0)+'，上月'+mktVals[mktVals.length-2].toFixed(0)+'）。</td></tr>';
} else if(trendC<-5){
html+='<tr><td style="color:#f59e0b;font-weight:600;text-align:center">P1</td><td>轻微下滑</td><td><b>卡西欧'+lastM.replace("2026-","")+'月GMV环比'+trendC.toFixed(0)+'%</b>，小幅下滑需关注。建议：①检查TOP10款UV变化，定位下滑源头；②对比竞品同价位段新品上架情况；③考虑小额度得物推测款（¥500-1000）刺激流量。</td></tr>';
}

// 蔻驰高客单价机会
if(asp_CH>asp_C*1.5){
html+='<tr><td style="color:'+GRN+';font-weight:600;text-align:center">P0</td><td>高客单</td><td><b>蔻驰笔单价¥'+asp_CH.toFixed(0)+'（'+((asp_CH/asp_C).toFixed(1))+'x卡西欧）</b>，高客单价=高利润空间。建议：①扩充蔻驰SPU至15-20个，补全¥800-2000价格带（目前仅'+spuCountCH+'个SPU）；②针对高客单意向用户开启精准得物推，出价可适度提高；③优化高客单款详情页视频和评价质量，提升转化率；④考虑推出限定/联名款提升品牌溢价。</td></tr>';
}

// 蔻驰增长空间
if(chPct<20 && totalGMV_CH>20000){
html+='<tr><td style="color:'+GRN+';font-weight:600;text-align:center">P0</td><td>增长空间</td><td><b>蔻驰仅占'+chPct.toFixed(1)+'%GMV</b>，'+lastM.replace('2026-','')+'月提升至'+lmChPct.toFixed(1)+'%。增长路径：①启动社区投放（参照美兰经验，首月预算¥5,000-8,000，选择穿搭/时尚类达人）；②为蔻驰独立开设得物推计划，月预算¥2,000-5,000，优先推TOP3款；③每周上新2-3款，丰富产品矩阵；④关注单肩包大盘动向（当前'+lastM.replace('2026-','')+'月指数'+(mktBagVals.length>0?mktBagVals[mktBagVals.length-1].toFixed(0):'—')+'），把握品类红利。</td></tr>';
}

// 卡西欧头部集中风险
if(concC>60){
html+='<tr><td style="color:#f59e0b;font-weight:600;text-align:center">P1</td><td>头部风险</td><td><b>TOP3占'+concC.toFixed(0)+'%GMV</b>，单点故障风险高。建议：①筛选第4-10名潜力款（GMV>¥5,000），加大详情页优化+得物推曝光测试；②优化腰部商品标题/主图/评价；③对TOP1款（占比最高）建立竞品监控机制，每周比对价格和评价。</td></tr>';
}

// 转化率对比
if(cvr_C<cvr_CH){
html+='<tr><td style="color:#f59e0b;font-weight:600;text-align:center">P1</td><td>转化优化</td><td><b>卡西欧商详转化率'+cvr_C.toFixed(2)+'%低于蔻驰('+cvr_CH.toFixed(2)+'%)</b>。虽然品类不同，但可从以下维度优化：①TOP10款详情页视频质量（对比蔻驰详情页结构取长补短）；②增加买家秀和问答区维护频率；③优化价格竞争力（关注平台最低价提醒）；④检查主图首图吸引力，考虑A/B测试。</td></tr>';
}

// 营销效率
if(totalMktC>0 && roiAvgC<3){
html+='<tr><td style="color:#f59e0b;font-weight:600;text-align:center">P1</td><td>营销效率</td><td><b>综合ROI仅'+roiAvgC.toFixed(1)+'x</b>，获客成本¥'+(cacC.length>0?cacC[cacC.length-1].toFixed(0):'—')+'/单。建议：①按SPU拆分得物推数据，砍掉ROI<1的无效款，资金集中到ROI>3的头部款；②社区投放优化选品策略，优先投TOP10高转化款；③'+lastM.replace('2026-','')+'月得物推¥'+pushC_last.toLocaleString()+'，社区¥'+commC_last.toLocaleString()+'，评估两个渠道ROI差异后调整预算配比。</td></tr>';
}

// 蔻驰付费投放建议
if(totalPushCH===0 && totalGMV_CH>20000){
html+='<tr><td style="color:#f59e0b;font-weight:600;text-align:center">P1</td><td>蔻驰投放</td><td><b>GMV已超¥2万但零付费投放</b>。启动方案：①选3-5款TOP商品做精准得物推，小预算测试（¥2,000/月起步）；②投放后跟踪7日ROI、加购率、收藏率；③投产比>2的款加码预算，<1的及时止损换款；④'+lastM.replace('2026-','')+'月底前完成首批测试，'+MO[MO.length-1].replace('2026-','')+'月中复盘。</td></tr>';
}

// UV下滑但GMV稳定 -> 转化率提升信号
if(uvTrendC<-10 && trendC>-5){
html+='<tr><td style="color:#f59e0b;font-weight:600;text-align:center">P1</td><td>流量预警</td><td><b>卡西欧UV环比'+uvTrendC.toFixed(0)+'%但GMV相对稳定</b>，说明转化率在提升（流失的是低质流量），但长期看流量下滑会限制增长。建议：①排查自然搜索排名是否下降；②检查是否因竞品新增投放抢占流量；③适度增加得物推预算补充流量；④优化商品标题关键词覆盖。</td></tr>';
}

// 大盘向好
if(mktTrend>10){
html+='<tr><td style="color:#22c55e;font-weight:600;text-align:center">P2</td><td>大盘红利</td><td><b>日韩表大盘累计增长'+mktTrend.toFixed(0)+'%</b>，品类处于上升期。建议：①适度加码整体预算10-20%抢占流量窗口；②扩充日韩表SPU数量（新品+配件），丰富价格带覆盖；③关注「得物趋势」板块，挖掘上升中的细分风格。</td></tr>';
}

// 常规监控
html+='<tr><td style="color:#f59e0b;font-weight:600;text-align:center">P2</td><td>常规监控</td><td><b>每周更新看板</b>，重点盯：①UV增长率（日环比>15%需排查原因）；②GMV月环比（连续2月下滑>5%启动专项诊断）；③单品UV异常（突然下降>30%排查竞品/价格）；④获客成本趋势（连续2月上升>20%调整投放策略）；⑤蔻驰占比月度变化（目标逐步提升至25-30%）。</td></tr>';

// 月度具体行动建议
html+='<tr><td style="color:#22c55e;font-weight:600;text-align:center">P2</td><td>月度行动</td><td><b>'+lastM.replace('2026-','')+'月重点行动：</b>①查看模块四异动款式，对异常下滑款重点排查（价格/竞品/评价）；②对比模块三TOP20三个维度，找出"高UV低转化"款优化详情页，找出"高转化低UV"款增加曝光投放；③对比模块七ROI表格，低效月份复盘原因，高效月份分析可复制经验；④更新货盘表确保新品已入库。</td></tr>';

html+='</table></div>';
el.innerHTML=html;
}

// ================================================================
// 单品弹窗 (不变)
// ================================================================
document.addEventListener("click",function(e){var el=e.target.closest(".slnk");if(!el)return;var b=el.dataset.brand,s=el.dataset.spu;if(b&&s)openMd(b,s);});

var curB="",curS="";
function openMd(brand,spu){curB=brand;curS=spu;
// Fuzzy match: if exact lookup fails, try prefix match
if(!(D.uv_by_spu[brand]&&D.uv_by_spu[brand][spu])){
  var pool=D.uv_by_spu[brand]||{},keys=Object.keys(pool);
  for(var i=0;i<keys.length;i++){if(keys[i].indexOf(spu)===0||spu.indexOf(keys[i])===0){curS=keys[i];break;}}
  if(!pool[curS]){for(var i=0;i<keys.length;i++){if(keys[i].indexOf(spu.replace(/&quot;/g,'"'))>=0){curS=keys[i];break;}}}
}
document.getElementById("md-t").textContent=curS;document.getElementById("md-ds").value=DR.start;document.getElementById("md-de").value=DR.end;document.getElementById("md").style.display="flex";renderSpu();}
function closeMd(){document.getElementById("md").style.display="none";if(spC){spC.destroy();spC=null;}}
function renderSpu(){
var sd2=D.uv_by_spu[curB]?D.uv_by_spu[curB][curS]:null;if(!sd2){document.getElementById("stb").innerHTML='<tr><td colspan="4" style="color:#94a3b8;text-align:center;padding:20px">该款式暂无商详访客数据</td></tr>';return;}
var ds=document.getElementById("md-ds").value||DR.start,de=document.getElementById("md-de").value||DR.end;
var daily=(sd2.daily||[]).filter(function(d){return d.date>=ds&&d.date<=de});
if(daily.length>60){var wk={};daily.forEach(function(d){var dt=new Date(d.date),ws=new Date(dt);ws.setDate(dt.getDate()-dt.getDay()+1);var k=ws.toISOString().slice(0,10);if(!wk[k])wk[k]={uv:0,gmv:0,orders:0};wk[k].uv+=d.uv;wk[k].gmv+=d.gmv;wk[k].orders+=d.orders;});daily=Object.entries(wk).sort().map(function(e){return{date:e[0],uv:e[1].uv,gmv:e[1].gmv,orders:e[1].orders}});}
var lbs=daily.map(function(d){return d.date.slice(5)}),uv=daily.map(function(d){return d.uv}),gmv=daily.map(function(d){return d.gmv});
var tbl=document.getElementById("stb");tbl.innerHTML="";
var mo2=sd2.monthly||{};
Object.keys(mo2).sort().forEach(function(m){var d=mo2[m],tr=document.createElement("tr");tr.innerHTML="<td>"+m+"</td><td style='text-align:right'>"+d.uv.toLocaleString()+"</td><td style='text-align:right'>"+d.orders+"</td><td style='text-align:right'>¥"+d.gmv.toLocaleString()+"</td>";tbl.appendChild(tr);});
if(spC)spC.destroy();
spC=new Chart(document.getElementById("sc"),{type:"bar",data:{labels:lbs,datasets:[
{label:"UV",data:uv,borderColor:"#60a5fa",backgroundColor:"transparent",yAxisID:"y",tension:0.3,borderWidth:2,pointRadius:0,type:"line",order:1},
{label:"GMV(元)",data:gmv,backgroundColor:CHC+"70",borderColor:CHC,borderWidth:1,yAxisID:"y1",order:2}
]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{position:"left",ticks:{color:TX},grid:{color:GG}},y1:{position:"right",ticks:{color:TX,callback:function(v){return"¥"+(v/1000).toFixed(0)+"k"}},grid:{drawOnChartArea:false}},x:{ticks:{color:TX,maxTicksLimit:15},grid:{color:GG}}},plugins:{legend:{display:false}}}});
}
document.addEventListener("keydown",function(e){if(e.key==="Escape")closeMd();});
document.getElementById("md").addEventListener("click",function(e){if(e.target===document.getElementById("md"))closeMd();});
