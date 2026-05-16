// ==================== 喜过运营分析看板 - 核心脚本 ====================
var MO=D.months,ML=MO.map(function(m){return m.replace("2026-","")+"月"});
var CC="#f59e0b",CHC="#a78bfa",TX="#94a3b8",GG="#1e2235",GRN="#22c55e",RED="#ef4444";

function gm(b,k){var m=D.orders_monthly[b]||{},r=[];MO.forEach(function(x){r.push(m[x]?m[x][k]||0:0)});return r;}
function uvm(b,k){var m=D.uv_monthly[b]||{},r=[];MO.forEach(function(x){r.push(m[x]?m[x][k]||0:0)});return r;}

window.addEventListener("load",function(){
// ===== KPI =====
var cg=gm("卡西欧","gmv"),co=gm("卡西欧","orders"),chg=gm("蔻驰","gmv"),cho=gm("蔻驰","orders");
var cguv=uvm("卡西欧","uv"),chuv=uvm("蔻驰","uv");
document.getElementById("kv-cgmv").textContent="¥"+(cg.reduce(function(a,b){return a+b},0)/10000).toFixed(1)+"万";
document.getElementById("kv-co").textContent=co.reduce(function(a,b){return a+b},0);
document.getElementById("kv-chgmv").textContent="¥"+(chg.reduce(function(a,b){return a+b},0)/10000).toFixed(1)+"万";
document.getElementById("kv-cho").textContent=cho.reduce(function(a,b){return a+b},0);

// ===== 模块一: UV曲线 + GMV柱状图 + 大盘指数 =====
var mktAvg=MO.map(function(m){return D.market_monthly_avg?D.market_monthly_avg[m]||null:null});
var gmvC=MO.map(function(m){var v=D.orders_monthly["卡西欧"]&&D.orders_monthly["卡西欧"][m]?D.orders_monthly["卡西欧"][m].gmv:0;return v/10000;});
var gmvCH=MO.map(function(m){var v=D.orders_monthly["蔻驰"]&&D.orders_monthly["蔻驰"][m]?D.orders_monthly["蔻驰"][m].gmv:0;return v/10000;});

new Chart(document.getElementById("c1"),{type:"bar",data:{labels:ML,datasets:[
// UV 曲线
{label:"卡西欧 UV",data:uvm("卡西欧","uv"),borderColor:CC,backgroundColor:"transparent",yAxisID:"y",tension:0.3,borderWidth:2.5,pointRadius:3,pointBackgroundColor:CC,type:"line",order:1},
{label:"蔻驰 UV",data:uvm("蔻驰","uv"),borderColor:CHC,backgroundColor:"transparent",yAxisID:"y",tension:0.3,borderWidth:2.5,borderDash:[5,5],pointRadius:3,pointBackgroundColor:CHC,type:"line",order:2},
// GMV 柱状图
{label:"卡西欧 GMV(万)",data:gmvC,backgroundColor:CC+"70",borderColor:CC,borderWidth:1,yAxisID:"y1",order:3},
{label:"蔻驰 GMV(万)",data:gmvCH,backgroundColor:CHC+"70",borderColor:CHC,borderWidth:1,yAxisID:"y1",order:4},
// 大盘指数 曲线
{label:"大盘指数(日韩表)",data:mktAvg,borderColor:"#64748b",backgroundColor:"transparent",yAxisID:"y2",tension:0.3,borderWidth:1.5,pointRadius:3,pointBackgroundColor:"#64748b",type:"line",order:0}
]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:"index",intersect:false},plugins:{legend:{labels:{color:TX,usePointStyle:true,padding:14,font:{size:11}},position:"top"},tooltip:{callbacks:{label:function(ctx){var v=ctx.raw;if(ctx.dataset.label.indexOf("UV")>=0)return ctx.dataset.label+": "+v.toLocaleString();if(ctx.dataset.label.indexOf("大盘")>=0)return ctx.dataset.label+": "+v;return ctx.dataset.label+": ¥"+v.toFixed(1)+"万";}}}},scales:{y:{position:"left",title:{display:true,text:"UV",color:TX},grid:{color:GG},ticks:{color:TX,callback:function(v){return v>=1000?(v/1000).toFixed(0)+"k":v}}},y1:{position:"right",title:{display:true,text:"GMV(万元)",color:TX},grid:{drawOnChartArea:false},ticks:{color:TX,callback:function(v){return"¥"+v.toFixed(1)}},beginAtZero:true},y2:{position:"right",title:{display:true,text:"大盘指数",color:"#64748b"},grid:{drawOnChartArea:false},ticks:{color:"#64748b"},min:0},x:{ticks:{color:TX},grid:{color:GG}}}}});


// ===== 模块二: 货号销售排行 =====
function bt(brand,tid){
var spus=D.orders_by_spu[brand]||[],total=0;spus.forEach(function(s){total+=s.gmv});
var tb=document.getElementById(tid);
spus.forEach(function(s){
var pct=total>0?(s.gmv/total*100).toFixed(1):"0.0",nm=s.spu.length>40?s.spu.slice(0,40)+"...":s.spu;
var tr=document.createElement("tr"),spuEnc=s.spu.replace(/"/g,"&quot;");
tr.innerHTML='<td><span class="slnk" data-brand="'+brand+'" data-spu="'+spuEnc+'">'+nm+'</span></td><td style="text-align:right">¥'+s.gmv.toLocaleString()+'</td><td style="text-align:right">'+pct+'%</td>';
tb.appendChild(tr);
});}
bt("卡西欧","tb-c");bt("蔻驰","tb-ch");

// ===== 模块三: 月度品牌销售占比（堆叠柱状图）=====
var sd=MO.map(function(m){
var c=0,ch=0;
if(D.orders_monthly["卡西欧"]&&D.orders_monthly["卡西欧"][m])c=D.orders_monthly["卡西欧"][m].gmv||0;
if(D.orders_monthly["蔻驰"]&&D.orders_monthly["蔻驰"][m])ch=D.orders_monthly["蔻驰"][m].gmv||0;
return {c:c,ch:ch,total:c+ch};
});
new Chart(document.getElementById("c3"),{type:"bar",data:{labels:ML,datasets:[
{label:"卡西欧",data:sd.map(function(d){return d.c}),backgroundColor:CC+"80",borderColor:CC,borderWidth:1},
{label:"蔻驰",data:sd.map(function(d){return d.ch}),backgroundColor:CHC+"80",borderColor:CHC,borderWidth:1}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:TX,usePointStyle:true}},tooltip:{callbacks:{label:function(ctx){var v=ctx.raw,t=sd[ctx.dataIndex].total,p=t>0?(v/t*100).toFixed(1):"0";return ctx.dataset.label+": \u00A5"+v.toLocaleString()+" ("+p+"%)";}}}},scales:{x:{stacked:true,ticks:{color:TX},grid:{color:GG}},y:{stacked:true,ticks:{color:TX,callback:function(v){return"\u00A5"+(v/10000).toFixed(0)+"\u4E07";}},grid:{color:GG}}}}});

// ===== 模块四: 营销效率 =====
var pc=[],pch=[],cc2=[];
MO.forEach(function(m,i){pc.push(D.push_monthly["卡西欧"]?D.push_monthly["卡西欧"][m]||0:0);pch.push(D.push_monthly["蔻驰"]?D.push_monthly["蔻驰"][m]||0:0);cc2.push(D.comm_monthly["卡西欧"]&&D.comm_monthly["卡西欧"][m]?D.comm_monthly["卡西欧"][m].cost||0:0)});
var mc=[];MO.forEach(function(m,i){mc.push(pc[i]+cc2[i])});
var mch=pch,cacC=[],cacCh=[],roiC=[],roiCh=[];
MO.forEach(function(m,i){cacC.push(co[i]>0?mc[i]/co[i]:0);cacCh.push(cho[i]>0?mch[i]/cho[i]:0);roiC.push(mc[i]>0?cg[i]/mc[i]:0);roiCh.push(mch[i]>0?chg[i]/mch[i]:0)});

new Chart(document.getElementById("c4a"),{type:"bar",data:{labels:ML,datasets:[
{label:"卡西欧",data:cacC,backgroundColor:CC+"60",borderColor:CC,borderWidth:1},
{label:"蔻驰",data:cacCh,backgroundColor:CHC+"60",borderColor:CHC,borderWidth:1}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:TX,usePointStyle:true,font:{size:10}}},title:{display:true,text:"获客成本(¥/单)",color:TX,font:{size:13}}},scales:{y:{ticks:{color:TX,callback:function(v){return"¥"+v.toFixed(0)}},grid:{color:GG}},x:{ticks:{color:TX},grid:{color:GG}}}}});

new Chart(document.getElementById("c4b"),{type:"bar",data:{labels:ML,datasets:[
{label:"卡西欧",data:roiC,backgroundColor:CC+"60",borderColor:CC,borderWidth:1},
{label:"蔻驰",data:roiCh,backgroundColor:CHC+"60",borderColor:CHC,borderWidth:1}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:TX,usePointStyle:true,font:{size:10}}},title:{display:true,text:"ROI (GMV/营销费)",color:TX,font:{size:13}}},scales:{y:{ticks:{color:TX,callback:function(v){return v.toFixed(1)+"x"}},grid:{color:GG}},x:{ticks:{color:TX},grid:{color:GG}}}}});

var mt=document.getElementById("tb-m");
MO.forEach(function(m,i){var tr=document.createElement("tr"),cr=roiC[i],chr=roiCh[i],t1=cr>3?"tg":cr>1?"ty":"tr",t2=chr>3?"tg":chr>1?"ty":"tr";
tr.innerHTML="<td>"+m.replace("2026-","")+"月</td><td>¥"+mc[i].toLocaleString()+"</td><td>"+(cacC[i]>0?"¥"+cacC[i].toFixed(0):"—")+"</td><td><span class='tag "+t1+"'>"+(cr>0?cr.toFixed(1)+"x":"—")+"</span></td><td>¥"+mch[i].toLocaleString()+"</td><td>"+(cacCh[i]>0?"¥"+cacCh[i].toFixed(0):"—")+"</td><td><span class='tag "+t2+"'>"+(chr>0?chr.toFixed(1)+"x":"—")+"</span></td>";
mt.appendChild(tr);});

// ===== 模块五: 动态运营诊断 =====
renderAnalysis();

}); // end window.load

// ==================== 模块二: 单品弹窗 ====================
document.addEventListener("click",function(e){var el=e.target.closest(".slnk");if(!el)return;var b=el.dataset.brand,s=el.dataset.spu;if(b&&s)openMd(b,s)});

var curB="",curS="",spC=null;
function openMd(brand,spu){curB=brand;curS=spu;document.getElementById("md-t").textContent=spu;document.getElementById("md-ds").value="2026-01-01";document.getElementById("md-de").value="2026-05-14";document.getElementById("md").style.display="block";renderSpu();}
function closeMd(){document.getElementById("md").style.display="none";if(spC){spC.destroy();spC=null;}}
function renderSpu(){
var sd2=D.uv_by_spu[curB]?D.uv_by_spu[curB][curS]:null;if(!sd2)return;
var ds=document.getElementById("md-ds").value||"2026-01-01",de=document.getElementById("md-de").value||"2026-05-14";
var daily=(sd2.daily||[]).filter(function(d){return d.date>=ds&&d.date<=de});
if(daily.length>60){var wk={};daily.forEach(function(d){var dt=new Date(d.date),ws=new Date(dt);ws.setDate(dt.getDate()-dt.getDay()+1);var k=ws.toISOString().slice(0,10);if(!wk[k])wk[k]={uv:0,gmv:0,orders:0};wk[k].uv+=d.uv;wk[k].gmv+=d.gmv;wk[k].orders+=d.orders;});daily=Object.entries(wk).sort().map(function(e){return{date:e[0],uv:e[1].uv,gmv:e[1].gmv,orders:e[1].orders}});}
var lbs=daily.map(function(d){return d.date.slice(5)}),uv=daily.map(function(d){return d.uv}),gmv=daily.map(function(d){return d.gmv});
var tbl=document.getElementById("stb");tbl.innerHTML="";
var mo2=sd2.monthly||{};
Object.keys(mo2).sort().forEach(function(m){var d=mo2[m],tr=document.createElement("tr");tr.innerHTML="<td>"+m+"</td><td style='text-align:right'>"+d.uv.toLocaleString()+"</td><td style='text-align:right'>"+d.orders+"</td><td style='text-align:right'>¥"+d.gmv.toLocaleString()+"</td>";tbl.appendChild(tr);});
if(spC)spC.destroy();
spC=new Chart(document.getElementById("sc"),{type:"line",data:{labels:lbs,datasets:[
{label:"UV",data:uv,borderColor:"#60a5fa",yAxisID:"y",tension:0.3,borderWidth:2,pointRadius:0},
{label:"GMV(元)",data:gmv,borderColor:CHC,yAxisID:"y1",tension:0.3,borderWidth:2,pointRadius:0}
]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{position:"left",ticks:{color:TX},grid:{color:GG}},y1:{position:"right",ticks:{color:TX,callback:function(v){return"¥"+(v/1000).toFixed(0)+"k"}},grid:{drawOnChartArea:false}},x:{ticks:{color:TX,maxTicksLimit:15},grid:{color:GG}}},plugins:{legend:{display:false}}}});
}
document.addEventListener("keydown",function(e){if(e.key==="Escape")closeMd()});
document.getElementById("md").addEventListener("click",function(e){if(e.target===document.getElementById("md"))closeMd()});

// ==================== 模块五: 动态运营诊断分析 ====================
function renderAnalysis(){
var el=document.getElementById("mod5");if(!el)return;

// 汇总数据计算
var totalGMV_C=gm("卡西欧","gmv").reduce(function(a,b){return a+b},0);
var totalGMV_CH=gm("蔻驰","gmv").reduce(function(a,b){return a+b},0);
var totalOrd_C=gm("卡西欧","orders").reduce(function(a,b){return a+b},0);
var totalOrd_CH=gm("蔻驰","orders").reduce(function(a,b){return a+b},0);
var asp_C=totalOrd_C>0?totalGMV_C/totalOrd_C:0;
var asp_CH=totalOrd_CH>0?totalGMV_CH/totalOrd_CH:0;
var totalGMV=totalGMV_C+totalGMV_CH;

// UV数据
var totalUV_C=uvm("卡西欧","uv").reduce(function(a,b){return a+b},0);
var totalUV_CH=uvm("蔻驰","uv").reduce(function(a,b){return a+b},0);
var uvGMV_C=uvm("卡西欧","gmv").reduce(function(a,b){return a+b},0);
var uvGMV_CH=uvm("蔻驰","gmv").reduce(function(a,b){return a+b},0);
var uvOrd_C=uvm("卡西欧","orders").reduce(function(a,b){return a+b},0);
var uvOrd_CH=uvm("蔻驰","orders").reduce(function(a,b){return a+b},0);

// 转化率
var cvr_C=totalUV_C>0?(uvOrd_C/totalUV_C*100):0;
var cvr_CH=totalUV_CH>0?(uvOrd_CH/totalUV_CH*100):0;

// 营销总费用
var totalPushC=pc.reduce(function(a,b){return a+b},0);
var totalPushCH=pch.reduce(function(a,b){return a+b},0);
var totalCommC=cc2.reduce(function(a,b){return a+b},0);
var totalMktC=totalPushC+totalCommC;
var totalMktCH=totalPushCH;
var roiAvgC=totalMktC>0?totalGMV_C/totalMktC:0;
var roiAvgCH=totalMktCH>0?totalGMV_CH/totalMktCH:0;

// 月度趋势判断
var gmvArr_C=gm("卡西欧","gmv");
var trendC=gmvArr_C.length>=2?(gmvArr_C[gmvArr_C.length-1]/Math.max(1,gmvArr_C[gmvArr_C.length-2])-1)*100:0;
var gmvArr_CH=gm("蔻驰","gmv");
var trendCH=gmvArr_CH.length>=2?(gmvArr_CH[gmvArr_CH.length-1]/Math.max(1,gmvArr_CH[gmvArr_CH.length-2])-1)*100:0;

// SPU数量
var spuCountC=(D.orders_by_spu["卡西欧"]||[]).length;
var spuCountCH=(D.orders_by_spu["蔻驰"]||[]).length;

// TOP3 SPU集中度
var top3GMV_C=0,top3GMV_CH=0;
(D.orders_by_spu["卡西欧"]||[]).slice(0,3).forEach(function(s){top3GMV_C+=s.gmv});
(D.orders_by_spu["蔻驰"]||[]).slice(0,3).forEach(function(s){top3GMV_CH+=s.gmv});
var concC=totalGMV_C>0?(top3GMV_C/totalGMV_C*100):0;
var concCH=totalGMV_CH>0?(top3GMV_CH/totalGMV_CH*100):0;

// 大盘趋势
var mktVals=D.market_monthly_avg?Object.values(D.market_monthly_avg):[];
var mktTrend=mktVals.length>=2?((mktVals[mktVals.length-1]/mktVals[0])-1)*100:0;

// 蔻驰占比
var chPct=totalGMV>0?(totalGMV_CH/totalGMV*100):0;

// 最近一月数据
var lastM=MO[MO.length-1];
var lmC=D.orders_monthly["卡西欧"]&&D.orders_monthly["卡西欧"][lastM]?D.orders_monthly["卡西欧"][lastM].gmv:0;
var lmCH=D.orders_monthly["蔻驰"]&&D.orders_monthly["蔻驰"][lastM]?D.orders_monthly["蔻驰"][lastM].gmv:0;
var lmTotal=lmC+lmCH;
var lmChPct=lmTotal>0?(lmCH/lmTotal*100):0;

// 获客成本趋势
var cacLastC=cacC.length>0?cacC[cacC.length-1]:0;
var cacPrevC=cacC.length>1?cacC[cacC.length-2]:0;

var html='<div class="analysis">';

// 卡西欧
html+='<h3 style="color:#f59e0b;margin-top:0">🟠 卡西欧 — 基本盘</h3>';
html+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px;font-size:13px">';
html+='<div><b>累计交易GMV：</b>¥'+(totalGMV_C/10000).toFixed(1)+'万</div>';
html+='<div><b>累计UV：</b>'+(totalUV_C/10000).toFixed(1)+'万</div>';
html+='<div><b>有效订单：</b>'+totalOrd_C+'单</div>';
html+='<div><b>笔单价：</b>¥'+asp_C.toFixed(0)+'</div>';
html+='<div><b>在售SPU：</b>'+spuCountC+'个</div>';
html+='<div><b>商详转化率：</b>'+cvr_C.toFixed(2)+'%</div>';
html+='<div><b>TOP3集中度：</b>'+concC.toFixed(0)+'%</div>';
html+='<div><b>月环比趋势：</b><span style="color:'+(trendC>=0?GRN:RED)+'">'+(trendC>=0?'+':'')+trendC.toFixed(1)+'%</span></div>';
html+='<div><b>累计营销费：</b>¥'+(totalMktC/10000).toFixed(1)+'万</div>';
html+='<div><b>综合ROI：</b>'+roiAvgC.toFixed(1)+'x</div>';
html+='</div>';

// 蔻驰
html+='<h3 style="color:#a78bfa">🟣 蔻驰 — 高潜增长</h3>';
html+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px;font-size:13px">';
html+='<div><b>累计交易GMV：</b>¥'+(totalGMV_CH/10000).toFixed(1)+'万</div>';
html+='<div><b>累计UV：</b>'+(totalUV_CH/10000).toFixed(1)+'万</div>';
html+='<div><b>有效订单：</b>'+totalOrd_CH+'单</div>';
html+='<div><b>笔单价：</b>¥'+asp_CH.toFixed(0)+'</div>';
html+='<div><b>在售SPU：</b>'+spuCountCH+'个</div>';
html+='<div><b>商详转化率：</b>'+cvr_CH.toFixed(2)+'%</div>';
html+='<div><b>TOP3集中度：</b>'+concCH.toFixed(0)+'%</div>';
html+='<div><b>月环比趋势：</b><span style="color:'+(trendCH>=0?GRN:RED)+'">'+(trendCH>=0?'+':'')+trendCH.toFixed(1)+'%</span></div>';
html+='<div><b>总GMV占比：</b>'+chPct.toFixed(1)+'%</div>';
html+='<div><b>综合ROI：</b>'+(roiAvgCH>0?roiAvgCH.toFixed(1)+'x':'无投放')+'</div>';
html+='</div>';

// 市场环境
html+='<h3 style="color:#64748b">🌐 市场环境</h3>';
html+='<div style="font-size:13px;margin-bottom:12px">';
html+='<div><b>大盘趋势（日韩表）：</b>年初指数 '+mktVals[0].toFixed(0)+' → '+lastM.replace("2026-","")+'月 '+mktVals[mktVals.length-1].toFixed(0)+'，累计变动 <span style="color:'+(mktTrend>=0?GRN:RED)+'">'+(mktTrend>=0?'+':'')+mktTrend.toFixed(1)+'%</span></div>';
html+='<div style="margin-top:4px"><b>蔻驰占比：</b>整体 '+chPct.toFixed(1)+'% | '+lastM.replace("2026-","")+'月 '+lmChPct.toFixed(1)+'%</div>';
html+='</div>';

// 核心建议
html+='<h3 style="color:#22c55e">⚡ 核心建议</h3>';
html+='<table class="st" style="font-size:13px">';

// P0
if(roiAvgC>2 && trendC<-5){
html+='<tr><td style="color:#ef4444;font-weight:600;width:40px">P0</td><td><b>卡西欧下滑预警：</b>GMV月环比'+trendC.toFixed(0)+'%，检查TOP款是否被竞品截流，重点加固TOP3商品（占'+concC.toFixed(0)+'%GMV）的详情页和评价</td></tr>';
}
if(asp_CH>asp_C*2){
html+='<tr><td style="color:'+GRN+';font-weight:600">P0</td><td><b>蔻驰高客单价机会：</b>笔单价¥'+asp_CH.toFixed(0)+'（'+((asp_CH/asp_C).toFixed(1))+'x卡西欧），建议增加蔻驰SPU至20+个，补齐中高端价格带</td></tr>';
}
if(chPct<15 && totalGMV_CH>30000){
html+='<tr><td style="color:'+GRN+';font-weight:600">P0</td><td><b>蔻驰增长空间大：</b>当前仅占'+chPct.toFixed(1)+'%，'+lastM.replace('2026-','')+'月提升至'+lmChPct.toFixed(1)+'%。建议启动社区投放+得物推，参照美兰经验</td></tr>';
}

// P1
if(concC>60){
html+='<tr><td style="color:#f59e0b;font-weight:600">P1</td><td><b>卡西欧头部风险：</b>TOP3占'+concC.toFixed(0)+'%GMV，过度集中。建议筛选4-8名潜力款（GMV>¥5,000），加大曝光测试</td></tr>';
}
if(cvr_C<cvr_CH){
html+='<tr><td style="color:#f59e0b;font-weight:600">P1</td><td><b>卡西欧转化优化：</b>转化率'+cvr_C.toFixed(2)+'%低于蔻驰('+cvr_CH.toFixed(2)+'%)，建议优化详情页视频、增加买家秀</td></tr>';
}
if(totalMktC>0 && roiAvgC<3){
html+='<tr><td style="color:#f59e0b;font-weight:600">P1</td><td><b>营销效率优化：</b>综合ROI仅'+roiAvgC.toFixed(1)+'x，建议按SPU拆分得物推数据，砍掉ROI<1的款，资金集中到ROI>3的头部款</td></tr>';
}
if(totalPushCH===0 && totalGMV_CH>30000){
html+='<tr><td style="color:#f59e0b;font-weight:600">P1</td><td><b>蔻驰付费测试：</b>GMV已超¥3万但零投放。建议'+lastM.replace('2026-','')+'月启动3-5款得物推新品测款，预算¥2,000-5,000/月</td></tr>';
}

// P2
if(mktTrend>10){
html+='<tr><td style="color:#f59e0b;font-weight:600">P2</td><td><b>大盘向好：</b>日韩表大盘指数累计增长'+mktTrend.toFixed(0)+'%，是品类红利期。建议适度加码整体预算</td></tr>';
}
html+='<tr><td style="color:#f59e0b;font-weight:600">P2</td><td><b>数据监控：</b>每周更新一次看板，重点盯UV增长率和GMV月环比。当单品UV突然下降>30%时排查竞品投放/价格变化</td></tr>';

html+='</table></div>';

el.innerHTML=html;
}
