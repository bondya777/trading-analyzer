// Synthetic fixtures only. Run: node tests/account-narrative.test.cjs
const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const html=fs.readFileSync(path.join(__dirname,'..','交易数据分析平台.html'),'utf8');
const moduleSource=html.match(/\/\/ ACCOUNT_NARRATIVE_BEGIN([\s\S]*?)\/\/ ACCOUNT_NARRATIVE_END/)[1];
const context=vm.createContext({Date});
new vm.Script(moduleSource).runInContext(context);
const A=vm.runInContext('AccountNarrative',context);
let nextId=0;
const account='DEMO_ACCOUNT';
function flow(code,date,direction,yi,category='信用债',variety='现券买卖',qty=yi*1e4,time=''){
  return{_code:account,_name:'示例账户',_class:category,_variety:variety,_biz:variety,_dir:direction,_dateStr:date,_amountYi:yi,_amount:yi*1e8,_mkt:'上交所',代码:code,名称:'示例证券'+code,交易编号:String(++nextId),'交易面额(万)':qty,成交时间:time,_source:{file:'synthetic.xlsx',sheet:'流水',row:nextId+2}};
}
function bond(code,yi,type='企业债',remain=2,rate='AA'){
  return{acct:account,code,name:'示例债券'+code,market:'上交所',bondType:type,assetType:'债券',mv:yi*1e4,face:yi*1e4,remain,rating:rate,_source:{file:'synthetic-holdings.xlsx',sheet:'证券余额06.30',row:++nextId}};
}
function fund(code,yi,name,kind='基金'){
  return{acct:account,code,name,market:'上交所',kind,mv:yi*1e4,face:yi*1e4,_source:{file:'synthetic-holdings.xlsx',sheet:'基金余额06.30',row:++nextId}};
}
function store(flows=[],end={},start={},extra={}){
  return{zqAll:flows,zq:flows.filter(r=>['现券买卖','基金买卖','国债期货'].includes(r._variety)),hg:[],pos:[],hdStart:{bonds:[],funds:[],futures:[],dateStr:'2026-01-01',...start},hdEnd:{bonds:[],funds:[],futures:[],dateStr:'2026-06-30',...end},...extra};
}
function result(s,dim={},options={}){const idx=A.build(s,dim,options);return{idx,r:A.overview(idx,account)};}

test('cross-month sales retain the earlier buy and do not invent an opening date',()=>{
  const buy=flow('B1','2026-01-31','买入',1),sell=flow('B1','2026-02-01','卖出',1.01,'信用债','现券买卖',10000);
  const {idx}=result(store([buy,sell]));const m=A.monthly(idx,account,'2026-02');const c=m.stats.categories[0];
  assert.equal(c.buyYi,0);assert.equal(c.knownYi,1.01);assert.equal(c.oneTwoYi,1.01);assert.equal(c.unmatchedYi,0);
  assert(!m.text.includes('截至2026年6月'));assert(!m.text.includes('没有买入来源'));
});
test('old opening inventory remains undated and is not described as a one-day holding',()=>{
  const {idx}=result(store([flow('B1','2026-01-02','卖出',1)],{}, {bonds:[bond('B1',1)]}));
  const c=A.monthly(idx,account,'2026-01').stats.categories[0];assert.equal(c.openingYi,1);assert.equal(c.knownYi,0);assert.equal(c.oneTwoYi,0);
});
test('primary allocation and later secondary sale match quantities, not equal cash amounts',()=>{
  const p=flow('B1','2026-01-03','分销买入',1,'信用债','分销买卖',10000);
  const s=flow('B1','2026-02-01','卖出',1.2,'信用债','现券买卖',6000);
  const {idx,r}=result(store([p,s]));assert.equal(r.stats.total,1.2);assert.equal(r.stats.primaryYi,1);assert.equal(r.stats.matchedPrimary,1.2);
  assert(r.text.includes('一级分销买入'));assert(r.text.includes('二级市场卖出'));
  const m=A.monthly(idx,account,'2026-02');assert(m.text.includes('此前分销买入'));assert.equal(m.stats.categories[0].primaryYi,1.2);
});
test('partial quantities cannot attribute the whole sale to primary allocation',()=>{
  const p=flow('B1','2026-01-03','分销买入',.4,'信用债','分销买卖',4000);
  const b=flow('B1','2026-01-04','买入',.2,'信用债','现券买卖',2000);
  const s=flow('B1','2026-01-05','卖出',.8,'信用债','现券买卖',8000);
  const {r}=result(store([p,b,s]));assert(Math.abs(r.stats.matchedPrimary-.4)<1e-8);assert.equal(r.stats.categories[0].unmatchedYi,.2);
});
test('same code in different markets does not consume an unrelated opening lot',()=>{
  const opening=bond('B1',1);opening.market='深交所';const {idx}=result(store([flow('B1','2026-02-01','卖出',1)],{}, {bonds:[opening]}));
  const c=A.monthly(idx,account,'2026-02').stats.categories[0];assert.equal(c.openingYi,0);assert.equal(c.unmatchedYi,1);
});
test('credit holdings and rate-bond transactions can form different portfolio roles',()=>{
  const flows=[];for(let i=0;i<3;i++){flows.push(flow('R'+i,`2026-06-0${i*2+1}`,'买入',.1,'利率债'));flows.push(flow('R'+i,`2026-06-0${i*2+2}`,'卖出',.101,'利率债','现券买卖',1000));}
  flows.push(flow('C1','2026-06-10','买入',1));
  const {r}=result(store(flows,{bonds:[bond('C1',9),bond('R0',.1,'国债',15)]}));
  assert(r.text.includes('信用债'));assert(r.text.includes('利率债'));assert(r.features.some(f=>f.key==='credit-rate'));
});
test('ETF themes aggregate managers and periods without looking through the fund',()=>{
  const f1=flow('F1','2026-06-01','买入',.2,'基金','基金买卖'),f2=flow('F2','2026-06-02','买入',.3,'基金','基金买卖');
  const {r}=result(store([f1,f2],{funds:[fund('F1',.2,'科创债示例甲ETF'),fund('F2',.3,'科创债示例乙ETF'),fund('F3',.5,'城投债示例ETF')]}));
  assert(r.text.includes('科创债类约占50%'));assert(r.text.includes('城投债类约占50%'));assert(r.text.includes('买入配置'));
});
test('sell then repurchase is distinct from a proved same-day buy-sell close',()=>{
  const sold=flow('F1','2026-06-03','卖出',.1,'基金','基金买卖',1000,'2026-06-03 10:00:00');
  const bought=flow('F1','2026-06-03','买入',.1,'基金','基金买卖',1000,'2026-06-03 14:00:00');
  const {r}=result(store([sold,bought],{funds:[fund('F1',1,'示例债券ETF')]},{funds:[fund('F1',1,'示例债券ETF')]}));
  const c=r.stats.categories[0];assert.equal(c.sameDayPairs,1);assert.equal(c.sellThenBuy,1);assert.equal(c.sameDayYi,0);assert(r.text.includes('同日卖出后买回'));
});
test('a later system-entry timestamp cannot prove same-day transaction order',()=>{
  const b=flow('F1','2026-06-03','买入',.1,'基金','基金买卖'),s=flow('F1','2026-06-03','卖出',.1,'基金','基金买卖');
  b.交易时间='2026-08-01 10:00:00';s.交易时间='2026-08-01 11:00:00';
  const {r}=result(store([b,s]));const c=r.stats.categories[0];assert.equal(c.sameDayPairs,1);assert.equal(c.sameDayYi,0);assert.equal(c.knownYi,0);
});
test('many small convertible trades cannot become the main capital-trading category',()=>{
  const flows=[flow('C1','2026-06-01','买入',10)];for(let i=0;i<12;i++)flows.push(flow('V'+(i%3),`2026-06-${String(i+2).padStart(2,'0')}`,i%2?'卖出':'买入',.0001,'转债'));
  const {r}=result(store(flows,{bonds:[bond('C1',10)]}));const f=r.features.find(f=>f.key==='small-转债');assert(f);assert(/金额(?:仅占|不足)/.test(f.text));assert(!f.text.includes('金额仅占0.0%'));assert(r.stats.categories[0].name==='信用债');
});
test('unknown classifications and ratings stay in their denominators',()=>{
  const {r}=result(store([flow('X1','2026-06-01','买入',.4,'其他'),flow('C1','2026-06-02','买入',.6)],{bonds:[bond('C1',.3,'企业债',null,'AA'),bond('C2',.7,'企业债',null,'')]}));
  assert.equal(r.stats.total,1);assert.equal(r.holding.rateNames.length,0);assert.equal(r.holding.term,null);assert(!r.text.includes('0—1年'));assert(r.flags.some(x=>x.includes('尚未识别')));
});
test('AA parenthetical notation combines with AA while unconfirmed spellings stay separate',()=>{
  const {r}=result(store([],{bonds:[bond('C1',.4,'企业债',2,'AA'),bond('C2',.4,'企业债',2,'AA(2)'),bond('C3',.2,'企业债',2,'AA2')]}));
  assert.equal(r.holding.rates.find(x=>x.name==='AA').value,.8);assert.equal(r.holding.rates.find(x=>x.name==='AA2').value,.2);assert(r.text.includes('评级以AA为主'));
});
test('negative positions do not get silently normalized to a clean allocation percentage',()=>{
  const {r}=result(store([],{bonds:[bond('C1',2),bond('C2',-.5)]}));assert.equal(r.holding.valid,false);assert(r.flags.some(x=>x.includes('负市值')));assert(!r.features.find(f=>f.key==='holdings').text.includes('占100%'));
});
test('historical monthly reports never substitute the uploaded end snapshot',()=>{
  const {idx}=result(store([flow('C1','2026-03-02','买入',1),flow('C2','2026-06-02','买入',1)],{bonds:[bond('C2',1)]}));
  assert(!A.monthly(idx,account,'2026-03').features.some(f=>f.key==='monthly-holdings'));assert(A.monthly(idx,account,'2026-06').features.some(f=>f.key==='monthly-holdings'));
});
test('strategy inception date takes priority over IAM date, not the first observed trade',()=>{
  const dims={strat:{[account]:{sacName:'示例账户',startDate:'2026-05-20',cat2:'FOF（纯债）策略'}},dim:{[account]:{startDate:'2025-01-01'}}};
  const {r}=result(store([flow('F1','2026-05-25','买入',1,'基金','基金买卖')]),dims);assert.equal(r.startDate,'2026-05-20');assert.equal(r.startLabel,'2026年5月新户');
});
test('holdings-only, primary-only, repo-only and unknown cases have factual fallbacks',()=>{
  const held=result(store([],{bonds:[bond('C1',1)]}));assert(held.r.text.includes('持仓'));
  const primary=result(store([flow('C1','2026-06-01','分销买入',1,'信用债','分销买卖')]));assert(primary.r.text.includes('一级分销买入'));assert(!primary.r.text.includes('未见常规交易记录'));
  const repo=flow('RP1','2026-06-01','逆回购',1);const rr=result(store([],{}, {},{hg:[repo]}));assert(rr.r.text.includes('逆回购'));
  const unknown=result(store([flow('X1','2026-06-01','买入',1,'其他')]));assert(unknown.r.text.includes('其他'));assert(!unknown.r.text.includes('信用债为主'));
});
test('each selected narrative feature resolves its own supporting records',()=>{
  const {idx,r}=result(store([flow('C1','2026-06-01','买入',1)],{bonds:[bond('C1',1)]}));
  for(const f of r.features){const e=A.evidence(idx,account,f.key,'all');assert(e);assert.equal(e.records.length,f.recordIds.length);for(const row of e.records)assert(row.acct===account);}
});
test('duplicate identifiers suppress inference of repeated trading without changing totals',()=>{
  const flows=[];for(let i=0;i<3;i++){flows.push(flow('R1',`2026-06-0${i*2+1}`,'买入',1,'利率债'));flows.push(flow('R1',`2026-06-0${i*2+2}`,'卖出',1,'利率债'));}flows[1].交易编号=flows[0].交易编号;
  const {r}=result(store(flows));assert.equal(r.stats.total,6);assert.equal(r.stats.categories[0].rolling,false);assert(r.flags.some(x=>x.includes('重复交易编号')));
});
test('calendar gaps are not presented as a previous-month scale comparison',()=>{
  const positions=[{_code:account,_dateStr:'2026-04-30',_size:1e8,_total:1e8},{_code:account,_dateStr:'2026-06-30',_size:2e8,_total:2e8}];
  const {idx}=result(store([],{}, {},{pos:positions}));const m=A.monthly(idx,account,'2026-06');assert.equal(m.stats.prior,null);assert(!m.text.includes('增加100'));
});
test('shorter rate holdings and longer rate trading are described as separate roles',()=>{
  const flows=[];for(let i=0;i<3;i++){const b=flow('L'+i,`2026-06-0${i*2+1}`,'买入',1,'利率债'),s=flow('L'+i,`2026-06-0${i*2+2}`,'卖出',1,'利率债');b.期限=s.期限='10-30年';flows.push(b,s);}
  const {r}=result(store(flows,{bonds:[bond('S1',3,'国债',4)]}));assert(r.features.some(f=>f.key==='long-rate'));assert(r.text.includes('持仓偏3—5年'));assert(r.text.includes('7年以上长端'));
});
test('equity-themed ETFs are not described as bond ETFs',()=>{
  const {r}=result(store([flow('E1','2026-06-01','买入',1,'基金','基金买卖')],{funds:[fund('E1',1,'红利示例ETF')]}));assert(r.text.includes('红利类'));assert(!r.text.includes('债券ETF'));
});
test('end-snapshot quality warnings are not attached to a historical month',()=>{
  const {idx}=result(store([flow('C1','2026-05-01','买入',1),flow('C1','2026-06-01','卖出',1)],{bonds:[bond('C1',2),bond('C2',-.1)]}));assert(!A.monthly(idx,account,'2026-05').flags.some(x=>x.includes('负市值')));assert(A.monthly(idx,account,'2026-06').flags.some(x=>x.includes('负市值')));
});
test('filtered records have a traceable evidence view without re-entering statistics',()=>{
  const clean=flow('C1','2026-06-01','买入',1),filtered=flow('C2','2026-06-02','买入',100);
  const {idx,r}=result(store([clean]),{}, {quality:[{rule:'R2',level:'过滤',row:filtered,detail:'演示金额超过最近规模'}]});const e=A.evidence(idx,account,'quality','all');assert.equal(r.stats.total,1);assert.equal(e.records.length,1);assert.equal(e.records[0].tradeId,filtered.交易编号);assert.equal(e.records[0].source.row,filtered._source.row);
});
test('observed gaps gain an empty-month report without fabricated positions',()=>{
  const positions=[{_code:account,_dateStr:'2026-04-30',_size:1e8,_total:1e8},{_code:account,_dateStr:'2026-06-30',_size:1e8,_total:1e8}];const {idx}=result(store([],{}, {},{pos:positions}));assert(idx.accounts.get(account).months.includes('2026-05'));const m=A.monthly(idx,account,'2026-05');assert(m.text.includes('未见常规交易'));assert(!m.features.some(f=>f.key==='monthly-holdings'));
});
test('monthly source evidence includes earlier primary purchases and opening holdings',()=>{
  const primary=flow('P1','2026-02-01','分销买入',1,'信用债','分销买卖'),sale=flow('P1','2026-06-01','卖出',.5,'信用债'),oldSale=flow('O1','2026-06-02','卖出',.5,'信用债');
  const {idx}=result(store([primary,sale,oldSale],{}, {bonds:[bond('O1',1)]}));const e=A.evidence(idx,account,'monthly-信用债','2026-06');
  assert(e.records.some(r=>r.tradeId===primary.交易编号));assert(e.records.some(r=>r.kind==='holding'&&r.snapshot==='start'));assert.equal(A.monthly(idx,account,'2026-06').stats.total,1);
});
test('material secondary holdings and trading categories are mentioned',()=>{
  const {r}=result(store([flow('C1','2026-06-01','买入',.7),flow('V1','2026-06-02','买入',.3,'转债')],{bonds:[bond('C1',.7)],funds:[fund('F1',.3,'示例债券基金')]}));
  assert(r.features.find(f=>f.key==='holdings').text.includes('另持有基金（约占30%）'));assert(r.features.find(f=>f.key==='main-trading').text.includes('兼有转债交易'));
});
test('rounding does not imply an exclusive holding when minor assets are present',()=>{
  const {r}=result(store([],{bonds:[bond('C1',9.99)],funds:[fund('F1',.01,'示例债券基金')]}));const text=r.features.find(f=>f.key==='holdings').text;assert(!text.includes('占100%'));assert(text.includes('少量持有基金'));
});
