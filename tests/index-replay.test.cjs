// Fictional data only. Run: node --test tests/*.test.cjs
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const html=fs.readFileSync(path.join(__dirname,'..','交易数据分析平台.html'),'utf8');
const extract=name=>html.match(new RegExp('// '+name+'_BEGIN([\\s\\S]*?)// '+name+'_END'))[1];
const c=vm.createContext({Date});new vm.Script(extract('BOND_MARKET')+'\n'+extract('INDEX_REPLAY')).runInContext(c);const I=vm.runInContext('IndexReplay',c),M=vm.runInContext('BondMarket',c);
const all={level:'department',dept:'all'},first='2026-01-05',last='2026-01-09';
const quotes=(dates=[first,'2026-01-06','2026-01-07','2026-01-08',last,'2026-01-12'])=>I.merge([I.parseSheet([['日期',...I.CATEGORIES.map(c=>c.name)],['Date',...I.CATEGORIES.map(c=>c.code)],...dates.map((d,i)=>[d,...I.CATEGORIES.map(()=>100+i)])],'fictional.xlsx','指数')]);
const bond=(acct='DEMO_A',code='DEMO_CREDIT',mv=1000,type='公司债',remain=3)=>({acct,code,name:'虚构债券',market:'银行间',mv,bondType:type,remain,_source:{file:'fictional-holdings.xlsx',sheet:'证券余额',row:2}});
const fund=(code,name,kind='基金')=>({acct:'DEMO_A',code,name,market:'上交所',mv:1000,kind});
const flow=(date,dir,amount,acct='DEMO_A',code='DEMO_CREDIT',variety='现券买卖')=>({_code:acct,_name:'虚构账户',_dateStr:date,_dir:dir,_amountYi:amount,_variety:variety,代码:code,_name2:'虚构债券',_mkt:'银行间',_source:{file:'fictional-flow.xlsx',sheet:'流水',row:3}});
const position=(date,size,acct='DEMO_A')=>({_code:acct,_dateStr:date,_size:size*1e8});
function fixture(options={}){const opening=options.opening||[bond()],ending=options.ending||[bond('DEMO_A','DEMO_CREDIT',1000)];return{_store:{hg:options.hg||[{_code:'DEMO_A'}],hdStart:{dateStr:'01.05',bonds:opening,funds:options.startFunds||[]},hdEnd:{dateStr:'01.09',bonds:ending,funds:options.endFunds||[]},zqAll:options.flows||[],pos:options.pos||[position(first,1),position(last,1)]},acctDim:{dim:options.dim||{DEMO_A:{line:'虚构部门',managers:['演示经理甲'],sacName:'虚构账户'}},strat:{}},indexMarket:quotes()};}
const build=o=>I.build(fixture(o));
test('all six index identifiers are parsed separately including CS CSI and leading-zero SH',()=>{
 const d=quotes();assert.equal(d.series.size,6);assert.equal(d.series.get('equityFund').code,'000300.SH');assert.equal(d.series.get('credit').quotes.length,6);assert.equal(d.start,first);
});
test('ordinary bond or fund sheets are not mistaken for the index sheet',()=>{
 assert.equal(I.parseSheet([['日期','虚构基金'],['Date','000300.SH'],[first,1]],'fictional.xlsx','基金'),null);
 assert.equal(I.parseSheet([['项目编号','账户'],['DEMO_A','虚构账户']],'fictional.xlsx','策略'),null);
});
test('duplicate equal index quotes collapse while conflicting dates are excluded',()=>{
 const p=I.parseSheet([['日期','虚构'],['Date','CBA20901.CS'],[first,100]],'fictional.xlsx','指数');assert.equal(I.merge([p,p]).series.get('credit').quotes.length,1);
 const p2=I.parseSheet([['日期','虚构'],['Date','CBA20901.CS'],[first,101]],'fictional-b.xlsx','指数');const d=I.merge([p,p2]);assert.equal(d.conflicts,1);assert.equal(d.series.get('credit').quotes.length,0);
});
test('the user cash-holding example is 10m then 30m unchanged then 20m independent of quotes',()=>{
 const idx=build({flows:[flow('2026-01-07','买入',.2),flow(last,'卖出',.1)],ending:[bond('DEMO_A','DEMO_CREDIT',2000)]}),s=I.scope(idx,all);
 for(const[d,expected]of [[first,.1],['2026-01-07',.3],['2026-01-08',.3],[last,.2]])assert(Math.abs(s.daily.get(d).values.credit.amount-expected)<1e-12);
 assert.equal(s.checks.length,0);assert(Math.abs(s.totals.credit.ratio-.2)<1e-12);
});
test('opening-day transactions are not counted twice and later dates stop at the closing snapshot',()=>{
 const idx=build({flows:[flow(first,'买入',99),flow('2026-01-06','买入',.2),flow('2026-01-12','买入',88)]}),s=I.scope(idx,all);assert.equal(s.events.length,1);assert(Math.abs(s.totals.credit.projected-.3)<1e-12);
});
test('primary buying increases holdings and net buying; repos futures and share adjustments are excluded',()=>{
 const flows=[flow('2026-01-06','分销买入',.2,'DEMO_A','DEMO_CREDIT','分销买卖'),flow('2026-01-07','买入',99,'DEMO_A','DEMO_CREDIT','国债期货'),flow('2026-01-07','买入',99,'DEMO_A','DEMO_CREDIT','质押式回购'),flow('2026-01-07','份额调整',99,'DEMO_A','DEMO_CREDIT','基金买卖')];
 const s=I.scope(build({flows}),all);assert.equal(s.events.length,1);assert.equal(s.totals.credit.buy,.2);assert.equal(s.totals.credit.net,.2);
});
test('bond put-back reduces holdings without appearing as active sell or negative net trade',()=>{
 const s=I.scope(build({flows:[flow('2026-01-07','债券回售',.04,'DEMO_A','DEMO_CREDIT','债券回售')]}),all);assert.equal(s.totals.credit.repay,.04);assert.equal(s.totals.credit.sell,0);assert.equal(s.daily.get('2026-01-07').values.credit.net,0);assert(Math.abs(s.totals.credit.projected-.06)<1e-12);
});
test('same-day gross buying and selling can cancel the marker net without disappearing from evidence',()=>{
 const s=I.scope(build({flows:[flow('2026-01-06','买入',.02),flow('2026-01-06','卖出',.02)]}),all),d=s.daily.get('2026-01-06').values.credit;assert.equal(d.net,0);assert.equal(d.records.length,2);assert.equal(d.buy,.02);assert.equal(d.sell,.02);
});
test('shared accounts split flows holdings and net size equally while department counts each once',()=>{
 const idx=build({dim:{DEMO_A:{line:'虚构部门',managers:['演示甲','演示乙']}},flows:[flow('2026-01-06','买入',.2)]});
 const a=I.scope(idx,{level:'manager',dept:'all',manager:'虚构部门|演示甲'}),b=I.scope(idx,{level:'manager',dept:'all',manager:'虚构部门|演示乙'}),d=I.scope(idx,all);
 assert.equal(a.opening.credit,.05);assert.equal(a.events[0].amount,.1);assert.equal(a.daily.get(last).netSize,.5);assert(Math.abs(a.totals.credit.projected+b.totals.credit.projected-d.totals.credit.projected)<1e-12);assert.equal(d.selected.length,1);assert.equal(d.daily.get(last).netSize,1);
});
test('group position ratio divides summed holdings by summed size rather than averaging account ratios',()=>{
 const idx=build({opening:[bond(),bond('DEMO_B','DEMO_CREDIT',4500)],ending:[bond(),bond('DEMO_B','DEMO_CREDIT',4500)],hg:[{_code:'DEMO_A'},{_code:'DEMO_B'}],pos:[position(first,1),position(last,1),position(first,9,'DEMO_B'),position(last,9,'DEMO_B')]}),d=I.scope(idx,all).daily.get(last);assert(Math.abs(d.values.credit.ratio-.055)<1e-12);
});
test('leverage may produce a position ratio above 100 percent',()=>{
 const s=I.scope(build({opening:[bond('DEMO_A','DEMO_CREDIT',15000)],pos:[position(first,1)]}),all);assert.equal(s.daily.get(last).values.credit.ratio,1.5);
});
test('market line can extend beyond analysis while future holdings and ratios remain missing',()=>{
 const s=I.scope(build(),all);assert(s.calendar.includes('2026-01-12'));assert.equal(s.daily.get('2026-01-12').values.credit.amount,null);assert.equal(s.daily.get('2026-01-12').values.credit.ratio,null);assert.equal(s.daily.get('2026-01-12').netSize,null);
});
test('convertible ETF beats equity ETF classification but a regular convertible fund stays other fund',()=>{
 const result=fixture({endFunds:[fund('DEMO_CONV','虚构可转债ETF','股票ETF'),fund('DEMO_NORMAL','虚构可转债债券基金'),fund('DEMO_LINK','虚构可转债ETF联接基金'),fund('DEMO_STOCK','虚构股票ETF','股票ETF'),fund('DEMO_GOLD','虚构黄金ETF','商品ETF'),fund('DEMO_REIT','虚构REIT','REITs')]});
 const meta=I.metadata(result._store),cat=code=>I.classify({code,market:'上交所'},'fund',meta);assert.equal(cat('DEMO_CONV'),'convertible');assert.equal(cat('DEMO_NORMAL'),'otherFund');assert.equal(cat('DEMO_LINK'),'otherFund');assert.equal(cat('DEMO_STOCK'),'equityFund');assert.equal(cat('DEMO_GOLD'),'otherFund');assert.equal(cat('DEMO_REIT'),'otherFund');
});
test('seven-year boundary missing term and greater-than-thirty-year rate bonds use confirmed buckets',()=>{
 const r=fixture({ending:[bond('DEMO_A','EXACT7',1000,'国债',7),bond('DEMO_A','OVER7',1000,'国债',7.01),bond('DEMO_A','OVER30',1000,'国债',50),bond('DEMO_A','MISSING',1000,'国债',0)]}),m=I.metadata(r._store);
 assert.equal(I.classify({code:'EXACT7',market:'银行间'},'bond',m),'rateShort');assert.equal(I.classify({code:'MISSING',market:'银行间'},'bond',m),'rateShort');assert.equal(I.classify({code:'OVER7',market:'银行间'},'bond',m),'rateLong');assert.equal(I.classify({code:'OVER30',market:'银行间'},'bond',m),'rateLong');
});
test('closing snapshot term applies to opening and historic flows consistently without migration',()=>{
 const idx=build({opening:[bond('DEMO_A','DEMO_RATE',1000,'国债',7.5)],ending:[bond('DEMO_A','DEMO_RATE',1000,'国债',6.9)],flows:[flow('2026-01-06','买入',.2,'DEMO_A','DEMO_RATE')]}),s=I.scope(idx,all);assert.equal(s.opening.rateShort,.1);assert.equal(s.opening.rateLong,0);assert.equal(s.events[0].bucket,'rateShort');
});
test('unrecognized bond category is retained as unclassified rather than invented as short rate',()=>{
 const r=fixture({opening:[],ending:[],flows:[flow('2026-01-06','买入',.1,'DEMO_A','UNKNOWN')]});r._store.hdStart.bonds=[];r._store.hdEnd.bonds=[];const s=I.scope(I.build(r),all);assert.equal(s.events.length,0);assert.equal(s.unknown.length,1);
});
test('exact ten percent is not flagged but greater-than-ten percent is flagged without cash floor',()=>{
 const at=I.scope(build({opening:[bond('DEMO_A','DEMO_CREDIT',1100)],ending:[bond('DEMO_A','DEMO_CREDIT',1000)]}),all);assert.equal(at.checks.length,0);
 const tiny=I.scope(build({opening:[bond('DEMO_A','DEMO_CREDIT',.0012)],ending:[bond('DEMO_A','DEMO_CREDIT',.001)]}),all);assert.equal(tiny.checks.length,1);assert(tiny.checks[0].diff<.000001);
});
test('zero actual with nonzero projection is flagged and negative balances are not clamped',()=>{
 const r=fixture({opening:[],ending:[bond('DEMO_A','DEMO_CREDIT',0)],flows:[flow('2026-01-06','卖出',.00001)]});r._store.hdStart.bonds=[];const s=I.scope(I.build(r),all);assert.equal(s.checks.length,1);assert.equal(s.checks[0].ratio,Infinity);assert(s.totals.credit.projected<0);
});
test('opposite account errors remain visible even when department totals reconcile',()=>{
 const idx=build({opening:[bond('DEMO_A','DEMO_CREDIT',2000),bond('DEMO_B','DEMO_CREDIT',0)],ending:[bond(),bond('DEMO_B','DEMO_CREDIT',1000)],hg:[{_code:'DEMO_A'},{_code:'DEMO_B'}],pos:[position(first,1),position(first,1,'DEMO_B')]}),s=I.scope(idx,all);assert.equal(s.checks.length,2);assert(Math.abs(s.totals.credit.projected-s.actual.credit)<1e-12);
});
test('scale gap fill never looks ahead and scope caches cannot leak between hierarchy objects',()=>{
 const idx=build({pos:[position('2026-01-07',2)]}),s=I.scope(idx,all);assert.equal(s.daily.get(first).values.credit.ratio,null);assert.equal(s.daily.get('2026-01-06').values.credit.ratio,null);assert.equal(s.daily.get('2026-01-08').values.credit.ratio,.05);assert.equal(s.daily.get('2026-01-08').carriedSize,1);
 assert.equal(s.daily.get('2026-01-08').oldestScaleDate,'2026-01-07');const empty=I.scope(idx,{level:'account',dept:'all',account:'NO_MATCH'});assert.equal(empty.selected.length,0);assert.equal(s.selected.length,1);
});
test('new modules have no networking or persistent business storage and index prices never join security data',()=>{
 const source=extract('INDEX_REPLAY')+extract('INDEX_REPLAY_UI');assert(!/\bfetch\s*\(|XMLHttpRequest|WebSocket|sendBeacon|localStorage|sessionStorage|indexedDB/.test(source));assert(html.indexOf('IndexReplay.parseSheet(aoa')<html.indexOf('BondMarket.parseSheet(aoa'));assert(html.includes("connect-src 'none'"));
});
test('account universe is flow and size intersection including repo and futures while references cannot add accounts',()=>{
 const r=fixture({opening:[bond('HELD_ONLY')],ending:[bond('HELD_ONLY')],flows:[flow('2026-01-06','买入',.01,'FLOW_ONLY'),flow('2026-01-06','买入',.01,'BOTH')],pos:[position(first,1,'SCALE_ONLY'),position(first,1,'BOTH'),position(first,1,'REPO_BOTH'),position(first,1,'FUTURE_BOTH')],dim:{IAM_ONLY:{line:'虚构部门'}}});
 r._store.hg=[{_code:'REPO_BOTH',_name:'虚构回购账户'},{_code:'REPO_NO_SCALE'}];r._store.zqAll.push(flow('2026-01-06','买入',.01,'FUTURE_BOTH','DEMO_F','国债期货'));r.acctDim.strat={STRATEGY_ONLY:{sacName:'虚构策略账户'}};
 const idx=I.build(r);assert.deepEqual([...idx.accounts.keys()].sort(),['BOTH','FUTURE_BOTH','REPO_BOTH']);assert.equal(I.scope(idx,all).opening.credit,0);
});
test('IAM is the sole department authority even when size records disagree or IAM is missing',()=>{
 const r=fixture({hg:[{_code:'DEMO_A'},{_code:'DEMO_B'}],pos:[{...position(first,1),_belong:'财富产品部'},{...position(first,1,'DEMO_B'),_belong:'财富产品部'}],dim:{DEMO_A:{line:'投资顾问部'}}});
 r.deptMap={DEMO_A:'财富产品部',DEMO_B:'财富产品部'};const idx=I.build(r);assert.equal(idx.accounts.get('DEMO_A').dept,'投资顾问部');assert.equal(idx.accounts.get('DEMO_B').dept,'未匹配');
});
test('chart dates are valid prices plus actual trades, not scale-only dates; internal end reconciliation remains complete',()=>{
 const r=fixture({pos:[position(first,1),position('2026-01-10',1)]});r._store.hdEnd.dateStr='01.10';const idx=I.build(r),s=I.scope(idx,all),series=idx.market.series.get('credit');
 assert(s.daily.get('2026-01-10').values.credit.amount>0);assert(!I.plotDates(series.quotes,[],first).includes('2026-01-10'));assert(I.plotDates(series.quotes,[{date:'2026-01-10'}],first).includes('2026-01-10'));
 const ui=extract('INDEX_REPLAY_UI')+extract('BOND_MARKET_UI');assert(!ui.includes('connectNulls:false'));assert(ui.includes('IndexReplay.plotDates'));
});
test('individual holding cash path includes primary buys and put-back but excludes opening-day double count',()=>{
 const r=fixture({flows:[flow(first,'买入',99),flow('2026-01-06','分销买入',.2,'DEMO_A','DEMO_CREDIT','分销买卖'),flow('2026-01-07','卖出',.1),flow(last,'债券回售',.05,'DEMO_A','DEMO_CREDIT','债券回售')],ending:[bond('DEMO_A','DEMO_CREDIT',1500)]});
 const p=I.securityPosition(I.build(r),'DEMO_A',{code:'DEMO_CREDIT',market:'IB',kind:'bond'});assert(Math.abs(p.daily.get('2026-01-06').amount-.3)<1e-12);assert(Math.abs(p.projected-.15)<1e-12);assert(Math.abs(p.daily.get(last).ratio-.15)<1e-12);assert.equal(p.buy,.2);assert.equal(p.repay,.05);assert(!p.flagged);assert.equal(p.daily.get('2026-01-12').amount,null);
});
test('individual position uses all earlier months even when displayed trade table selects a later month',()=>{
 const r=fixture({flows:[flow('2026-01-06','买入',.2),flow('2026-02-02','卖出',.1)]});r._store.hdEnd.dateStr='02.03';r.indexMarket=quotes([first,'2026-01-06','2026-02-02','2026-02-03']);
 const idx=I.build(r),before=I.securityPosition(idx,'DEMO_A',{code:'DEMO_CREDIT',market:'IB',kind:'bond'});const table=M.summarize(M.index(r._store.zqAll,M.merge()),'DEMO_A','2026-02');assert.equal(table.tradeCount,1);assert(Math.abs(before.projected-.2)<1e-12);assert(Math.abs(before.daily.get('2026-02-02').amount-.2)<1e-12);
});
test('held-only securities within an in-scope account can be selected without changing trade summary',()=>{
 const idx=build(),options=I.securityOptions(idx,'DEMO_A',[]);assert.equal(options.length,1);assert.equal(options[0].count,0);assert.equal(options[0].kind,'bond');const p=I.securityPosition(idx,'DEMO_A',options[0]);assert.equal(p.opening,.1);assert.equal(p.projected,.1);
});
test('individual holding identity separates bond fund and exchange markets with the same code',()=>{
 const r=fixture({opening:[{...bond(),code:'SHARED',market:'银行间'},{...bond(),code:'SHARED',market:'上交所',mv:2000}],startFunds:[{...fund('SHARED','虚构基金'),mv:3000}],endFunds:[{...fund('SHARED','虚构基金'),mv:3000}]});
 const idx=I.build(r),get=(kind,market)=>I.securityPosition(idx,'DEMO_A',{code:'SHARED',kind,market}).opening;
 assert.equal(get('bond','IB'),.1);assert.equal(get('bond','SH'),.2);assert.equal(get('fund','SH'),.3);assert.equal(get('bond','SZ'),0);
});
test('fund leading zeros and generic market reconcile with full quote identity; fund net price is ignored',()=>{
 const r=fixture({opening:[],ending:[],startFunds:[{...fund('000123','虚构基金'),market:'场外市场',mv:1000}],endFunds:[{...fund('000123','虚构基金'),market:'场外市场',mv:3000}],flows:[{...flow('2026-01-06','买入',.2,'DEMO_A','123','基金买卖'),_mkt:'场外市场','全价(元)':1.01,'净价(元)':999}]});
 r.bondMarket=M.merge([M.parseSheet([['日期','虚构基金'],['Date','000123.OF'],[first,1],[last,1.01]],'fictional.xlsx','基金')]);
 const idx=I.build(r),options=I.securityOptions(idx,'DEMO_A',[],'fund');assert.equal(options.length,1);assert(options[0].series);const p=I.securityPosition(idx,'DEMO_A',options[0]);assert(Math.abs(p.projected-.3)<1e-12);assert(!p.flagged);
});
test('individual zero actual nonzero estimate flags discrepancy and missing scale leaves position intact',()=>{
 const r=fixture({pos:[position('2026-01-08',1)],ending:[bond('DEMO_A','DEMO_CREDIT',0)]}),idx=I.build(r),p=I.securityPosition(idx,'DEMO_A',{code:'DEMO_CREDIT',market:'IB',kind:'bond'});assert.equal(p.daily.get(first).amount,.1);assert.equal(p.daily.get(first).ratio,null);assert(p.flagged);assert.equal(p.discrepancy,Infinity);
});
test('science innovation debt ETFs remain other funds while equity science ETFs stay equity funds',()=>{
 const r=fixture({endFunds:[fund('DEMO_DEBT1','虚构科创债ETF'),fund('DEMO_DEBT2','虚构科创公司债ETF'),fund('DEMO_DEBT3','虚构科技创新债券ETF'),fund('DEMO_EQUITY','虚构科创50ETF'),fund('DEMO_CONV','虚构转债ETF')]});const m=I.metadata(r._store),cat=code=>I.classify({code,market:'上交所'},'fund',m);
 for(const c of ['DEMO_DEBT1','DEMO_DEBT2','DEMO_DEBT3'])assert.equal(cat(c),'otherFund');assert.equal(cat('DEMO_EQUITY'),'equityFund');assert.equal(cat('DEMO_CONV'),'convertible');
});
test('missing-day guide is a display annotation and never mutates raw quotes or later return pricing',()=>{
 const q=[['2026-01-05',100],['2026-01-07',104]],guide=I.quoteGuide(q,'2026-01-06');assert.equal(guide.value,102);assert(guide.missing);assert.equal(guide.reference,'line');assert.equal(q.length,2);assert(!q.some(x=>x[0]==='2026-01-06'));
 assert(!I.quoteGuide(q,'2026-01-05').missing);assert.equal(I.quoteGuide(q,'2026-01-08').reference,'nearest');assert.equal(I.quoteGuide([],'2026-01-06').reference,'none');
});
test('a missing-price trading day retains its true cash position and net trade while a no-trade gap is omitted',()=>{
 const r=fixture({flows:[flow('2026-01-06','买入',.2)]});r.indexMarket=quotes([first,'2026-01-08',last]);const idx=I.build(r),s=I.scope(idx,all),dates=I.plotDates(idx.market.series.get('credit').quotes,s.events,idx.start);
 assert(dates.includes('2026-01-06'));assert(!dates.includes('2026-01-07'));assert(Math.abs(s.daily.get('2026-01-06').values.credit.amount-.3)<1e-12);assert.equal(s.daily.get('2026-01-06').values.credit.net,.2);
});
test('negative source values are preserved and a cash outflow balance can be negative with zero actual shares',()=>{
 const r=fixture({opening:[],ending:[],startFunds:[{...fund('DEMO_NAV','虚构普通基金'),mv:1000}],endFunds:[{...fund('DEMO_NAV','虚构普通基金'),mv:0}],flows:[{...flow(last,'卖出',.11,'DEMO_A','DEMO_NAV','基金买卖'),_mkt:'上交所'}]});const idx=I.build(r),p=I.securityPosition(idx,'DEMO_A',{code:'DEMO_NAV',market:'SH',kind:'fund'});
 assert(Math.abs(p.projected+.01)<1e-12);assert.equal(p.actual,0);assert.equal(p.sell,.11);
 r._store.hdStart.funds[0].mv=-1000;r._store.zqAll=[];const n=I.securityPosition(I.build(r),'DEMO_A',{code:'DEMO_NAV',market:'SH',kind:'fund'});assert.equal(n.opening,-.1);assert.equal(n.sell,0);
});
test('zero opening metadata before account activity does not block another accounts valid department ratio',()=>{
 const r=fixture({hg:[{_code:'DEMO_A'},{_code:'DEMO_B'}],opening:[bond(),bond('DEMO_B','DEMO_CREDIT',0)],pos:[position(first,1),position(last,1,'DEMO_B')]}),s=I.scope(I.build(r),all);assert.equal(s.daily.get(first).missingSize,0);assert.equal(s.daily.get(first).values.credit.ratio,.1);
});
