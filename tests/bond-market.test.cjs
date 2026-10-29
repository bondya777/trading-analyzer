// Fictional fixtures only. Run: node --test tests/bond-market.test.cjs
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const html=fs.readFileSync(path.join(__dirname,'..','交易数据分析平台.html'),'utf8');
const moduleSource=html.match(/\/\/ BOND_MARKET_BEGIN([\s\S]*?)\/\/ BOND_MARKET_END/)[1];
const ctx=vm.createContext({Date});new vm.Script(moduleSource).runInContext(ctx);const M=vm.runInContext('BondMarket',ctx);
const plain=v=>JSON.parse(JSON.stringify(v));
const sheet=(rows,codes=['DEMO_BOND.IB'],names=['虚构债券'])=>[['日期',...names],['Date',...codes],...rows];
const data=(rows,codes,names)=>M.merge([M.parseSheet(sheet(rows,codes,names),'fictional-quotes.xlsx','行情')]);
let next=0;
function flow(date,direction,amount=.01,price=100,face=100,code='DEMO_BOND',mkt='银行间',account='DEMO_ACCOUNT'){
 return{_code:account,_name:'虚构账户',_variety:'现券买卖',_dir:direction,_dateStr:date,_mkt:mkt,_amountYi:amount,_name2:'虚构债券',代码:code,'全价(元)':price,'交易面额(万)':face,交易编号:'DEMO_T'+(++next),_source:{file:'fictional-flow.xlsx',sheet:'流水',row:next+2}};
}
const first=(rows,d)=>M.summarize(M.index(rows,d)).groups[0];

test('wide quote headers preserve leading zero codes and accept Excel date serials',()=>{
 const d=data([[46026,100],['2026-01-05',101]],['000000000.IB']);
 assert(d.series.has('000000000|IB'));assert.equal(d.series.get('000000000|IB').quotes[0][0],'2026-01-04');assert.equal(d.stats.quoteCount,2);
});
test('irrelevant account sheets are not recognized as quote matrices',()=>{
 assert.equal(M.parseSheet([['项目编号','静态收益率'],['DEMO_ACCOUNT',2]],'fake.xlsx','策略'),null);
});
test('blank, fetching, errors, nonpositive prices and numeric prefixes are never coerced into prices',()=>{
 const values=[null,'Fetching...','#N/A',0,-3,'103oops',true,100.25];
 const d=data(values.map((v,i)=>['2026-06-'+String(i+1).padStart(2,'0'),v]));
 assert.equal(d.stats.quoteCount,1);assert.equal(d.series.values().next().value.quotes[0][1],100.25);assert.equal(d.stats.invalidValues,6);
});
test('a fetching name keeps the usable code and quote data',()=>{
 const d=data([['2026-06-01',100]],undefined,['Fetching...']);assert.equal(d.stats.quoteCount,1);assert.equal(d.series.values().next().value.name,'');
});
test('duplicate identical quotes do not change the price or observation count',()=>{
 const p=M.parseSheet(sheet([['2026-06-01',100]]),'fake.xlsx','行情'),d=M.merge([p,p]);assert.equal(d.stats.quoteCount,1);assert.equal(d.stats.duplicates,1);
});
test('conflicting same-day quotes are removed and cannot be silently carried over',()=>{
 const a=M.parseSheet(sheet([['2026-06-01',100],['2026-06-08',101]]),'fake-a.xlsx','行情');
 const b=M.parseSheet(sheet([['2026-06-08',102]]),'fake-b.xlsx','行情');
 const d=M.merge([a,b]),g=first([flow('2026-06-01','卖出')],d);assert.equal(d.stats.conflicts,1);assert.equal(g.week.full.count,0);
});
test('code plus market separates interbank and exchange securities',()=>{
 const d=data([['2026-06-01',101,202]],['DEMO_BOND.IB','DEMO_BOND.SH']);
 assert.equal(M.resolve(d,M.identity('DEMO_BOND','银行间')).series.quotes[0][1],101);
 assert.equal(M.resolve(d,M.identity('DEMO_BOND','上交所')).series.quotes[0][1],202);
 assert.equal(M.resolve(d,M.identity('DEMO_BOND','深交所')).series,null);
});
test('generic exchange location only matches a unique exchange market',()=>{
 const unique=data([['2026-06-01',101,202]],['DEMO_BOND.IB','DEMO_BOND.SH']);assert.equal(M.resolve(unique,M.identity('DEMO_BOND','交易所')).series.market,'SH');
 const ambiguous=data([['2026-06-01',101,202]],['DEMO_BOND.SH','DEMO_BOND.SZ']);assert.equal(M.resolve(ambiguous,M.identity('DEMO_BOND','交易所')).series,null);assert.equal(M.resolve(ambiguous,M.identity('DEMO_BOND','交易所')).reason,'市场不明确');
});
test('calendar-month horizons clamp month-end and handle year transitions',()=>{
 assert.equal(M.after('2026-01-31','month'),'2026-02-28');assert.equal(M.after('2028-01-31','month'),'2028-02-29');assert.equal(M.after('2026-12-31','month'),'2027-01-31');assert.equal(M.after('2026-06-01','week'),'2026-06-08');
});
test('invalid dates are rejected rather than rolled into a different month',()=>{
 assert.equal(M.day('2026-02-30'),'');assert.equal(M.day('2026-13-01'),'');assert.equal(M.day('20260203'),'2026-02-03');
});
test('a weekend target uses the latest prior quote and never borrows a future quote',()=>{
 const d=data([['2026-06-06',100],['2026-06-12',102],['2026-06-15',120]]),g=first([flow('2026-06-06','卖出')],d),o=g.sales[0].week;
 assert.equal(o.target,'2026-06-13');assert.equal(o.date,'2026-06-12');assert(Math.abs(o.return-.02)<1e-12);assert.equal(o.source.row,4);
});
test('quotes no later than the sell date are not a zero-return follow-up observation',()=>{
 const d=data([['2026-06-01',100]]),g=first([flow('2026-06-01','卖出')],d);assert.equal(g.sales[0].week.return,null);assert.equal(g.week.full.count,0);
});
test('short observation windows are explicitly partial',()=>{
 const d=data([['2026-06-15',100],['2026-06-20',101]]),g=first([flow('2026-06-15','卖出')],d);
 assert.equal(g.sales[0].week.status,'partial');assert.equal(g.week.full.count,0);assert.equal(g.week.partial.count,1);assert.equal(g.sales[0].week.date,'2026-06-20');
});
test('complete and partial sale cohorts remain separate when aggregated',()=>{
 const d=data([['2026-06-01',100],['2026-06-08',102],['2026-06-15',100],['2026-06-20',105]]),g=first([flow('2026-06-01','卖出'),flow('2026-06-15','卖出',.03)],d);
 assert.equal(g.week.full.count,1);assert.equal(g.week.partial.count,1);assert(Math.abs(g.week.full.return-.02)<1e-12);assert(Math.abs(g.week.partial.return-.05)<1e-12);
});
test('long-stale quotes are missing instead of assumed current',()=>{
 const d=data([['2026-06-01',100],['2026-06-10',103],['2026-07-01',null]]),g=first([flow('2026-06-01','卖出')],d);
 assert.equal(g.sales[0].month.status,'估值缺失');assert.equal(g.sales[0].month.return,null);
});
test('weighted sell price uses face while follow-up changes use each sale date and sell-amount weights',()=>{
 const d=data([['2026-06-01',100],['2026-06-02',110],['2026-06-08',102],['2026-06-09',111],['2026-07-02',115]]);
 const g=first([flow('2026-06-01','卖出',.01,100,100),flow('2026-06-02','卖出',.033,110,300)],d);
 assert.equal(g.price,107.5);const expected=(.01*.02+.033*(111/110-1))/.043;assert(Math.abs(g.week.full.return-expected)<1e-12);assert.equal(g.sell,.043000000000000003);
});
test('missing sell prices retain sold amounts but do not contribute invented returns',()=>{
 const d=data([['2026-06-01',100],['2026-06-08',102]]),g=first([flow('2026-06-01','卖出',.01,null)],d);
 assert.equal(g.sell,.01);assert.equal(g.price,null);assert.equal(g.week.full.return,null);assert(g.week.reasons.includes('缺卖出全价'));assert.equal(M.marks(g).length,0);
});
test('buy-only securities are not assigned zero follow-up returns',()=>{
 const d=data([['2026-06-01',100],['2026-06-08',102]]),g=first([flow('2026-06-01','买入')],d);assert.equal(g.week.total,0);assert.equal(g.week.full.return,null);
});
test('primary allocations, futures, repos and share adjustments do not enter the replay',()=>{
 const rows=[flow('2026-06-01','买入')];for(const variety of ['分销买卖','国债期货','质押式回购'])rows.push({...flow('2026-06-01','买入',10),_variety:variety});rows.push(flow('2026-06-01','分销买入',10));rows.push({...flow('2026-06-01','份额调整',10),_variety:'基金买卖'});
 const s=M.summarize(M.index(rows,M.merge()));assert.equal(s.tradeCount,1);assert.equal(s.amount,.01);
});
test('account and month filtering change trades without clipping later valuations',()=>{
 const d=data([['2026-05-01',100],['2026-06-01',101],['2026-06-08',102],['2026-07-01',103]]);
 const idx=M.index([flow('2026-05-01','买入'),flow('2026-06-01','卖出'),flow('2026-06-01','卖出',5,100,100,'DEMO_BOND','银行间','OTHER_DEMO')],d);
 const s=M.summarize(idx,'DEMO_ACCOUNT','2026-06');assert.equal(s.tradeCount,1);assert.equal(s.groups[0].buy,0);assert.equal(s.groups[0].sales[0].month.date,'2026-07-01');assert.equal(s.groups[0].sales[0].month.status,'complete');
});
test('uncovered bonds remain in the table and lower coverage instead of disappearing',()=>{
 const d=data([['2026-06-01',100]]),s=M.summarize(M.index([flow('2026-06-01','买入'),flow('2026-06-01','卖出',.03,100,100,'OTHER_DEMO_BOND')],d));
 assert.equal(s.groups.length,2);assert.equal(s.coveredCount,1);assert(Math.abs(s.coveredAmount/s.amount-.25)<1e-12);assert.equal(s.groups.find(g=>g.code==='OTHER_DEMO_BOND').week.full.return,null);
});
test('display grouping preserves all raw same-day same-price transactions',()=>{
 const d=data([['2026-06-01',100]]),g=first([flow('2026-06-01','买入'),flow('2026-06-01','买入')],d),marks=M.marks(g);assert.equal(marks.length,1);assert.equal(marks[0].count,2);assert.equal(marks[0].amount,.02);assert.equal(g.count,2);
});
test('fifth-slot parsing accepts quote-only and mixed-information workbooks',()=>{
 const c=vm.createContext({Date,console,setTimeout,clearTimeout,document:{querySelectorAll:()=>[],getElementById:()=>({addEventListener(){}})},window:{addEventListener(){}},XLSX:{utils:{decode_range:()=>({s:{r:0}}),sheet_to_json:ws=>ws.rows}}});
 new vm.Script(html.match(/<script>([\s\S]*?)<\/script>\s*<\/body>/)[1]).runInContext(c);
 c.workbook={SheetNames:['策略','债券','基金'],Sheets:{策略:{rows:[['项目编号','SAC账户名称','静态收益率'],['DEMO_ACCOUNT','虚构账户',2]]},债券:{rows:sheet([['2026-06-01',100]])},基金:{rows:sheet([['2026-06-01',1.2]],['000123.OF'],['虚构基金'])}}};
 const part=vm.runInContext('parseAcctInfoWb(workbook,"fictional-info.xlsx")',c);assert.equal(part.strategy.length,1);assert.equal(part.marketParts.length,2);assert.equal(part.found.filter(x=>x.type==='market').length,2);assert(part.found.some(x=>x.label.includes('基金单位净值')));
 assert(html.includes('&&!acctInfo.marketParts.length&&!acctInfo.indexParts.length)acctInfo=null;'));
});
test('module adds no networking or persistent business storage',()=>{
 const ui=html.match(/\/\/ BOND_MARKET_UI_BEGIN([\s\S]*?)\/\/ BOND_MARKET_UI_END/)[1];
 assert(!/\bfetch\s*\(|XMLHttpRequest|WebSocket|sendBeacon|localStorage|sessionStorage|indexedDB|setWidgetState/.test(moduleSource+ui));assert(html.includes('每页15条'));assert(html.includes('trading-analyzer-release" content="2026.10.29.2'));
});

const fundData=(rows,codes=['000123.OF'],names=['虚构基金'])=>M.merge([M.parseSheet(sheet(rows,codes,names),'fictional-quotes.xlsx','基金')]);
const fundFlow=(date,direction,amount=.01,price=1,shares=100,code='000123',mkt='场外市场')=>({...flow(date,direction,amount,price,shares,code,mkt),_variety:'基金买卖',_name2:'虚构基金','净价(元)':999});
test('fund unit NAV sheets support OF, SH and SZ markets without losing leading zeros',()=>{
 const d=fundData([['2026-06-01',1.2,.8,2]],['000123.OF','DEMO_ETF.SH','DEMO_REIT.SZ']);
 assert.equal(d.stats.fundCount,3);assert.equal(d.stats.bondCount,0);assert(d.series.has('000123|OF|fund'));
 assert.equal(M.identity(123,'场外市场','fund').code,'000123');assert.equal(M.identity('000123.OF').market,'OF');
 for(const s of d.series.values())assert.equal(s.kind,'fund');
});
test('fund trades match NAV using ledger full price and ignore deliberately incorrect net price',()=>{
 const d=fundData([['2026-06-01',1.1],['2026-06-08',1.2],['2026-07-01',1.3]]),g=first([fundFlow('2026-06-01','卖出',.01,1)],d);
 assert.equal(g.price,1);assert.equal(g.kind,'fund');assert(Math.abs(g.week.full.return-.2)<1e-12);assert(Math.abs(g.month.full.return-.3)<1e-12);assert.equal(M.marks(g)[0].price,1);
});
test('fund identifier suffixes are preserved rather than stripped into a different share class',()=>{
 const d=fundData([['2026-06-01',1.2]],['001234!1.OF']);assert(d.series.has('001234!1|OF|fund'));
 assert.equal(M.resolve(d,M.identity('001234','场外市场','fund'),'fund').series,null);assert.equal(M.resolve(d,M.identity('001234!1.OF'),'fund').series.quotes[0][1],1.2);
});
test('missing fund full price never falls back to net price, unit NAV or an amount-derived price',()=>{
 const d=fundData([['2026-06-01',1],['2026-06-08',1.2]]),g=first([fundFlow('2026-06-01','卖出',.01,null)],d);
 assert.equal(g.price,null);assert.equal(g.sell,.01);assert.equal(g.week.full.return,null);assert.equal(M.marks(g).length,0);assert(g.week.reasons.includes('缺卖出全价'));
});
test('fund sell price is share weighted and follow-up is weighted by sold amount',()=>{
 const d=fundData([['2026-06-01',1],['2026-06-02',2],['2026-06-08',1.2],['2026-06-09',2.2]]);
 const g=first([fundFlow('2026-06-01','卖出',.01,1,100),fundFlow('2026-06-02','卖出',.06,2,300)],d);
 assert.equal(g.price,1.75);assert(Math.abs(g.week.full.return-(.01*.2+.06*.1)/.07)<1e-12);
});
test('same code and market in bond and fund sheets cannot cross-match or merge trades',()=>{
 const bonds=M.parseSheet(sheet([['2026-06-01',101]],['DEMO_SAME.SH']),'fictional.xlsx','债券');
 const funds=M.parseSheet(sheet([['2026-06-01',1.1]],['DEMO_SAME.SH']),'fictional.xlsx','基金');
 const d=M.merge([bonds,funds]),s=M.summarize(M.index([flow('2026-06-01','买入',.01,100,100,'DEMO_SAME','上交所'),fundFlow('2026-06-01','买入',.02,1,100,'DEMO_SAME','上交所')],d));
 assert.equal(s.groups.length,2);assert.equal(s.groups.find(g=>g.kind==='fund').series.quotes[0][1],1.1);assert.equal(s.groups.find(g=>g.kind==='bond').series.quotes[0][1],101);
 const fundOnly=M.merge([funds]);assert.equal(M.resolve(fundOnly,M.identity('DEMO_SAME','上交所'),'bond').series,null);
});
test('asset filtering uses separate caches while preserving account and month scope',()=>{
 const idx=M.index([flow('2026-06-01','买入'),fundFlow('2026-06-01','买入',.02)],M.merge());
 assert.equal(M.summarize(idx).groups.length,2);assert.equal(M.summarize(idx,'','all','fund').groups.length,1);assert.equal(M.summarize(idx,'','all','bond').amount,.01);assert.equal(M.summarize(idx,'','2026-05','fund').groups.length,0);
});
test('a later fund sheet cutoff does not make a bond observation complete',()=>{
 const b=M.parseSheet(sheet([['2026-06-01',100],['2026-06-07',101]]),'fictional.xlsx','债券');
 const f=M.parseSheet(sheet([['2026-06-01',1],['2026-06-08',1.01]],['000123.OF']),'fictional.xlsx','基金');
 const d=M.merge([b,f]),idx=M.index([flow('2026-06-01','卖出'),fundFlow('2026-06-01','卖出')],d),s=M.summarize(idx);
 assert.equal(s.groups.find(g=>g.kind==='bond').sales[0].week.status,'partial');assert.equal(s.groups.find(g=>g.kind==='bond').sales[0].week.cutoff,'2026-06-07');assert.equal(s.groups.find(g=>g.kind==='fund').sales[0].week.status,'complete');
});
test('flat unit NAV is retained as a price change and labelled without implying total return',()=>{
 const d=fundData([['2026-06-01',1],['2026-06-08',1]]),g=first([fundFlow('2026-06-01','卖出')],d);
 assert.equal(g.week.full.return,0);assert(html.includes('零净值变化不代表零投资收益'));assert(html.includes('单位净值变动未加入分红和收益分配'));
});
