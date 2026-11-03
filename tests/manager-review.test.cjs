// All accounts, managers, prices and transactions in these tests are fictional.
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const html=fs.readFileSync(path.join(__dirname,'..','交易数据分析平台.html'),'utf8'),extract=n=>html.match(new RegExp('// '+n+'_BEGIN([\\s\\S]*?)// '+n+'_END'))[1];
const c=vm.createContext({Date});new vm.Script(['BOND_MARKET','INDEX_REPLAY','MANAGER_REVIEW','BUSINESS_VIEWS'].map(extract).join('\n')).runInContext(c);
const I=vm.runInContext('IndexReplay',c),R=vm.runInContext('ManagerReview',c),V=vm.runInContext('BusinessViews',c);
const dates=[];for(let i=0;dates.length<110;i++){const d=new Date(Date.UTC(2026,0,5+i));if(![0,6].includes(d.getUTCDay()))dates.push(d.toISOString().slice(0,10));}
const price=i=>i<20?110-i*.5:i<40?100+(i-20)*.5:i<60?110-(i-40):i<80?90+(i-60)*.75:105-(i-80)*.4;
function fixture(trades=[],options={}){
 const codes=[...new Set(trades.map(t=>t.a||'DEMO_A'))],dim={},pos=[],bonds=[],flows=[];
 for(const a of codes){dim[a]={line:'虚构部门',managers:options.managers?.[a]||['演示甲'],sacName:'虚构账户'+a};for(const d of [dates[0],dates[79]])pos.push({_code:a,_dateStr:d,_size:(options.sizes?.[a]||1)*1e8});bonds.push({acct:a,code:'DEMO_BOND',name:'虚构转债',bondType:'可转债',market:'银行间',mv:1000,face:1000,remain:2});}
 for(const [i,t]of trades.entries())flows.push({_code:t.a||'DEMO_A',_dateStr:dates[t.i],_variety:t.variety||'现券买卖',_dir:t.dir||'买入',_amountYi:t.amount||.05,'交易面额(万)':t.face===null?null:(t.face??t.amount??.05)*1e4,'全价(元)':100,代码:'DEMO_BOND',名称:'虚构转债',_mkt:'银行间','交易编号':'DEMO_'+i,_source:{file:'fictional.xlsx',sheet:'流水',row:i+2}});
 const qs=dates.map((d,i)=>[d,price(i)]),part={calendar:dates,series:I.CATEGORIES.map(x=>({key:x.key,code:x.code,name:x.name,quotes:qs}))};
 return I.build({_store:{hg:codes.map(_code=>({_code})),zqAll:flows,pos,hdStart:{dateStr:dates[0],bonds,funds:[]},hdEnd:{dateStr:dates[79],bonds,funds:[]}},acctDim:{dim,strat:options.strat||{}},indexMarket:I.merge([part])});
}
const key='虚构部门|演示甲',report=(trades,options)=>R.report(R.build(fixture(trades,options)),key);
test('trade shares use cash amounts and co-managed accounts use 1/N once',()=>{
 const r=report([{i:20,a:'DEMO_A',amount:.2},{i:20,a:'DEMO_B',amount:.1}],{managers:{DEMO_A:['演示甲','演示乙'],DEMO_B:['演示甲']}});assert(Math.abs(r.amount-.2)<1e-10);assert(r.profiles.every(a=>Math.abs(a.share-.5)<1e-10));assert.equal(r.accountCount,2);assert(r.introduction.includes('50.0%')||r.primary[0].text.includes('50.0%'));
});
test('valid unclassified transactions remain in account shares without invented benchmark events',()=>{
 const idx=fixture([{i:20,a:'DEMO_A',amount:.2},{i:20,a:'DEMO_B',amount:.05}]),a=idx.accounts.get('DEMO_A');a.unknown=a.events.map(e=>({...e,bucket:'',recordKind:'flow'}));a.events=[];const r=R.report(R.build(idx),key);assert(Math.abs(r.amount-.25)<1e-10);assert(Math.abs(r.profiles.find(p=>p.code==='DEMO_A').share-.8)<1e-10);assert(r.events.every(e=>e.account!=='DEMO_A'));
});
test('market ranks only use prices known on the operation date',()=>{
 const short=dates.slice(0,25).map((d,i)=>[d,price(i)]),long=[...short,...dates.slice(25).map((d,i)=>[d,10000+i])],a=R.marketFeatures(short),b=R.marketFeatures(long);assert.equal(a.ranks.get(dates[20]).rank,b.ranks.get(dates[20]).rank);assert.equal(a.ranks.get(dates[5]),undefined);
});
test('repeated buying near a trailing low is supported by positive later market',()=>{
 const r=report([18,19,20].map(i=>({i,amount:.05}))),e=r.events.find(e=>e.quality==='favorable'&&e.direction===1);assert(e);assert.equal(e.location,'low');assert(e.repeated);assert(e.text.includes('持续加仓'));assert(e.text.includes('指数'));assert(!e.text.includes('超额收益'));
});
test('selling near a stage high followed by falling market is favorable timing',()=>{
 const r=report([37,38,39].map(i=>({i,dir:'卖出',amount:.05}))),e=r.events.find(e=>e.quality==='favorable'&&e.direction===-1);assert(e);assert(e.month.value<0);assert(e.text.includes('减持后的市场走势回落'));
});
test('opposite direction requires two thirds of at least three actual participants',()=>{
 const trades=['DEMO_A','DEMO_B','DEMO_C'].flatMap(a=>[37,38,39].map(i=>({i,a,dir:a==='DEMO_A'?'买入':'卖出',amount:.05}))),options={managers:{DEMO_A:['演示甲'],DEMO_B:['演示乙'],DEMO_C:['演示丙']}};
 const r=report(trades,options),e=r.events.find(e=>e.peer?.divergent);assert(e);assert.equal(e.peer.rows.length,3);assert.equal(e.peer.sell,2);assert(e.text.includes('3名参与经理中有2名净卖出'));
 const two=report(trades.filter(t=>t.a!=='DEMO_C'),options);assert(two.events.every(e=>!e.peer?.divergent));
});
test('majority does not require majority managers to dominate aggregate amount',()=>{
 const trades=['DEMO_A','DEMO_B','DEMO_C'].flatMap(a=>[37,38,39].map(i=>({i,a,dir:a==='DEMO_A'?'买入':'卖出',amount:a==='DEMO_A'?.5:.05}))),r=report(trades,{managers:{DEMO_A:['演示甲'],DEMO_B:['演示乙'],DEMO_C:['演示丙']},sizes:{DEMO_A:2}});assert(r.events.some(e=>e.peer?.divergent));
});
test('missing face prevents invented net directions but retains cash shares',()=>{
 const r=report([18,19,20].map(i=>({i,amount:.05,face:null})));assert.equal(r.coverage.missingFace,3);assert.equal(r.events.length,0);assert(Math.abs(r.amount-.15)<1e-10);
});
test('flat prices do not create false high-low timing claims',()=>{
 const idx=fixture([18,19,20].map(i=>({i,amount:.05})));for(const s of idx.market.series.values())s.quotes=s.quotes.map(q=>[q[0],100]);const r=R.report(R.build(idx),key);assert.equal(r.events.length,0);
});
test('ordinary selling into a steadily rising wealth index is not automatically a bad trade',()=>{
 const idx=fixture([37,38,39].map(i=>({i,dir:'卖出',amount:.05})));for(const s of idx.market.series.values())s.quotes=s.quotes.map((q,i)=>[q[0],100+i*.1]);const r=R.report(R.build(idx),key);assert(r.events.every(e=>e.quality!=='adverse'));assert(r.events.every(e=>!e.text.includes('减持较早')));
});
test('near-zero scale outliers do not dominate performance prose or change original cash totals',()=>{
 const r=report([18,19,20].map(i=>({i,amount:.05})),{sizes:{DEMO_A:.00001}});assert.equal(r.events.length,0);assert(Math.abs(r.amount-.15)<1e-10);
});
test('one-month dates clamp to month end and incomplete windows stay labelled',()=>{
 assert.equal(R.plusMonth('2026-01-31'),'2026-02-28');const q=[['2026-01-01',100],['2026-01-08',101]];const v=R.forward(q,'2026-01-01','2026-02-01');assert.equal(v.complete,false);assert.equal(v.date,'2026-01-08');assert.equal(R.forward(q,'2026-01-08','2026-02-01'),null);
});
test('main account transfers compare before and after shares and verify new-account dates',()=>{
 const trades=[...Array.from({length:8},(_,i)=>({i:8+i,a:'DEMO_OLD',amount:.2})),...Array.from({length:12},(_,i)=>({i:45+i,a:'DEMO_NEW',amount:.2}))],r=report(trades,{strat:{DEMO_NEW:{startDate:dates[40]}}});assert(r.shifts.length);assert(r.shifts.some(s=>s.text.includes('交易重心由')));assert(r.shifts.every(s=>s.share>=.35));
});
test('unrelated managers and departments never leak into peer tables',()=>{
 const idx=fixture(['DEMO_A','DEMO_B','DEMO_C','DEMO_D'].flatMap(a=>[{a,i:37,dir:a==='DEMO_A'?'买入':'卖出',amount:.1}]),{managers:{DEMO_A:['演示甲'],DEMO_B:['演示乙'],DEMO_C:['演示丙'],DEMO_D:['演示丁']}});idx.accounts.get('DEMO_D').dept='其他虚构部门';idx.managers.find(m=>m.name==='演示丁').dept='其他虚构部门';const r=R.report(R.build(idx),key);assert(r.events.filter(e=>e.peer).every(e=>e.peer.rows.every(p=>p.name!=='演示丁')));
});
test('small accounts can contribute conspicuous operations even with little manager volume',()=>{
 const trades=[];for(let a=0;a<6;a++)for(let i=5;i<65;i+=3)for(const dir of ['买入','卖出'])trades.push({i,a:'DEMO_BIG'+a,dir,amount:.1});for(const i of [18,19,20])trades.push({i,a:'DEMO_TINY',amount:.03});const r=report(trades);assert(r.profiles.find(p=>p.code==='DEMO_TINY').share<.01);assert(r.events.some(e=>e.account==='DEMO_TINY'));assert(r.events.find(e=>e.account==='DEMO_TINY').text.includes('虽然仅占'));
});
test('no full-market API, persistence or saved business output is added',()=>{
 const source=extract('MANAGER_REVIEW')+extract('MANAGER_REVIEW_UI')+extract('BUSINESS_VIEWS');assert(!/\bfetch\s*\(|XMLHttpRequest|WebSocket|sendBeacon|localStorage|sessionStorage|indexedDB/.test(source));assert(html.includes("connect-src 'none'"));
});
test('repo daily sums keep all six components without netting and keep accounts apart',()=>{
 const rows=V.DIRS.flatMap((d,i)=>V.TYPES.map((t,j)=>({_code:'DEMO_A',_name:'虚构回购账户',_dateStr:'2026-01-06',_dir:d,_biz:t,_amountYi:(i+1)*10+j})));rows.push({...rows[0],_amountYi:2},{...rows[0],_code:'DEMO_B',_amountYi:99});const m=V.repoDaily(rows),day=m.accounts.get('DEMO_A').days.get('2026-01-06');assert.equal(day['正回购|协议回购'],12);assert.equal(day['逆回购|通用质押式回购'],22);assert.equal(m.accounts.get('DEMO_B').amount,99);assert.equal(Object.keys(day).length,6);
});
test('comparison sorting is numerical and leaves nulls last in both directions',()=>{
 const rows=[{values:['演示甲',2]},{values:['演示乙',10]},{values:['演示丙',null]}];assert.deepEqual(Array.from(V.sortRows(rows,1,false),r=>r.values[1]),[10,2,null]);assert.deepEqual(Array.from(V.sortRows(rows,1,true),r=>r.values[1]),[2,10,null]);assert.equal(rows[0].values[1],2);
});
