// All names, codes, quotes and amounts below are fictional.
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const html=fs.readFileSync(path.join(__dirname,'..','交易数据分析平台.html'),'utf8');
const extract=name=>html.match(new RegExp('// '+name+'_BEGIN([\\s\\S]*?)// '+name+'_END'))[1];
const c=vm.createContext({Date});new vm.Script(['BOND_MARKET','INDEX_REPLAY','SUBJECT_REPLAY'].map(extract).join('\n')).runInContext(c);
const I=vm.runInContext('IndexReplay',c),M=vm.runInContext('BondMarket',c),S=vm.runInContext('SubjectReplay',c),first='2026-01-05',last='2026-01-30',all={level:'department',dept:'all'};
const near=(a,b)=>assert(Math.abs(a-b)<1e-11,`${a} != ${b}`);
const ref=rows=>S.merge([S.parseSheet([['代码','名称','债务主体'],...rows],'fictional.xlsx','债务主体')]);
const data=ref([['DEMO_1','虚构债券甲','虚构主体甲'],['DEMO_2','虚构债券乙','虚构主体甲'],['DEMO_3','虚构债券丙','虚构主体乙']]);
const bond=(acct,code,face,mv=face)=>({acct,code,name:code==='DEMO_1'?'虚构债券甲':'虚构债券乙',market:'银行间',face:face*1e4,mv:mv*1e4,bondType:'公司债',remain:3});
const flow=(acct,date,dir,face,code='DEMO_1',variety='现券买卖')=>({_code:acct,_name:'虚构账户'+acct,_dateStr:date,_dir:dir,_variety:variety,_amountYi:face*1.01,'交易面额(万)':face*1e4,'全价(元)':101,代码:code,_name2:code==='DEMO_1'?'虚构债券甲':'虚构债券乙',_mkt:'银行间'});
function fixture(){
 const flows=[flow('DEMO_A','2026-01-06','买入',.05),flow('DEMO_B','2026-01-07','分销买入',.02,'DEMO_2','分销买卖'),flow('DEMO_C','2026-01-08','卖出',.02),flow('DEMO_A','2026-01-09','债券回售',.01,'DEMO_2','债券回售')];
 return{_store:{hg:['DEMO_A','DEMO_B','DEMO_C'].map(_code=>({_code})),zqAll:flows,hdStart:{dateStr:first,bonds:[bond('DEMO_A','DEMO_1',.1),bond('DEMO_A','DEMO_2',.2),bond('DEMO_B','DEMO_1',.05),bond('DEMO_C','DEMO_1',.06)],funds:[]},hdEnd:{dateStr:last,bonds:[bond('DEMO_A','DEMO_1',.15,.151),bond('DEMO_A','DEMO_2',.2),bond('DEMO_B','DEMO_1',.05),bond('DEMO_B','DEMO_2',.02),bond('DEMO_C','DEMO_1',.04)],funds:[]},pos:['DEMO_A','DEMO_B','DEMO_C'].flatMap((a,i)=>[first,last].map(d=>({_code:a,_dateStr:d,_size:(i+2)*1e8})))},acctDim:{dim:{DEMO_A:{line:'演示投顾部',sacName:'虚构账户甲',managers:['演示经理甲','演示经理乙']},DEMO_B:{line:'演示投顾部',sacName:'虚构账户乙',managers:['演示经理丙']},DEMO_C:{line:'演示财富部',sacName:'虚构账户丙',managers:['演示经理丁']}},strat:{}},indexMarket:I.merge(),bondMarket:M.merge()};
}
const issuer=(idx,state=all,name='虚构主体甲')=>S.query(idx,data,state,{mode:'issuer',issuer:name});
const security={kind:'bond',code:'DEMO_1',market:'IB',key:'DEMO_1|IB',name:'虚构债券甲'};
test('reference parser preserves leading zeros, optional markets and source row offsets',()=>{
 const p=S.parseSheet([['说明'],['代码','名称','债务主体','市场'],['032200001','虚构零头债','虚构主体','银行间'],['DEMO_1.SH','虚构甲','虚构主体']], 'fictional.xlsx','债务主体',8);
 assert.equal(p.rows[0].code,'032200001');assert.equal(p.rows[0].market,'IB');assert.equal(p.rows[0].source.row,11);assert.equal(p.rows[1].market,'SH');
});
test('price matrices and unrelated account references are not issuer mappings',()=>{
 assert.equal(S.parseSheet([['日期','虚构债'],['Date','DEMO_1.IB'],[first,100]],'fictional.xlsx','债券1'),null);
 assert.equal(S.parseSheet([['项目编号','名称'],['DEMO_A','虚构账户']],'fictional.xlsx','IAM'),null);
});
test('blank subjects and spreadsheet errors are not treated as actual issuers',()=>{
 const p=S.parseSheet([['代码','债务主体'],['DEMO_1','#N/A'],['DEMO_2',''],['#VALUE!','虚构主体'],['DEMO_3','虚构主体']],'fictional.xlsx','债务主体');assert.equal(p.rows.length,1);assert.equal(p.invalid,3);assert.equal(p.rows[0].code,'DEMO_3');
});
test('identical reference rows deduplicate and references do not mutate prices',()=>{
 const p=S.parseSheet([['代码','债务主体'],['DEMO_1','虚构主体']],'fictional.xlsx','对照');const r=S.merge([p,p]);assert.equal(r.stats.rows,2);assert.equal(r.stats.duplicates,1);assert.equal(r.stats.codes,1);
 assert.equal(S.resolve(r,{code:'DEMO_1',market:'IB',kind:'bond'}).issuer,'虚构主体');
});
test('ambiguous issuer relations remain unresolved without a matching supplied name',()=>{
 const r=ref([['DEMO_1','甲债','甲主体'],['DEMO_1','乙债','乙主体']]);assert.equal(S.resolve(r,{code:'DEMO_1',kind:'bond'}).issuer,'');assert.equal(S.resolve(r,{code:'DEMO_1',kind:'bond',name:'甲债'}).issuer,'甲主体');assert.equal(r.stats.conflicts,1);
});
test('explicit markets prevent cross-market mapping leakage and funds are excluded',()=>{
 const r=ref([['DEMO_1.IB','甲债','甲主体'],['DEMO_1.SH','乙债','乙主体']]);assert.equal(S.resolve(r,{code:'DEMO_1',market:'IB',kind:'bond'}).issuer,'甲主体');assert.equal(S.resolve(r,{code:'DEMO_1',market:'SZ',kind:'bond'}).issuer,'');assert.equal(S.resolve(r,{code:'DEMO_1',market:'IB',kind:'fund'}).issuer,'');
});
test('empty issuer does not combine unmapped securities',()=>{
 const idx=I.build(fixture()),q=S.query(idx,S.merge(),all,{mode:'issuer',issuer:''});assert.equal(q.records.length,0);assert.equal(q.at(last).holders,0);
});
test('issuer projection uses face and includes primary, puts, and cash separately',()=>{
 const q=issuer(I.build(fixture()));near(q.at(first).projected,.41);near(q.at(last).projected,.45);near(q.at(last).actual,.46);near(q.at(last).actualMV,.461);near(q.dailyCash.get('2026-01-06').buy,.0505);near(q.dailyCash.get('2026-01-07').buy,.0202);near(q.dailyCash.get('2026-01-09').repay,.0101);assert.equal(q.securities.size,2);
});
test('selected date changes projected holders without changing period-end actuals',()=>{
 const q=issuer(I.build(fixture())),before=q.at(first),after=q.at(last);near(before.actual,after.actual);near(before.actualMV,after.actualMV);near(before.projected,.41);near(q.at('2026-01-06').projected,.46);assert.equal(before.date,first);
});
test('individual holders include accounts with no period trades',()=>{
 const idx=I.build(fixture()),q=S.query(idx,data,all,{mode:'security',security});assert.equal(q.at(first).holders,3);near(q.at(last).rows.find(r=>r.account==='DEMO_B').projected,.05);near(q.at(last).projected,.24);
});
test('department and account scopes never expose outside holders',()=>{
 const idx=I.build(fixture()),d=issuer(idx,{level:'department',dept:'演示投顾部'}),a=issuer(idx,{level:'account',dept:'演示投顾部',account:'DEMO_A'});
 assert.equal(d.at(last).rows.length,2);assert(!d.at(last).rows.some(r=>r.account==='DEMO_C'));near(d.at(last).projected,.41);near(d.at(last).ratio,.41/5);assert.equal(a.at(last).rows.length,1);near(a.at(last).projected,.34);
});
test('manager scope splits all holdings, cash and account scale by one over N exactly once',()=>{
 const idx=I.build(fixture()),state={level:'manager',dept:'演示投顾部',manager:'演示投顾部|演示经理甲'},q=issuer(idx,state),snap=q.at(last);
 near(snap.projected,.17);near(snap.actual,.175);near(snap.netSize,1);near(q.dailyCash.get('2026-01-06').buy,.02525);near(q.distribution(last,'manager')[0].value,.17);assert.equal(snap.rows[0].managers.length,2);
});
test('manager distribution conserves department projection and account ratio is account-specific',()=>{
 const q=issuer(I.build(fixture()),{level:'department',dept:'演示投顾部'}),snap=q.at(last);near(q.distribution(last,'manager').reduce((n,r)=>n+r.value,0),snap.projected);near(snap.rows.find(r=>r.account==='DEMO_A').ratio,.34/2);
});
test('account bond details and individual distribution reconcile to the query total',()=>{
 const q=issuer(I.build(fixture()));for(const r of q.at(last).rows){const ds=q.details(r.account,last);near(ds.reduce((n,r)=>n+r.projected,0),r.projected);near(ds.reduce((n,r)=>n+r.actual,0),r.actual);}near(q.distribution(last,'security').reduce((n,r)=>n+r.value,0),q.at(last).projected);
});
test('a zero selected holding and positive actual holding remains visible and is not called missing',()=>{
 const r=fixture();r._store.hdStart.bonds=r._store.hdStart.bonds.filter(h=>h.acct!=='DEMO_B');r._store.zqAll=r._store.zqAll.filter(e=>e._code!=='DEMO_B');const snap=issuer(I.build(r)).at(first),b=snap.rows.find(r=>r.account==='DEMO_B');assert.equal(b.projected,0);near(b.actual,.07);assert.equal(snap.holders,2);assert.equal(snap.actualHolders,3);
});
test('missing face is unavailable rather than zero while actual market value remains visible',()=>{
 const r=fixture();r._store.hdStart.bonds[0].faceKnown=false;r._store.hdEnd.bonds[0].faceKnown=false;const snap=issuer(I.build(r)).at(last);assert.equal(snap.projected,null);assert.equal(snap.actual,null);near(snap.actualMV,.461);assert.equal(snap.rows[0].projected==null||snap.rows.some(r=>r.projected==null),true);
});
test('an entirely absent opening snapshot does not produce a false zero projection',()=>{
 const r=fixture();r._store.hdStart={dateStr:'',bonds:[],funds:[]};const q=issuer(I.build(r));assert.equal(q.at(last).projected,null);assert(q.history.every(r=>r.amount==null));
});
test('arbitrary dates use only prior operations and prior scale, without reading future data',()=>{
 const q=issuer(I.build(fixture()));near(q.at('2026-01-10').projected,.45);assert.equal(q.at('2026-02-10').date,last);assert.equal(q.at('2025-12-31').date,first);
});
test('negative reconstructed holdings are retained and not labeled as positive holders',()=>{
 const r=fixture();r._store.zqAll.push(flow('DEMO_A','2026-01-15','卖出',.5));const q=S.query(I.build(r),data,{level:'account',dept:'all',account:'DEMO_A'},{mode:'security',security}),snap=q.at(last);assert(snap.projected<0);assert.equal(snap.holders,0);assert.equal(snap.negative,1);assert.equal(snap.rows.length,1);
});
test('fund negative-only holdings and trades over 30 percent remain excluded',()=>{
 const r=fixture();const f={acct:'DEMO_A',code:'DEMO_FUND',name:'虚构基金',market:'上交所',mv:1000};r._store.hdStart.funds=[f,{...f,code:'DEMO_NEG',mv:-1000}];r._store.hdEnd.funds=[{...f,mv:1200}];r._store.zqAll.push({...flow('DEMO_A','2026-01-10','买入',.7,'DEMO_FUND','基金买卖'),_mkt:'上交所','全价(元)':1});
 const idx=I.build(r),q=S.query(idx,data,all,{mode:'security',security:{kind:'fund',code:'DEMO_FUND',market:'SH',key:'DEMO_FUND|SH|fund'}});near(q.at(last).projected,.1);near(q.at(last).actual,.12);assert.equal(q.events.length,0);assert.equal(S.options(idx,data,all).subjects.length,1);
});
test('cross-market fund holdings follow existing pooling rules',()=>{
 const r=fixture(),f={acct:'DEMO_A',code:'DEMO_POOL',name:'虚构跨市场基金',kind:'基金'};r._store.hdStart.funds=[{...f,market:'上交所',mv:-1000},{...f,market:'场外市场',mv:3000}];r._store.hdEnd.funds=r._store.hdStart.funds;
 const q=S.query(I.build(r),data,all,{mode:'security',security:{kind:'fund',code:'DEMO_POOL',market:'SH',key:'DEMO_POOL|SH|fund'}});near(q.at(last).projected,.2);near(q.at(last).actual,.2);assert.equal(q.at(last).rows[0].crossMarket,true);
});
test('query caches isolate manager selection and distinct issuer and security targets',()=>{
 const idx=I.build(fixture()),a=issuer(idx),b=issuer(idx,{level:'account',dept:'all',account:'DEMO_A'}),other=issuer(idx,all,'虚构主体乙');assert.notEqual(a,b);assert.equal(other.at(last).holders,0);near(a.at(last).projected,.45);assert.equal(a,issuer(idx));
});
test('zero-only dormant snapshot rows do not inflate issuer or unmatched-bond counts',()=>{
 const r=fixture();r._store.hdEnd.bonds.push(bond('DEMO_A','DEMO_3',0),bond('DEMO_A','DEMO_UNMAPPED',0));const idx=I.build(r),options=S.options(idx,data,all);assert.equal(options.subjects.length,1);assert.equal(options.unmapped.length,0);assert.equal(issuer(idx,all,'虚构主体乙').securities.size,0);
});
test('new model and UI have no network APIs or persistent business storage',()=>{
 const source=extract('SUBJECT_REPLAY')+extract('SUBJECT_REPLAY_UI');assert(!/\bfetch\s*\(|XMLHttpRequest|WebSocket|sendBeacon|localStorage|sessionStorage|indexedDB/.test(source));assert(html.includes("connect-src 'none'"));
});
