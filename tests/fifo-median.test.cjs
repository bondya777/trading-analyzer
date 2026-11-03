// Fictional FIFO records only. No embedded business inputs.
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const html=fs.readFileSync(path.join(__dirname,'..','交易数据分析平台.html'),'utf8');
const app=vm.createContext({Date,console,setTimeout,clearTimeout,document:{querySelectorAll:()=>[],getElementById:()=>({addEventListener(){}})},window:{addEventListener(){}}});
new vm.Script(html.match(/<script>([\s\S]*?)<\/script>\s*<\/body>/)[1]).runInContext(app);
const build=vm.runInContext('buildFifo',app),kind='现券买卖';
function trades(returns,qty=returns.map(()=>1)){
 return returns.flatMap((r,i)=>[{_code:'DEMO_A',代码:'DEMO_BOND_'+i,_variety:kind,_dir:'买入',_date:new Date('2026-01-06'),_name2:'虚构债券','交易面额(万)':qty[i],'全价(元)':100},{_code:'DEMO_A',代码:'DEMO_BOND_'+i,_variety:kind,_dir:'卖出',_date:new Date('2026-01-07'),_name2:'虚构债券','交易面额(万)':qty[i],'全价(元)':100*(1+r/100)}]);
}
test('odd FIFO median is the middle return rather than the cash-weighted average',()=>{
 const a=build(trades([1,2,10],[1000,1,1]),{bonds:[]}).accounts[0];assert.equal(a.medianRet,2);assert.notEqual(a.medianRet,a.avgRet);assert.equal(a.closed,3);assert.equal(a.winRate,100);
});
test('even FIFO median averages the middle two values including negative results',()=>{
 const a=build(trades([-1,1,3,50],[500,50,5,1]),{bonds:[]}).accounts[0];assert.equal(a.medianRet,2);assert.equal(a.closed,4);assert.equal(a.wins,3);assert.equal(a.winRate,75);
});
test('FIFO matching segments remain the observation unit when one sale closes several lots',()=>{
 const rows=trades([1,2,3]);rows[1]['交易面额(万)']=.5;rows.push({...rows[1],_date:new Date('2026-01-08'),'交易面额(万)':.5,'全价(元)':110});const a=build(rows,{bonds:[]}).accounts[0];assert.equal(a.closed,4);assert.equal(a.medianRet,2.5);
});
test('opening inventory uses supplied full-price cost and unclosed purchases do not enter the median',()=>{
 const rows=trades([1,3]);rows.push({_code:'DEMO_A',代码:'DEMO_START',_variety:kind,_dir:'卖出',_date:new Date('2026-01-08'),'交易面额(万)':1,'全价(元)':102},{_code:'DEMO_A',代码:'DEMO_OPEN',_variety:kind,_dir:'买入',_date:new Date('2026-01-08'),'交易面额(万)':1,'全价(元)':1});const a=build(rows,{bonds:[{acct:'DEMO_A',code:'DEMO_START',name:'虚构期初债券',face:1,costFull:100}]}).accounts[0];assert.equal(a.medianRet,2);assert.equal(a.closed,3);
});
test('the existing fewer-than-three-pairs filter and ranking order are preserved',()=>{
 assert.equal(build(trades([1,2]),{bonds:[]}).accounts.length,0);const a=trades([1,2,10]).concat(trades([-1,-2,-3,-4]).map(r=>({...r,_code:'DEMO_B'}))),r=build(a,{bonds:[]});assert.equal(r.accounts[0].code,'DEMO_B');assert.equal(r.accounts[0].closed,4);assert.equal(r.accounts[0].medianRet,-2.5);
});
test('both FIFO display locations name and use the median consistently',()=>{
 assert(!html.includes('平均往返收益'));assert(html.includes('往返收益中位数(%)'));assert(html.includes('ff.medianRet'));assert(html.includes('a.medianRet!=null?a.medianRet'));
});
test('sale-detail panel and manager-flow table are removed without removing market computations',()=>{
 const market=html.match(/\/\/ BOND_MARKET_UI_BEGIN([\s\S]*?)\/\/ BOND_MARKET_UI_END/)[1],manager=html.match(/\/\/ MANAGER_REVIEW_UI_BEGIN([\s\S]*?)\/\/ MANAGER_REVIEW_UI_END/)[1];assert(!market.includes('marketReplayDetails'));assert(!market.includes('renderMarketReplayDetails'));assert(!manager.includes('对应流水'));assert(manager.includes('查看指标依据'));assert(manager.includes('后续市场'));assert(manager.includes('同部门同区间'));assert(market.includes('导出Excel'));assert(market.includes('买卖金额与卖出后价格'));
});
