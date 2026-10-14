import {test} from 'node:test';
import assert from 'node:assert/strict';
import {loadActions,implOf} from './helpers/actions.mjs';
function fixture({supported=true,failSource=false}={}){
 const events=[];let runs=0,phase=0,adds=0;
 const chart={name:'Chart 1',title:{text:'',visible:false,load(){}},load(){events.push('chart-load');},setPosition(a,b){events.push(['position',a,b]);}};
 for(const [key,value]of Object.entries({left:0,top:300,width:600,height:280}))Object.defineProperty(chart,key,{get(){assert.ok(phase>=2,'geometry must be synchronized after creation');return value;}});
 const sheet={name:'Generated_Overview',load(){events.push('sheet-load');},charts:{items:[{name:'Earlier',left:0,top:0,width:200,height:100}],load(){events.push('charts-load');},add(type,data,by){assert.ok(phase>=1,'source must preflight before any chart creation');assert.equal(data,source);assert.equal(by,'Auto');adds++;events.push(['add',type]);return chart;}}};
 const source={worksheet:sheet,address:'Generated_Overview!A14:C18',rowCount:5,columnCount:3,values:[['类别','预算','实际'],['获客',20000,18000]],formulas:[['','',''],['','','']],numberFormat:[['General','0.00','0.00']],load(){events.push('source-load');},getCell(){return{getResizedRange(){return source;}};}};
 const util={okResult:result=>({ok:true,result}),errResult:(code,error)=>({ok:false,code,error}),excelRange:()=>source};
 const {A,sandbox}=loadActions({host:'Excel',supported,util});
 sandbox.Excel.run=async callback=>{runs++;return callback({workbook:{worksheets:{getActiveWorksheet(){throw Error('must not use unrelated active worksheet');}}},async sync(){events.push('sync');if(failSource)throw Error('invalid source');phase++;}});};
 return{run:args=>implOf(A,'Excel','insert_chart')(args),preview:args=>A.hosts.Excel.previewReaders.insert_chart(args),events,adds:()=>adds,runs:()=>runs};
}
test('chart is hosted on dataRange worksheet with native position and actual geometry',async()=>{
 const f=fixture();const out=await f.run({dataRange:'Generated_Overview!A14:C18',type:'column',title:'费用类别预算与实际（元）',position:{startCell:'$A$30',endCell:'H46'}});
 assert.equal(out.ok,true,out.error);assert.equal(f.adds(),1);assert.equal(out.result.chart.sheet,'Generated_Overview');assert.equal(out.result.chart.titleVisible,true);assert.equal(out.result.chart.geometry.height,280);assert.deepEqual(f.events.find(e=>Array.isArray(e)&&e[0]==='position'),['position','A30','H46']);
});
test('all invalid chart arguments fail before entering Excel or creating a chart',async()=>{
 for(const args of [{dataRange:1},{dataRange:'S!A1:B2',type:'unknown'},{dataRange:'S!A1:B2',title:12},{dataRange:'S!A1:B2',position:{startCell:'A1',endCell:'Other!B4'}},{dataRange:'S!A1:B2',position:{startCell:'B4',endCell:'A1'}},{dataRange:'S!A1:B2',position:{startCell:'A1',endCell:'XFE2'}},{dataRange:'S!A1:B2',position:{startCell:'A1',endCell:'B1048577'}}]){
  const f=fixture();assert.equal((await f.run(args)).code,'bad_args');assert.equal(f.runs(),0);assert.equal(f.adds(),0);
 }
});
test('invalid native source preflight prevents chart creation',async()=>{const f=fixture({failSource:true});assert.equal((await f.run({dataRange:'Missing!A1:B3'})).code,'execution');assert.equal(f.adds(),0);});
test('chart approval preview reads actual source and target-sheet charts without writes',async()=>{const f=fixture();const out=await f.preview({dataRange:'Generated_Overview!A14:C18',position:{startCell:'A30',endCell:'H46'}});assert.equal(out.ok,true,out.error);assert.equal(out.result.sheet,'Generated_Overview');assert.equal(out.result.existingCharts[0].name,'Earlier');assert.equal(out.result.source.values[1][1],20000);assert.equal(f.adds(),0);});
test('chart gate uses supported baseline and default column remains compatible',async()=>{const absent=fixture({supported:false});assert.equal((await absent.run({dataRange:'S!A1:B2'})).code,'requirement');assert.equal(absent.runs(),0);const f=fixture();assert.equal((await f.run({dataRange:'S!A1:B2'})).result.chart.type,'ColumnClustered');});
