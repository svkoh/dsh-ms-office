#!/usr/bin/env node
/** One native editable ten-slide atlas. Office.js plugin HTTP is the only writer.
 * Default: print plan. --inspect: native read-only preflight. --execute --allow-write:
 * run only AFTER the coordinating agent has preserved/approved the target document.
 * node scripts/create-ppt-style-atlas.mjs --inspect --instance inst-... --port 3000
 * node scripts/create-ppt-style-atlas.mjs --execute --allow-write --instance inst-... --port 3000 --run-id atlas-v1
 * Add --allow-default-size only after reviewing native read_presentation_info.
 * Native pptx/PNG bytes are transported unchanged. ZIP parsing below is READ ONLY.
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { inflateRawSync } from 'node:zlib';

export const DEMO = Object.freeze({ categories: ['研发','运营','营销','支持'], budget: [120,80,60,40], actual: [114,84,54,38], totalBudget: 300, totalActual: 290, balance: 10, execution: '96.7%', unit: '万元' });
export const STYLE = Object.freeze({ width: 960, height: 540, titleFont: 'SimHei', bodyFont: 'KaiTi', titleSize: 30, bodySize: 18, labelSize: 16, dataSize: 14, footerSize: 12 });
const C = { navy: '#17365D', blue: '#2E75B6', teal: '#008C95', gold: '#D49A27', red: '#B64646', ink: '#243746', muted: '#607486', pale: '#EDF3F8', white: '#FFFFFF', grid: '#D6E0E9' };
const text = (value,x,y,w,h,size=18,extra={}) => ({ type:'text',text:value,x,y,w,h,font:'KaiTi',size,color:C.ink,autoFit:false,lineWidth:0,...extra });
const rect = (x,y,w,h,fill=C.pale,extra={}) => ({ type:'shape',shapeType:'Rectangle',x,y,w,h,text:'',font:'KaiTi',size:18,fill,lineWidth:0,autoFit:false,...extra });
const line = (x,y,w,h=0,color=C.grid,extra={}) => ({ type:'line',x,y,w,h,color,lineWidth:1,connectorType:'Straight',...extra });
const title = (value,extra={}) => text(value,48,32,864,52,30,{font:'SimHei',color:C.navy,...extra});
const base = (name,index,dark=false) => [rect(0,0,960,540,dark?C.navy:C.white), ...(dark?[]:[rect(48,97,48,4,C.teal)]), title(name,dark?{color:C.white}:{}), line(48,492,864,0,dark?'#426183':C.grid), text('Office 原生样式图谱  ·  预算执行示例',48,502,790,22,12,{color:dark?'#D8E5F2':C.muted}), text(String(index).padStart(2,'0')+' / 10',838,502,74,22,12,{align:'Right',color:dark?'#D8E5F2':C.muted})];

export function buildAtlasPlan() {
  const slides = [];
  slides.push({name:'cover',title:'预算执行 · 原生样式图谱',elements:[...base('预算执行 · 原生样式图谱',1,true),rect(48,145,6,228,C.teal),text('用同一组业务数据\n展示十种可编辑表达',78,153,520,106,30,{font:'SimHei',color:C.white}),text('文字、表格、图形、图片与流程\n统一字体与视觉层级，支持后续编辑。',80,286,520,78,18,{color:'#D8E5F2'}),rect(663,156,230,213,'#244C78'),text('300',690,179,177,59,30,{font:'SimHei',color:C.white}),text('预算总额 / 万元',690,246,177,31,16,{color:'#D8E5F2'}),line(690,290,174,0,'#52779B'),text('实际 290  ·  结余 10',690,315,180,48,16,{color:C.white}),text('共享示例：研发 / 运营 / 营销 / 支持',80,411,780,30,16,{color:'#D8E5F2'})]});
  const kpi = [['预算总额','300','万元'],['实际执行','290','万元'],['预算结余','10','万元'],['执行率','96.7%','实际 / 预算']];
  slides.push({name:'typography-kpis',title:'文字层级与关键指标',elements:[...base('文字层级与关键指标',2),text('黑体标题 30 pt',48,125,390,48,30,{font:'SimHei',color:C.navy}),text('楷体正文 18 pt：数字与中文保持清晰层级。',48,190,850,42),text('标签 16 pt  /  表格图表 14 pt  /  页脚 12 pt',48,244,850,36,16,{color:C.muted}),...kpi.flatMap(([label,value,unit],i)=>[rect(48+i*222,313,204,135,i===3?C.navy:C.pale),text(label,64+i*222,329,172,27,16,{color:i===3?C.white:C.muted}),text(value,64+i*222,364,172,43,30,{font:'SimHei',color:i===3?C.white:C.navy}),text(unit,64+i*222,414,172,25,14,{color:i===3?'#D8E5F2':C.muted})])]});
  slides.push({name:'primitives',title:'基础图形 · 填充与轮廓',elements:[...base('基础图形 · 填充与轮廓',3),text('实心色块、真圆、菱形与轮廓共同构成视觉语言。',48,120,864,42),rect(64,200,164,140,C.blue),rect(292,200,140,140,C.teal,{shapeType:'Ellipse'}),rect(496,200,140,140,C.gold,{shapeType:'Diamond'}),rect(704,200,164,140,null,{lineColor:C.navy,lineWidth:2}),...['实心矩形','真圆 140 × 140','决策菱形','透明填充 / 轮廓'].map((t,i)=>text(t,48+i*222,363,204,48,16,{align:'Center'})),line(64,437,804,0,C.teal,{lineWidth:3}),text('填充强调区域；细线建立秩序；所有图形均可编辑。',48,450,864,30,14,{color:C.muted})]});
  const values=[['类别','预算（万元）','实际（万元）','结余（万元）','执行率'],...DEMO.categories.map((cat,i)=>[cat,String(DEMO.budget[i]),String(DEMO.actual[i]),String(DEMO.budget[i]-DEMO.actual[i]),(100*DEMO.actual[i]/DEMO.budget[i]).toFixed(1)+'%']),['合计','300','290','10','96.7%']];
  slides.push({name:'native-table',title:'原生表格 · 预算执行明细',elements:[...base('原生表格 · 预算执行明细',4),text('同一组示例数据，保留原生单元格编辑能力。',48,121,864,40),text('运营超预算 4 万元；其余类别合计结余 14 万元。',48,425,864,42,18,{color:C.navy})],table:{values,left:48,top:185,width:864,height:216,columns:[{columnWidth:144},{columnWidth:180},{columnWidth:180},{columnWidth:180},{columnWidth:180}],rows:Array.from({length:6},()=>({rowHeight:36})),uniformCellProperties:{font:{name:'KaiTi',size:14,color:C.ink},margins:{left:10,right:10,top:4,bottom:4},horizontalAlignment:'Center',verticalAlignment:'Middle'},specificCellProperties:values.map((row,r)=>row.map(()=>({fill:{color:r===0?C.navy:r===5?'#DEEAF5':r%2?C.white:C.pale},font:{name:r===0?'SimHei':'KaiTi',size:14,bold:r===0||r===5,color:r===0?C.white:C.ink}})))}});
  slides.push({name:'editable-data-shapes',title:'数据图形 · 预算与实际对照',elements:[...base('数据图形 · 预算与实际对照',5),rect(48,125,14,14,C.blue),text('预算',70,120,90,28,14),rect(175,125,14,14,C.teal),text('实际',197,120,90,28,14),text('单位：万元',752,120,160,28,14,{align:'Right',color:C.muted}),...[0,40,80,120].flatMap(v=>[line(174+v*4.5,167,0,256),text(String(v),156+v*4.5,429,40,24,14,{align:'Center',color:C.muted})]),...DEMO.categories.flatMap((cat,i)=>[text(cat,48,185+i*62,100,32,16),rect(174,177+i*62,DEMO.budget[i]*4.5,17,C.blue),rect(174,199+i*62,DEMO.actual[i]*4.5,17,C.teal),text(`${DEMO.budget[i]} / ${DEMO.actual[i]}`,745,181+i*62,166,35,14,{color:i===1?C.red:C.ink})]),text('原生矩形与文字组成数据图；数值、颜色和长度均可直接编辑。',48,458,864,28,14,{color:C.muted})]});
  slides.push({name:'embedded-image',title:'原始图片 · 嵌入与图文编排',elements:[...base('原始图片 · 嵌入与图文编排',6),rect(48,136,568,313,C.pale),text('原始 PNG',654,145,258,40,16,{color:C.teal}),text('嵌入素材保持原图\n清晰边界与留白\n文字说明独立可编辑',654,201,258,128,18),text('用于视觉参考与流程概览。\n详细步骤见下一页。',654,359,258,73,16,{color:C.muted})],image:{left:64,top:152,width:536,height:281}});
  slides.push({name:'native-flow',title:'四节点流程 · 原生连接线',elements:[...base('四节点流程 · 原生连接线',7),text('每一步都有明确输出，连接线以 Office 原生线条保留。',48,122,864,42),...['需求归集','预算评审','执行跟踪','复盘改进'].flatMap((label,i)=>[rect(48+i*228,217,180,92,i===3?C.teal:C.navy,{shapeType:'RoundRectangle',text:label,size:18,color:C.white,align:'Center',valign:'Middle'}),text(['收集业务目标','确认资源边界','记录实际支出','形成改进清单'][i],48+i*228,345,180,56,16,{align:'Center',color:C.muted})]),...[0,1,2].map(i=>line(228+i*228,263,48,0,C.blue,{lineWidth:3})),text('输入：四类预算 300 万元     输出：实际 290 万元，执行率 96.7%',48,434,864,42,16,{color:C.navy})]});
  slides.push({name:'timeline-three-columns',title:'时间线 · 三栏行动计划',elements:[...base('时间线 · 三栏行动计划',8),line(120,183,720,0,C.blue,{lineWidth:3}),...['月初 · 规划','月中 · 跟踪','月末 · 复盘'].flatMap((label,i)=>[rect(144+i*288,171,24,24,C.teal,{shapeType:'Ellipse'}),text(label,48+i*296,224,272,38,18,{font:'SimHei',color:C.navy}),rect(48+i*296,278,272,171,C.pale),text(['按四类配置预算\n建立责任人与基线\n锁定总额 300 万元','记录实际支出\n运营超支及时提醒\n滚动更新执行进度','核对实际 290 万元\n确认结余 10 万元\n归档执行率 96.7%'][i],66+i*296,296,236,133,18)])]});
  const roles=[['主色',C.navy,'标题与结构'],['信息',C.blue,'预算与基线'],['执行',C.teal,'实际与完成'],['提醒',C.gold,'关注与重点'],['风险',C.red,'超支与异常'],['底色',C.pale,'容器与分区']];
  slides.push({name:'color-roles-grid',title:'色彩角色 · 对齐网格',elements:[...base('色彩角色 · 对齐网格',9),text('语义化色彩与统一留白，使页面可以复用和扩展。',48,120,864,42),...roles.flatMap(([label,color,meaning],i)=>{const x=48+(i%3)*296,y=184+Math.floor(i/3)*138;return[rect(x,y,272,114,C.pale),rect(x,y,70,114,color),text(label,x+88,y+16,168,33,16,{font:'SimHei',color:C.navy}),text(meaning,x+88,y+57,168,30,16,{color:C.muted})]}),text('设计网格 960 × 540 pt  ·  外边距 48 pt  ·  三栏内容宽 272 pt  ·  栏间距 24 pt',48,465,864,24,12,{color:C.muted})]});
  slides.push({name:'closing',title:'从样式一致，到表达清楚',elements:[...base('从样式一致，到表达清楚',10,true),text('预算 300 万元\n实际 290 万元',48,150,520,114,30,{font:'SimHei',color:C.white}),text('结余 10 万元  ·  执行率 96.7%',48,293,520,52,18,{color:'#D8E5F2'}),rect(640,153,272,246,'#244C78'),text('可编辑交付',664,176,224,45,18,{font:'SimHei',color:C.white}),text('文字与基础图形\n原生表格与连接线\n嵌入原始 PNG\nOffice 原生导出与截图',664,240,224,140,16,{color:'#D8E5F2'}),text('研发 / 运营 / 营销 / 支持',48,421,570,40,16,{color:'#D8E5F2'})]});
  return {version:1,title:'预算执行 · PowerPoint 原生样式图谱',style:STYLE,demo:DEMO,slides};
}

// Minimal bounded ZIP reader for exported native evidence. Never creates an Office ZIP.
export function readNativeZip(buffer) {
  let eocd=-1; for(let i=buffer.length-22;i>=Math.max(0,buffer.length-65557);i--) if(buffer.readUInt32LE(i)===0x06054b50){eocd=i;break;}
  if(eocd<0) throw new Error('Native export is not a ZIP package');
  const result=new Map(); let cursor=buffer.readUInt32LE(eocd+16); const count=buffer.readUInt16LE(eocd+10);
  for(let i=0;i<count;i++) {if(buffer.readUInt32LE(cursor)!==0x02014b50) throw new Error('Invalid native ZIP directory'); const method=buffer.readUInt16LE(cursor+10),size=buffer.readUInt32LE(cursor+20),uncompressed=buffer.readUInt32LE(cursor+24),nl=buffer.readUInt16LE(cursor+28),el=buffer.readUInt16LE(cursor+30),cl=buffer.readUInt16LE(cursor+32),offset=buffer.readUInt32LE(cursor+42),name=buffer.subarray(cursor+46,cursor+46+nl).toString('utf8'); if(uncompressed>64*1024*1024)throw new Error('Native part too large'); const start=offset+30+buffer.readUInt16LE(offset+26)+buffer.readUInt16LE(offset+28),data=buffer.subarray(start,start+size); if(method!==0&&method!==8)throw new Error('Unsupported native ZIP method'); result.set(name,method===8?inflateRawSync(data,{maxOutputLength:64*1024*1024}):data);cursor+=46+nl+el+cl;}
  return result;
}
const unescapeXml=s=>s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
function orderedNativeParts(zip) {
  const xml=zip.get('ppt/presentation.xml')?.toString('utf8')||'',rels=zip.get('ppt/_rels/presentation.xml.rels')?.toString('utf8')||'';
  const attributes=tag=>Object.fromEntries([...tag.matchAll(/([\w:]+)="([^"]*)"/g)].map(m=>[m[1],unescapeXml(m[2])]));
  const relationships=new Map([...rels.matchAll(/<Relationship\b[^>]*\/?>/g)].map(m=>{const a=attributes(m[0]);return[a.Id,a];}));
  return [...xml.matchAll(/<p:sldId\b[^>]*\/?>/g)].map(m=>{const rel=relationships.get(attributes(m[0])['r:id']);if(!rel||rel.TargetMode==='External')throw new Error('Missing native slide relationship');const name=path.posix.normalize(rel.Target.startsWith('/')?rel.Target.slice(1):path.posix.join('ppt',rel.Target));if(!name.startsWith('ppt/')||!zip.has(name))throw new Error('Missing native slide part');return[name,zip.get(name)];});
}
export function inspectNativeExport(buffer, plan=buildAtlasPlan()) {
  const zip=readNativeZip(buffer),parts=orderedNativeParts(zip),fontChecks=[];
  const defaultFontSize=Number(zip.get('ppt/presentation.xml')?.toString('utf8').match(/<a:lvl1pPr\b[\s\S]*?<a:defRPr\b[^>]*sz="(\d+)"/)?.[1]||0)/100;
  const nativeSize=zip.get('ppt/presentation.xml')?.toString('utf8').match(/<p:sldSz\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/),sx=Number(nativeSize?.[1]||0)/12700/960,sy=Number(nativeSize?.[2]||0)/12700/540;
  const fontMatches=(name,expected)=>({SimHei:['SimHei','黑体'],KaiTi:['KaiTi','楷体']}[expected]||[expected]).includes(name);
  function checkRun(run,part,expected,expectedSize,table=false) {
    const text=[...run.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m=>unescapeXml(m[1])).join('');
    if(!text)return; const latin=run.match(/<a:latin\b[^>]*typeface="([^"]+)"/)?.[1]||null,eastAsia=run.match(/<a:ea\b[^>]*typeface="([^"]+)"/)?.[1]||null;
    const size=Number(run.match(/\bsz="(\d+)"/)?.[1]||defaultFontSize*100)/100,hasChinese=/[\u3400-\u9FFF]/.test(text);
    fontChecks.push({part,text,expected,expectedSize,size,table,latin,eastAsia,hasChinese,verified:fontMatches(latin,expected)&&(!hasChinese||fontMatches(eastAsia,expected))&&size===expectedSize});
  }
  const slides=parts.map(([name,b],index)=>{
    const xml=b.toString('utf8');
    for(const shape of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
      const block=shape[0],value=[...block.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m=>unescapeXml(m[1])).join('');
      if(!value)continue; const element=plan.slides[index]?.elements.find(e=>e.text?.replace(/[\r\n]/g,'')===value.replace(/[\r\n]/g,''));
      if(!element) { fontChecks.push({part:name,text:value,verified:false,reason:'Unmapped native text'});continue; }
      for(const run of block.matchAll(/<a:r\b[\s\S]*?<\/a:r>/g))checkRun(run[0],name,element.font,element.size);
    }
    const tableDetails=[...xml.matchAll(/<a:tbl>[\s\S]*?<\/a:tbl>/g)].map(table=>{
      const rows=[...table[0].matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)];
      rows.forEach((row,r)=>{for(const run of row[0].matchAll(/<a:r\b[\s\S]*?<\/a:r>/g))checkRun(run[0],name,r===0?'SimHei':'KaiTi',14,true);});
      return {rows:rows.length,cols:(table[0].match(/<a:gridCol\b/g)||[]).length,values:rows.map(row=>[...row[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map(cell=>[...cell[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m=>unescapeXml(m[1])).join('')))};
    });
    const circles=[...xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)].filter(m=>/<a:prstGeom\b[^>]*prst="ellipse"/.test(m[0])).map(m=>{const ext=m[0].match(/<a:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);return ext?{width:Number(ext[1])/12700,height:Number(ext[2])/12700,trueCircle:ext[1]===ext[2]}:{trueCircle:false};});
    const lineGeometries=[...xml.matchAll(/<p:cxnSp\b[\s\S]*?<\/p:cxnSp>/g)].map(m=>{const ext=m[0].match(/<a:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);return ext?{width:Number(ext[1])/12700,height:Number(ext[2])/12700}:null;});
    const expectedLines=plan.slides[index].elements.filter(e=>e.type==='line'),lineGeometryVerified=lineGeometries.length===expectedLines.length&&lineGeometries.every((g,i)=>g&&Math.abs(g.width-expectedLines[i].w*sx)<.01&&Math.abs(g.height-expectedLines[i].h*sy)<.01);
    const pictureGeometries=[...xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)].map(m=>{const ext=m[0].match(/<a:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);return ext?{width:Number(ext[1])/12700,height:Number(ext[2])/12700}:null;});
    return {lineGeometries,lineGeometryVerified,pictureGeometries,part:name,text:[...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m=>unescapeXml(m[1])).join(''),nativeTables:tableDetails.length,tableDetails,nativeConnectors:(xml.match(/<p:cxnSp\b/g)||[]).length,nativePictures:(xml.match(/<p:pic\b/g)||[]).length,noShrink:!/<a:normAutofit\b/.test(xml),circles};
  });
  const presentation=zip.get('ppt/presentation.xml')?.toString('utf8')||'',size=presentation.match(/<p:sldSz\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
  const fontsVerified=fontChecks.length>0&&fontChecks.every(c=>c.verified)&&fontChecks.some(c=>c.hasChinese&&c.expected==='SimHei')&&fontChecks.some(c=>c.hasChinese&&c.expected==='KaiTi');
  const tableVerified=JSON.stringify(slides[3]?.tableDetails[0]?.values)===JSON.stringify(plan.slides[3].table.values);
  return {source:'plugin ppt_export_file native PPTX readback',slideCount:slides.length,slideSize:size?{width:Number(size[1])/12700,height:Number(size[2])/12700}:null,fontsVerified,fontChecks,slides,tableVerified,originalPngEmbedded:false,passed:slides.length===10&&fontsVerified&&tableVerified&&slides.every(s=>s.noShrink&&s.lineGeometryVerified)&&slides[2]?.circles.some(c=>c.trueCircle)&&slides[3]?.nativeTables===1&&slides[5]?.nativePictures>=1&&slides[6]?.nativeConnectors>=3};
}

function options(argv) { const out={}; for(let i=0;i<argv.length;i++){if(!argv[i].startsWith('--'))throw new Error('Unexpected argument '+argv[i]); const key=argv[i].slice(2);out[key]=argv[i+1]&&!argv[i+1].startsWith('--')?argv[++i]:true;}return out; }
const hash=x=>createHash('sha256').update(x).digest('hex');
async function locateImage(root, explicit) {if(explicit)return path.resolve(explicit); const preferred=path.join(root,'docs/style-atlas/atlas-workflow.png');try{await fs.access(preferred);return preferred;}catch{/* Search the original image filename below. */} const entries=await fs.readdir(path.join(root,'docs/style-atlas'),{recursive:true});const match=entries.find(p=>path.basename(p)==='atlas-workflow.png')||entries.find(p=>path.basename(p)==='original-workflow.png');if(!match)throw new Error('Missing original workflow PNG: use --image <path>');return path.join(root,'docs/style-atlas',match);}
async function main() {
  const opt=options(process.argv.slice(2)),root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..'),plan=buildAtlasPlan();
  if(!opt.execute&&!opt.inspect){console.log(JSON.stringify(plan,null,2));return;}
  if(typeof opt.instance!=='string')throw new Error('Explicit --instance is required');
  const baseUrl=String(opt.url||`http://127.0.0.1:${opt.port||3000}`),outDir=path.resolve(root,String(opt.out||'docs/style-atlas/plugin-ppt')),runId=String(opt['run-id']||'ppt-style-atlas-v1');
  if(!/^[\w.-]{1,64}$/.test(runId))throw new Error('run-id must be 1-64 stable filename characters');
  await fs.mkdir(outDir,{recursive:true});const journalPath=path.join(outDir,`${runId}.journal.json`);let journal={runId,instance:opt.instance,baseUrl,commands:{}};try{journal=JSON.parse(await fs.readFile(journalPath,'utf8'));}catch(e){if(e.code!=='ENOENT')throw e;}
  if(journal.instance!==opt.instance||journal.baseUrl!==baseUrl)throw new Error('Journal target differs; choose a fresh run-id');
  const persist=()=>fs.writeFile(journalPath,JSON.stringify(journal,null,2),'utf8');
  const get=async route=>{const r=await fetch(baseUrl+route,{signal:AbortSignal.timeout(20000)});return r.json();};
  async function command(step,action,args={},mutating=false) {
    const key=`${runId}:${step}`,fingerprint=hash(JSON.stringify({action,args,instance:opt.instance})),old=journal.commands[step];
    if(old&&old.fingerprint!==fingerprint)throw new Error(`Changed arguments for ${step}; use another run-id`);
    if(old?.status==='completed')return old.result;
    if(old) {if(old.commandId){const view=await get('/office/commands/'+old.commandId);if(view.status==='completed'&&view.response?.ok){old.status='completed';old.result=view.response.result;await persist();return old.result;}old.lastQuery=view;await persist();}throw new Error(`Unknown/incomplete outcome at ${step}; journal retained. Do not resend; query its commandId.`);}
    if(mutating&&(!opt.execute||!opt['allow-write']))throw new Error('Content writes require --execute --allow-write after target preservation');
    journal.commands[step]={action,fingerprint,idempotencyKey:key,status:'sending',sentAt:new Date().toISOString()};await persist();
    let response;try {const r=await fetch(baseUrl+'/office/command',{method:'POST',headers:{'content-type':'application/json; charset=utf-8'},body:Buffer.from(JSON.stringify({host:'PowerPoint',instance:opt.instance,action,args,idempotencyKey:key}),'utf8'),signal:AbortSignal.timeout(110000)});response=await r.json();}catch(e){journal.commands[step].status='transport_unknown';journal.commands[step].error=String(e);const status=await get('/office/status').catch(()=>null);journal.commands[step].pendingCandidates=status?.pendingCommands?.filter(c=>c.instance===opt.instance&&c.action===action)||[];await persist();throw new Error(`Transport outcome unknown for ${step}; no retry. Inspect ${journalPath}`);}
    const record=journal.commands[step];record.commandId=response.commandId;record.response=response;record.status=response.ok?'completed':response.code==='execution_unknown'?'execution_unknown':'failed';if(response.ok)record.result=response.result;await persist();
    if(!response.ok){if(response.commandId)record.lastQuery=await get('/office/commands/'+response.commandId).catch(()=>null);await persist();throw new Error(`${action}: ${response.code}: ${response.error||'failed'}; no retry`);}return response.result;
  }
  const status=await get('/office/status');const instances=Array.isArray(status.instances)?status.instances:Object.entries(status.instances||{}).map(([instanceId,v])=>({instanceId,...v}));const target=instances.find(i=>(i.instanceId||i.instance)===opt.instance);if(!target||target.host!=='PowerPoint')throw new Error('Target PowerPoint instance is not online');
  const info=await command('preflight-size','ppt_read_presentation_info'),before=await command('preflight-slides','ppt_inspect_slides',{withStyles:true}),notes=await command('preflight-notes','ppt_read_notes'),document=await command('preflight-document','read_document');
  journal.preflight={target,info,before,notes,document};await persist();
  if(opt.inspect&&!opt.execute){console.log(JSON.stringify(journal.preflight,null,2));return;}
  if(!opt['allow-write'])throw new Error('No write authorization supplied; use --allow-write only after coordinator preservation');
  const size=info.slideSize;if(!size||!Number.isFinite(size.width)||!Number.isFinite(size.height))throw new Error('Cannot establish native slide dimensions');
  if(Math.abs(size.width/size.height-16/9)>.001)throw new Error('Native presentation must use 16:9 to retain true circles under uniform scaling');
  if((Math.abs(size.width-960)>.01||Math.abs(size.height-540)>.01)&&!opt['allow-default-size'])throw new Error(`Native slide size ${size.width}x${size.height} pt; review it and use --allow-default-size to scale geometry`);
  if(document.text.replace(/--- 幻灯片 \d+ ---/g,'').trim()||before.slides.some(s=>s.shapes.some(sh=>sh.text?.trim()||!['TextBox','GeometricShape','Placeholder'].includes(sh.type))))throw new Error('Target contains user content; preserve it and open an empty presentation');
  const imagePath=await locateImage(root,opt.image),imageBytes=await fs.readFile(imagePath);if(imageBytes.subarray(0,8).toString('hex')!=='89504e470d0a1a0a')throw new Error('Workflow asset must be original PNG');
  const added=await command('add-ten-slides','ppt_add_slides',{count:10,clearPlaceholders:true},true),slideIds=added.added.map(s=>String(s.id));if(slideIds.length!==10)throw new Error('Expected ten created native slides');journal.slideIds=slideIds;await persist();
  if(!opt['repair-native']) for(let i=0;i<10;i++) {const slide=plan.slides[i],slideId=slideIds[i];await command(`slide-${i+1}-elements`,'ppt_add_elements',{slideId,coordinateSpace:{width:960,height:540},targetSize:{width:size.width,height:size.height},elements:slide.elements},true);if(slide.table){const t=structuredClone(slide.table),sx=size.width/960,sy=size.height/540;for(const key of ['left','width'])t[key]*=sx;for(const key of ['top','height'])t[key]*=sy;t.columns.forEach(c=>c.columnWidth*=sx);t.rows.forEach(r=>r.rowHeight*=sy);await command('slide-4-table','ppt_insert_table',{slideId,...t},true);}if(slide.image){const box=slide.image,ratio=Math.min(box.width/imageBytes.readUInt32BE(16),box.height/imageBytes.readUInt32BE(20)),g={width:imageBytes.readUInt32BE(16)*ratio,height:imageBytes.readUInt32BE(20)*ratio};g.left=box.left+(box.width-g.width)/2;g.top=box.top+(box.height-g.height)/2;await command('slide-6-image','ppt_insert_image',{slideId,base64:imageBytes.toString('base64'),mode:'positioned',left:g.left*size.width/960,top:g.top*size.height/540,width:g.width*size.width/960,height:g.height*size.height/540},true);}}
  if(!opt['repair-native']&&before.slides.length)await command('remove-original-empty-slides','ppt_delete_slides',{slideIds:before.slides.map(s=>String(s.id))},true);
  const suffix=opt['repair-native']?'-repair-v1':'';
  if(opt['repair-native']) {
    if(journal.commands['slide-10-elements']?.status!=='completed'||journal.commands['remove-original-empty-slides']?.status!=='completed')throw new Error('Repair requires a fully generated atlas journal');
    const live=await command('repair-native-preflight','ppt_inspect_slides',{withStyles:true});
    if(live.slides.length!==10||live.slides.some((s,i)=>String(s.id)!==slideIds[i]))throw new Error('Live slides differ from owned atlas');
    for(let i=0;i<10;i++) {
      const result=journal.commands['slide-'+(i+1)+'-elements'].result,updates=[];
      plan.slides[i].elements.forEach((e,index)=>{if(e.type==='line'){const shapeId=String(result.elements[index].id);if(!live.slides[i].shapes.some(s=>String(s.id)===shapeId&&s.type==='Line'))throw new Error('Owned native line missing');updates.push({shapeId,width:e.w*size.width/960,height:e.h*size.height/540});}});
      if(i===5){const image=live.slides[i].shapes.filter(s=>s.type==='Image');if(image.length!==1)throw new Error('Expected one owned image');const b=plan.slides[i].image,r=Math.min(b.width/imageBytes.readUInt32BE(16),b.height/imageBytes.readUInt32BE(20)),width=imageBytes.readUInt32BE(16)*r,height=imageBytes.readUInt32BE(20)*r;updates.push({shapeId:String(image[0].id),left:(b.left+(b.width-width)/2)*size.width/960,top:(b.top+(b.height-height)/2)*size.height/540,width:width*size.width/960,height:height*size.height/540});}
      if(updates.length)await command('repair-slide-'+(i+1)+'-geometry-v1','ppt_update_shapes',{slideId:slideIds[i],updates},true);
    }
  }
  const inspection=await command('inspect-complete'+suffix,'ppt_inspect_slides',{slideIds,withStyles:true});await fs.writeFile(path.join(outDir,'ppt-base-styles.inspect.json'),JSON.stringify(inspection,null,2),'utf8');
  for(let i=0;i<10;i++){const png=await command(`render-${i+1}${suffix}`,'ppt_render_slides',{slideIds:[slideIds[i]],width:960,height:540});if(png.slides.length!==1)throw new Error('Unexpected native render count');await fs.writeFile(path.join(outDir,`ppt-base-styles-slide-${String(i+1).padStart(2,'0')}.png`),Buffer.from(png.slides[0].base64,'base64'));}
  const exported=await command('export-native'+suffix,'ppt_export_file'),native=Buffer.from(exported.base64,'base64');await fs.writeFile(path.join(outDir,'ppt-base-styles.pptx'),native);
  const pdf=await command('export-native-pdf'+suffix,'ppt_export_file',{type:'pdf'}); const pdfBytes=Buffer.from(pdf.base64,'base64'); if(pdfBytes.subarray(0,5).toString()!=='%PDF-')throw new Error('Native PDF export did not return PDF bytes'); await fs.writeFile(path.join(outDir,'ppt-base-styles.pdf'),pdfBytes);
  const evidence=inspectNativeExport(native,plan),zip=readNativeZip(native);evidence.originalPngEmbedded=[...zip].some(([name,b])=>name.startsWith('ppt/media/')&&b.equals(imageBytes));evidence.imageAspectPreserved=evidence.slides[5].pictureGeometries.every(g=>g&&Math.abs(g.width/g.height-imageBytes.readUInt32BE(16)/imageBytes.readUInt32BE(20))<.001);evidence.passed&&=evidence.originalPngEmbedded&&evidence.imageAspectPreserved;evidence.instance=opt.instance;evidence.imageSource=path.relative(root,imagePath);evidence.imageSha256=hash(imageBytes);evidence.nativeSizeVerified=info.slideSize;evidence.requestedSize={width:960,height:540};evidence.usedDefaultSize=!!opt['allow-default-size'];await fs.writeFile(path.join(outDir,'ppt-base-styles.evidence.json'),JSON.stringify(evidence,null,2),'utf8');console.log(JSON.stringify({outDir,nativeSlideCount:evidence.slideCount,passed:evidence.passed,fontsVerified:evidence.fontsVerified,originalPngEmbedded:evidence.originalPngEmbedded},null,2));if(!evidence.passed)throw new Error('Native export verification failed; inspect evidence before claiming completion');
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url))main().catch(e=>{console.error(e.message);process.exitCode=1;});
