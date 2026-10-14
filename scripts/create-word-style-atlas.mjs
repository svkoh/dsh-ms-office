#!/usr/bin/env node
/** Build a five-page Word atlas. Only plugin HTTP commands write Office content.
 * No args: plan only. --inspect: read-only. --execute --allow-write: explicit target.
 * node scripts/create-word-style-atlas.mjs --inspect --instance inst-... --port 3000
 * node scripts/create-word-style-atlas.mjs --execute --allow-write --instance inst-... --port 3000 --run-id word-v1
 * Native DOCX bytes are transported unchanged. ZIP inspection is read-only. No PDF export.
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {readNativeZip} from './create-ppt-style-atlas.mjs';

export const DEMO=Object.freeze({categories:['研发','运营','营销','支持'],budget:[120,80,60,40],actual:[114,84,54,38],totalBudget:300,totalActual:290,balance:10,execution:'96.7%',unit:'万元'});
const MARK={toc:'DSH_WORD_ATLAS_TOC',table:'DSH_WORD_ATLAS_TABLE',image:'DSH_WORD_ATLAS_IMAGE'};
const xmlText=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
const para=(text,role='Normal',extra={})=>({text,role,...extra});
export function buildAtlasPlan(){
  const pages=[
    {name:'cover',title:'Office 文档基础样式范本',paragraphs:[para('Office 文档基础样式范本','Title',{alignment:'center',spaceBefore:120,spaceAfter:24}),para('一份文件 · 常用元素 · 可复用样式','Normal',{alignment:'center',firstLineIndent:0,spaceAfter:12}),para('封面 / 目录 / 章节 / 段落 / 表格 / 图片 / 页眉页脚','Normal',{alignment:'center',firstLineIndent:0,spaceAfter:12}),para('适用：一般企业办公文档','Normal',{alignment:'center',firstLineIndent:0,spaceAfter:12}),para('DEMO / 示例数据 · 2026年10月','Normal',{alignment:'center',firstLineIndent:0})]},
    {name:'contents',title:'目录',paragraphs:[para('目录','Title',{alignment:'left',firstLineIndent:0}),para(MARK.toc,'Normal',{firstLineIndent:0}),para('目录由 Word 原生 TOC 域读取实际标题层级；修改标题后更新域即可同步。','Normal')]},
    {name:'paragraphs',title:'一、章节、段落与列表',paragraphs:[para('一、章节、段落与列表','Heading1'),para('（一）标题层级与正文','Heading2'),para('1. 三级标题示例','Heading3'),para('正文使用宋体12pt、1.5倍行距、首行缩进24pt和两端对齐。标题使用全局语义样式，自动目录读取实际标题层级。'),para('演示数据统一为预算300万元、实际支出290万元、余额10万元、执行率96.7%。这些数值仅供样式示例，正式使用时应替换并核验。'),para('（二）按用途选择对齐方式','Heading2'),para('左对齐：短说明、标签和列表。','Normal',{alignment:'left',firstLineIndent:0}),para('居中：独立标题与图注。','Normal',{alignment:'center',firstLineIndent:0}),para('右对齐：日期与签署位置。','Normal',{alignment:'right',firstLineIndent:0}),para('两端对齐：连续正文保持左右边界整齐，末行自然结束。'),para('（三）项目符号和编号','Heading2'),para('并列事项使用项目符号，保持同级缩进。','ListBullet',{list:'ul'}),para('状态同时使用文字说明，颜色仅作辅助。','ListBullet',{list:'ul'}),para('先确认内容与数据口径。','ListNumber',{list:'ol'}),para('再应用元素样式并检查效果。','ListNumber',{list:'ol'})]},
    {name:'table',title:'二、表格、对齐与颜色',paragraphs:[para('二、表格、对齐与颜色','Heading1'),para('（一）原生表格与汇总','Heading2'),para('金额右对齐，文字左对齐；表头可重复，数据单元格保留可编辑结构。下表单位为万元，全部数值为 DEMO。'),para(MARK.table),para('表1　预算与实际对比（DEMO）','Caption',{alignment:'center',firstLineIndent:0}),para('（二）强调、颜色与签署','Heading2'),para('默认正文使用黑色。重要判断使用加粗；辅助标签少量使用哑铜色，不能只靠颜色区分状态。','Normal',{font:{bold:true}}),para('辅助标签：统计口径 / 示例数据 / 待复核','Normal',{alignment:'left',firstLineIndent:0,font:{color:'#836344'}}),para('复核人：文档编制者与复核者','Normal',{alignment:'right',firstLineIndent:0})]},
    {name:'image',title:'三、图片与页眉页脚',paragraphs:[para('三、图片与页眉页脚','Heading1'),para('（一）原始图片与图注','Heading2'),para('下图为原创 PNG，以原始宽高比通过 Word 原生图片接口嵌入。图注独立可编辑。'),para(MARK.image,'Normal',{alignment:'center',firstLineIndent:0}),para('图1　从内容规划到效果复核（原创示意图）','Caption',{alignment:'center',firstLineIndent:0}),para('（二）页眉、页脚与真实域','Heading2'),para('页脚自动显示页码，目录按实际标题层级生成；修改标题或分页后应更新目录与页码。封面采用首页不同设置，不显示页眉和页码。'),para('（三）按元素复用','Heading2'),para('先另存工作副本，再替换演示内容。保留标题层级、表格结构与图片比例；完成后逐页检查标题孤行、表格跨页、图片边界和页码。')]}
  ];
  const rows=[['类别','预算/万元','实际/万元','余额/万元','说明'],...DEMO.categories.map((v,i)=>[v,String(DEMO.budget[i]),String(DEMO.actual[i]),String(DEMO.budget[i]-DEMO.actual[i]),i===1?'待复核':'按计划']),['合计','300','290','10','96.7%']];
  return {version:1,title:'Office 文档基础样式范本',logicalPages:5,pages,rows,image:'docs/style-atlas/assets/original-workflow.png',demo:DEMO};
}
export function buildHtml(plan=buildAtlasPlan()){
  let html='',list=null;for(const p of plan.pages.flatMap(p=>p.paragraphs)){if(list!==p.list){if(list)html+='</'+list+'>';list=p.list||null;if(list)html+='<'+list+'>';}html+=list?'<li>'+xmlText(p.text)+'</li>':'<p>'+xmlText(p.text)+'</p>';}if(list)html+='</'+list+'>';return html;
}
export function isOwnedTemplate(text){const s=String(text||'').trim();return !s || (s.includes('Office 文档基础样式范本')&&s.includes('DEMO')&&s.includes('300')&&s.includes('290')&&s.includes('目录')&&(s.includes('三、图片与页眉页脚')||s.includes('三、图片、图注与页眉页脚')));}
export function resolveStyleNames(styles){
  const wanted={Normal:['Normal','正文'],Title:['Title','标题'],Heading1:['Heading 1','Heading1','标题 1','标题1'],Heading2:['Heading 2','Heading2','标题 2','标题2'],Heading3:['Heading 3','Heading3','标题 3','标题3'],Caption:['Caption','题注']},found={};
  for(const [role,names] of Object.entries(wanted)){const item=styles.find(s=>names.includes(s.name));if(!item)throw new Error('Missing global style '+role+'; inspect localized names');found[role]=item.name;}return found;
}

export function buildParagraphTargets(readback,specs){
  const used=new Set();return specs.map(spec=>{const {text,...patch}=spec;const p=readback.paragraphs?.find(p=>!used.has(p.index)&&String(p.text).trim()===text);if(!p)throw new Error('Exact paragraph not found: '+text);used.add(p.index);return {index:p.index,expectedText:p.text,...patch};});
}
const hash=value=>createHash('sha256').update(value).digest('hex');
const attr=(xml,key)=>xml.match(new RegExp('(?:^|[\\s])'+key+'="([^"]*)"'))?.[1]||null;
export function countNativeListParagraphs(document,stylesXml,numberingXml){
  const ids=new Set([...numberingXml.matchAll(/<w:num\b[^>]*>/g)].map(m=>attr(m[0],'w:numId'))),styles=new Map([...stylesXml.matchAll(/<w:style\b[\s\S]*?<\/w:style>/g)].map(m=>[attr(m[0].split('>')[0],'w:styleId'),m[0]]));
  const numId=xml=>attr(xml.match(/<w:numId\b[^>]*>/)?.[0]||'','w:val');
  const inherited=(id,seen=new Set())=>{if(!id||seen.has(id)||!styles.has(id))return null;seen.add(id);const style=styles.get(id);return numId(style)??inherited(attr(style.match(/<w:basedOn\b[^>]*>/)?.[0]||'','w:val'),seen);};
  return [...document.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].filter(m=>{const props=m[0].match(/<w:pPr\b[\s\S]*?<\/w:pPr>/)?.[0]||'',id=numId(props)??inherited(attr(props.match(/<w:pStyle\b[^>]*>/)?.[0]||'','w:val'));return id&&id!=='0'&&ids.has(id);}).length;
}
export function inspectNativeExport(buffer,imageBytes){
  const zip=readNativeZip(buffer),document=zip.get('word/document.xml')?.toString('utf8')||'',stylesXml=zip.get('word/styles.xml')?.toString('utf8')||'',app=zip.get('docProps/app.xml')?.toString('utf8')||'';
  const findStyle=id=>[...stylesXml.matchAll(/<w:style\b[\s\S]*?<\/w:style>/g)].find(m=>[attr(m[0].split('>')[0],'w:styleId'),attr(m[0].match(/<w:name[^>]*>/)?.[0]||'','w:val')].some(name=>name?.replaceAll(' ','').toLowerCase()===id.toLowerCase()))?.[0]||'';
  const definitions={};
  for(const [id,font,size] of [['Normal','宋体',24],['Body','宋体',24],['Heading1','黑体',28],['Title','黑体',36]]){const block=findStyle(id==='Body'?'范本正文':id),fonts=block.match(/<w:rFonts\b[^>]*>/)?.[0]||'',spacing=block.match(/<w:spacing\b[^>]*>/)?.[0]||'',sz=block.match(/<w:sz\b[^>]*>/)?.[0]||'',indent=block.match(/<w:ind\b[^>]*>/)?.[0]||'';definitions[id]={styleId:attr(block.split('>')[0],'w:styleId'),name:attr(block.match(/<w:name[^>]*>/)?.[0]||'','w:val'),font:attr(fonts,'w:ascii'),eastAsia:attr(fonts,'w:eastAsia'),sizeHalfPoints:Number(attr(sz,'w:val')),line:attr(spacing,'w:line'),lineRule:attr(spacing,'w:lineRule'),firstLineTwips:attr(indent,'w:firstLine'),verified:[font,font==='宋体'?'SimSun':'SimHei'].includes(attr(fonts,'w:eastAsia'))&&Number(attr(sz,'w:val'))===size};}
  const nativeFields=[...document.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)].map(m=>m[1]);
  const footers=[...zip].filter(([n])=>/^word\/footer\d+\.xml$/.test(n)).map(([part,b])=>({part,xml:b.toString('utf8')}));
  const hasPage=footers.some(f=>/\bPAGE\b/i.test(f.xml)),hasToc=nativeFields.some(f=>/\bTOC\b/i.test(f));
  const pageSize=document.match(/<w:pgSz\b[^>]*>/)?.[0]||'',margin=document.match(/<w:pgMar\b[^>]*>/)?.[0]||'';
  const pageCount=Number(app.match(/<(?:\w+:)?Pages>(\d+)<\/(?:\w+:)?Pages>/)?.[1])||null;
  const nativeTables=(document.match(/<w:tbl\b/g)||[]).length,nativePictures=(document.match(/<pic:pic\b/g)||[]).length,originalPngEmbedded=[...zip].some(([n,b])=>n.startsWith('word/media/')&&b.equals(imageBytes));
  const nativeListParagraphs=countNativeListParagraphs(document,stylesXml,zip.get('word/numbering.xml')?.toString('utf8')||''),hasNativeNumbering=!!zip.get('word/numbering.xml');
  const titlePg=document.match(/<w:titlePg\b[^>]*>/)?.[0]||'',firstPageDifferent=!!titlePg&&!['0','false','off'].includes(attr(titlePg,'w:val')),firstPageReferences=[...document.matchAll(/<w:(header|footer)Reference\b[^>]*w:type="first"[^>]*>/g)].map(m=>m[0]);
  const decode=xml=>xml.replaceAll('&lt;','<').replaceAll('&gt;','>').replaceAll('&quot;','"').replaceAll('&apos;',"'").replaceAll('&amp;','&');
  const actualParagraphs=[...document.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map(m=>{const props=m[0].match(/<w:pPr\b[\s\S]*?<\/w:pPr>/)?.[0]||'',spacing=props.match(/<w:spacing\b[^>]*>/)?.[0]||'';return {text:[...m[0].matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map(t=>decode(t[1])).join(''),styleId:attr(props.match(/<w:pStyle\b[^>]*>/)?.[0]||'','w:val'),line:attr(spacing,'w:line'),lineRule:attr(spacing,'w:lineRule')};});
  const tocStyleIds=new Set([...stylesXml.matchAll(/<w:style\b[\s\S]*?<\/w:style>/g)].filter(m=>/^toc[1-3]$/i.test((attr(m[0].match(/<w:name[^>]*>/)?.[0]||'','w:val')||'').replaceAll(' ',''))).map(m=>attr(m[0].split('>')[0],'w:styleId')));
  const tocParagraphs=actualParagraphs.filter(p=>tocStyleIds.has(p.styleId));
  const tocCacheApplications=buildAtlasPlan().pages.flatMap((page,i)=>page.paragraphs.filter(p=>/^Heading[1-3]$/.test(p.role)).map(p=>({heading:p.text,page:i+1,found:tocParagraphs.some(actual=>actual.text===p.text+String(i+1))})));
  const tocCacheVerified=tocCacheApplications.length>0&&tocCacheApplications.every(p=>p.found);
  const bodyApplications=buildAtlasPlan().pages.flatMap(p=>p.paragraphs).filter(p=>p.role==='Normal'&&!Object.values(MARK).includes(p.text)).map(p=>{const found=actualParagraphs.find(a=>a.text===p.text),applied=!!found&&found.styleId===definitions.Body.styleId,spacing=!!found&&(!found.line||(found.line==='360'&&(!found.lineRule||found.lineRule==='auto')));return {text:p.text,found:!!found,styleId:found?.styleId,applied,spacing};});
  const bodyParagraphsVerified=bodyApplications.length>0&&bodyApplications.every(p=>p.applied&&p.spacing);
  const trueOnePointFive=definitions.Body.line==='360'&&(!definitions.Body.lineRule||definitions.Body.lineRule==='auto')&&bodyParagraphsVerified;
  const fontsVerified=[definitions.Body,definitions.Heading1,definitions.Title].every(v=>v.verified),firstLineVerified=definitions.Body.firstLineTwips==='480';
  const a4=Math.abs(Number(attr(pageSize,'w:w'))-11906)<=1&&Math.abs(Number(attr(pageSize,'w:h'))-16838)<=1;
  const margins25mm=['top','bottom','left','right'].every(k=>Math.abs(Number(attr(margin,'w:'+k))-1417)<=1);
  return {source:'Office.js plugin word_export_file native DOCX',logicalPages:5,nativePageCount:pageCount,nativeTables,nativePictures,nativeListParagraphs,hasNativeNumbering,hasToc,tocCacheVerified,tocCacheApplications,hasPage,originalPngEmbedded,globalStyles:definitions,bodyStyle:{name:definitions.Body.name,styleId:definitions.Body.styleId},bodyApplications,bodyParagraphsVerified,normalStyleUsedForBody:false,fontsVerified,trueOnePointFive,firstLineVerified,a4,margins25mm,firstPageDifferent,firstPageReferences,passed:pageCount===5&&nativeListParagraphs>=4&&hasNativeNumbering&&nativeTables===1&&nativePictures>=1&&hasToc&&tocCacheVerified&&hasPage&&originalPngEmbedded&&fontsVerified&&trueOnePointFive&&firstLineVerified&&a4&&margins25mm&&firstPageDifferent};
}
function parseOptions(argv){const out={};for(let i=0;i<argv.length;i++){if(!argv[i].startsWith('--'))throw new Error('Unexpected argument '+argv[i]);const key=argv[i].slice(2);out[key]=argv[i+1]&&!argv[i+1].startsWith('--')?argv[++i]:true;}return out;}
async function main(){
  const opt=parseOptions(process.argv.slice(2)),root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..'),plan=buildAtlasPlan();
  if(!opt.execute&&!opt.inspect){console.log(JSON.stringify(plan,null,2));return;}
  if(typeof opt.instance!=='string')throw new Error('Explicit --instance required');
  if(opt.execute&&!opt['allow-write'])throw new Error('Execution requires --allow-write');
  const baseUrl=String(opt.url||`http://127.0.0.1:${opt.port||3000}`),runId=String(opt['run-id']||'word-style-atlas-v1'),outDir=path.resolve(root,String(opt.out||'docs/style-atlas/plugin-word'));
  const url=new URL(baseUrl);if(!['127.0.0.1','localhost','[::1]'].includes(url.hostname)||url.protocol!=='http:')throw new Error('Bridge must be local HTTP');
  if(!/^[\w.-]{1,64}$/.test(runId))throw new Error('run-id must be stable 1..64 filename characters');
  await fs.mkdir(outDir,{recursive:true});const journalPath=path.join(outDir,runId+'.journal.json');let journal={runId,instance:opt.instance,baseUrl,commands:{}};
  try{journal=JSON.parse(await fs.readFile(journalPath,'utf8'));}catch(e){if(e.code!=='ENOENT')throw e;}
  if(journal.instance!==opt.instance||journal.baseUrl!==baseUrl)throw new Error('Journal target differs; use a fresh run-id');
  const persist=()=>fs.writeFile(journalPath,JSON.stringify(journal,null,2),'utf8');
  const get=async route=>{const r=await fetch(baseUrl+route,{signal:AbortSignal.timeout(20000)});if(!r.ok)throw new Error('Bridge HTTP '+r.status);return r.json();};
  async function command(step,action,args={},write=false){
    const fingerprint=hash(JSON.stringify({action,args,instance:opt.instance})),key=runId+':'+step+(write?'':':read-'+Date.now()),old=journal.commands[step];
    if(write&&old&&old.fingerprint!==fingerprint)throw new Error('Changed step arguments: '+step);
    if(write&&old?.status==='completed')return old.result;
    if(write&&old){if(old.commandId){const state=await get('/office/commands/'+old.commandId);old.lastQuery=state;if(state.status==='completed'&&state.response?.ok){old.status='completed';old.result=state.response.result;await persist();return old.result;}await persist();}throw new Error('Unresolved previous command '+step+'; query commandId, never blindly resend');}
    if(write&&(!opt.execute||!opt['allow-write']))throw new Error('Writes require --execute --allow-write');
    if(old&&!write)(journal.readHistory??=[]).push({step,...old});
    const record=journal.commands[step]={action,fingerprint,idempotencyKey:key,status:'sending',sentAt:new Date().toISOString()};await persist();if(write)console.log('[Word] '+step+' -> '+action);
    let response;try{const r=await fetch(baseUrl+'/office/command',{method:'POST',headers:{'content-type':'application/json; charset=utf-8'},body:Buffer.from(JSON.stringify({action,host:'Word',instance:opt.instance,args,idempotencyKey:key}),'utf8'),signal:AbortSignal.timeout(200000)});response=await r.json();}catch(e){record.status='transport_unknown';record.error=String(e);record.bridgeStatus=await get('/office/status').catch(()=>null);await persist();throw new Error('Unknown transport outcome at '+step+'; stopped without retry');}
    record.commandId=response.commandId;record.response=response;record.status=response.ok?'completed':response.code||'failed';if(response.ok)record.result=response.result;await persist();
    if(!response.ok){if(response.commandId)record.lastQuery=await get('/office/commands/'+response.commandId).catch(()=>null);await persist();throw new Error(action+': '+response.code+': '+response.error+'; stopped without retry');}return response.result;
  }
  const status=await get('/office/status'),instance=status.instances?.find(i=>i.instanceId===opt.instance);if(!instance?.online||instance.host!=='Word')throw new Error('Explicit target is not an online Word instance');
  const capabilities=await get('/office/capabilities?lang=zh'),required=['read_document','get_environment','word_read_styles','word_set_styles','word_insert_html','word_format_paragraphs','locate_select','word_insert_break','insert_table','set_table_borders','word_insert_image','word_insert_field','word_read_fields','word_set_header_footer','word_read_headers_footers','word_set_page_setup','read_tables','word_list_paragraphs','word_list_sections','write_selection','word_export_file'];
  const missing=required.filter(n=>!capabilities.actions?.some(a=>a.name===n));if(missing.length)throw new Error('Missing plugin capabilities: '+missing.join(', '));
  const before=await command('preflight-document-'+Date.now(),'read_document');
  if(!isOwnedTemplate(before.text))throw new Error('Target is neither blank nor the owned DEMO template; no writes sent');
  const environment=await command('preflight-environment-'+Date.now(),'get_environment'),sections=await command('preflight-sections-'+Date.now(),'word_list_sections'),available=await command('preflight-styles-'+Date.now(),'word_read_styles'),names=resolveStyleNames(available.styles||[]);
  if(!environment.probe?.supportedDesktopSets?.includes('1.3'))throw new Error('WordApiDesktop 1.3 page setup not verified');
  if(sections.count!==1&&sections.sections?.length!==1)throw new Error('Owned template must have exactly one section');
  if(opt.inspect){console.log(JSON.stringify({instance,before,environment,sections,names,logicalPages:5,writePlanned:false},null,2));return;}
  const imageBytes=await fs.readFile(path.join(root,plan.image));
  await command('body','word_insert_html',{target:'body',mode:'replace',html:buildHtml(plan)},true);
  await command('body-style-preset','word_set_styles',{styles:[{name:'范本正文',preset:'body-1.5'}]},true);
  names.Normal='范本正文';
  await command('styles','word_set_styles',{styles:Object.entries(names).map(([role,name])=>({name,font:{name:['Title','Heading1'].includes(role)?'黑体':'宋体',size:role==='Title'?18:role==='Heading1'?14:12,color:'#000000',bold:role!=='Normal'&&role!=='Caption'},paragraphFormat:{alignment:role==='Normal'?'justify':'left',firstLineIndent:role==='Normal'?24:0,spaceBefore:role.startsWith('Heading')?12:0,spaceAfter:role.startsWith('Heading')?6:0}}))},true);
  const targets=plan.pages.flatMap(p=>p.paragraphs).map(p=>{const {text,role,list,...format}=p;return {text,...(!list?{style:text===MARK.toc?'Normal':role==='Normal'?'范本正文':role}:{}),...(list?{font:{name:'宋体',size:12,color:'#000000'},firstLineIndent:0,leftIndent:18}:{}),...format};});
  const initialParagraphs=await command('read-initial-paragraphs','word_list_paragraphs');
  await command('paragraph-styles','word_format_paragraphs',{targets:buildParagraphTargets(initialParagraphs,targets)},true);
  await command('page-setup','word_set_page_setup',{sections:'all',pageWidth:210*72/25.4,pageHeight:297*72/25.4,differentFirstPageHeaderFooter:true,topMargin:25*72/25.4,bottomMargin:25*72/25.4,leftMargin:25*72/25.4,rightMargin:25*72/25.4,headerDistance:36,footerDistance:36},true);
  async function select(step,text){const found=await command(step+'-'+Date.now(),'locate_select',{text,blinks:0});if(!found.selected)throw new Error('Missing unique insertion anchor '+text);return found;}
  for(let i=1;i<plan.pages.length;i++){const current=await command('read-break-anchor-'+i,'word_list_paragraphs'),paragraph=current.paragraphs.find(p=>String(p.text).trim()===plan.pages[i].title);if(!paragraph)throw new Error('Missing exact page heading');await command('break-'+i,'word_insert_break',{type:'page',location:'beforeSelection',paragraph:{index:paragraph.index,expectedText:paragraph.text}},true);}
  await select('select-table',MARK.table);await command('table','insert_table',{rows:plan.rows,location:'beforeSelection',styleBuiltIn:'TableGrid',headerRow:true},true);
  await select('select-table-marker',MARK.table);await command('remove-table-marker','write_selection',{text:''},true);
  await command('table-borders','set_table_borders',{tableIndex:0,color:'#D6E0E9',type:'Single',width:0.5,enableGrid:true},true);
  const tableParas=await command('read-table-paragraphs','word_list_paragraphs');
  const tableTargets=[];
  for(const row of plan.rows)for(let c=0;c<row.length;c++){const text=row[c];tableTargets.push({text,font:{name:'宋体',size:12,color:'#000000',bold:row===plan.rows[0]||row===plan.rows.at(-1)},style:'范本正文',alignment:c>=1&&c<=3?'right':'left',firstLineIndent:0,spaceBefore:0,spaceAfter:0});}
  // All table cell texts are exact paragraph matches; duplicate occurrences are explicit.
  await command('table-typography','word_format_paragraphs',{targets:buildParagraphTargets(tableParas,tableTargets)},true);
  await select('select-image',MARK.image);await command('image','word_insert_image',{mode:'replace',base64:imageBytes.toString('base64'),width:440,height:440/3},true);
  await command('header','word_set_header_footer',{kind:'header',section:0,type:'primary',mode:'replace',text:'Office 基础样式范本 · DEMO'},true);
  await command('footer','word_set_header_footer',{kind:'footer',section:0,type:'primary',mode:'replace',field:{type:'Page'}},true);
  await command('first-header','word_set_header_footer',{kind:'header',section:0,type:'firstPage',mode:'replace',text:''},true);
  await command('first-footer','word_set_header_footer',{kind:'footer',section:0,type:'firstPage',mode:'replace',text:''},true);
  const tocParas=(await command('toc-anchor-paragraphs','word_list_paragraphs')).paragraphs;const tocAnchor=tocParas.find(p=>p.text===MARK.toc);if(!tocAnchor)throw new Error('Exact TOC marker paragraph missing');
  await command('toc-paragraph','word_insert_field',{type:'TOC',target:'selection',mode:'replace',paragraph:{index:tocAnchor.index,expectedText:tocAnchor.text}},true);
  const document=await command('read-final-document','read_document'),styles=await command('read-final-styles','word_read_styles',{names:Object.values(names)}),fields=await command('read-final-fields','word_read_fields',{target:'body'}),headers=await command('read-final-headers','word_read_headers_footers',{section:0,types:['primary','firstPage'],includeFields:true}),tables=await command('read-final-tables','read_tables'),paragraphs=await command('read-final-paragraphs','word_list_paragraphs');
  if(Object.values(MARK).some(m=>document.text.includes(m)))throw new Error('Insertion marker remains; export stopped');
  if(!headers.items?.some(i=>i.type==='firstPage'&&!String(i.header||'').trim()&&!String(i.footer||'').trim()&&!i.headerFields?.length&&!i.footerFields?.length))throw new Error('First-page header/footer blank readback missing');
  if(!fields.fields?.some(f=>f.type==='TOC')||!headers.items?.some(i=>i.footerFields?.some(f=>f.type==='Page')))throw new Error('Native TOC/PAGE readback missing; export stopped');
  const finalEnvironment=await command('read-final-environment','get_environment');
  if(finalEnvironment.probe?.nativeDiagnostics?.pagination?.count!==5)throw new Error('Native pagination is not five pages; export stopped');
  const native=await command('export-docx','word_export_file',{type:'docx'});if(!native.base64)throw new Error('Native DOCX export missing');
  const bytes=Buffer.from(native.base64,'base64'),destination=path.join(outDir,'word-base-styles.docx');await fs.writeFile(destination,bytes);
  const evidence={...inspectNativeExport(bytes,imageBytes),instance:opt.instance,runId,pluginOnlyWrites:true,docxSha256:hash(bytes),originalPngSha256:hash(imageBytes),readback:{document,styles,fields,headers,tables,paragraphs,environment:finalEnvironment},firstPageSettingWritable:true,pdfAttempted:false};
  await fs.writeFile(path.join(outDir,'word-base-styles.evidence.json'),JSON.stringify(evidence,null,2),'utf8');console.log(JSON.stringify({outDir,passed:evidence.passed,nativePageCount:evidence.nativePageCount,trueOnePointFive:evidence.trueOnePointFive,fontsVerified:evidence.fontsVerified},null,2));if(!evidence.passed)throw new Error('Native DOCX verification incomplete; inspect evidence before claiming success');
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url))main().catch(e=>{console.error(e.message);process.exitCode=1;});
