import {test} from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import fs from "node:fs";
const source = fs.readFileSync(new URL("../server/actions/word.js",import.meta.url),"utf8");
const plain = x => JSON.parse(JSON.stringify(x));
function fixture({supported=true,missingStyles=false,missingFields=false,missingSecondFooter=false,missingStyleFont=false,missingFarEast=false,nativeHtml=false,missingClear=false,createDocument=false,missingBreakRange=false,nativeBodyPreset=false,nativeToc=false,failRead=false}={}) {
  let writes=0,runs=0; const pending=new Set(),events=[];
  function tracked(data) {
    let ready=false; const proxy=new Proxy({load(){pending.add(()=>{ready=true;});}}, {
      has(o,k) { return k in data || k in o; },
      get(o,k) { if (k in o) return o[k]; assert.equal(ready,true,"readback must load and sync"); return data[k]; },
      set(o,k,v) {writes++;events.push(k);data[k]=v;ready=false;return true;}
    }); return proxy;
  }
  function style(name,type="Paragraph",fontName="Arial",fontSize=11) {return {nameLocal:name,type,builtIn:name!=="范本正文",isNullObject:false,load(){},font:tracked({name:fontName,nameAscii:fontName,nameOther:fontName,nameBidirectional:fontName,...(missingFarEast?{}:{nameFarEast:fontName==="Arial"?"等线":fontName}),size:fontSize,color:"#000000",bold:false,italic:false}),paragraphFormat:tracked({alignment:"Left",spaceBefore:0,spaceAfter:0,lineSpacing:name==="范本正文"?18:12,leftIndent:0,rightIndent:0,firstLineIndent:name==="范本正文"?24:0})};}
  const styles=[style("正文"),style("标题 1"),style("标题"),style("字符","Character")];
  if (missingStyleFont) delete styles[1].font;
  const collection={items:styles,load(){},getByNameOrNullObject(n){return styles.find(s=>s.nameLocal===n)||{isNullObject:true,load(){}};}};
  function range(noField=false) {const r={text:"old",load(){},getRange(){return r;},insertText(text,location){writes++;events.push(["rangeText",text,location]);return r;},fields:{items:[],load(){}}};
    if (!noField) r.insertField=(location,type,switches,removeFormatting)=>{
      writes++; events.push([location,type,switches,removeFormatting]);
      let loaded=false;
      const f={result:{get text(){assert.equal(loaded,true);return type==="TOC"?"一、示例\t1":"7";}},get type(){assert.equal(loaded,true);return type;},get code(){assert.equal(loaded,true);return type.toUpperCase()+" "+switches;},load(fields){assert.equal(fields,"type,code,result/text");pending.add(()=>{loaded=true;});},updateResult(){writes++;}};r.fields.items.push(f);return f;
    };return r;
  }
  const selection=range(missingFields),bodyRange=range(missingFields);
  const footers=[range(missingFields),range(missingSecondFooter)];
  const bodies=footers.map(r=>({text:"old",load(){},getRange:()=>r,insertText(){writes++;}}));
  const ctx={document:{body:{text:"old",load(){},getRange:()=>bodyRange},getSelection:()=>selection,sections:{items:bodies.map(b=>({getHeader:()=>b,getFooter:()=>b})),load(){}}},async sync(){if(failRead)throw new Error("read failed");for(const f of pending)f();pending.clear();}};
  ctx.document.body.paragraphs={load(){},items:["封面 / 目录 / 章节","目录"].map((text,index)=>({text,...(missingBreakRange?{}:{getRange:(location)=>{if(location)events.push(["paragraphRange",index,location]);return {...range(missingFields),insertBreak(type,location){writes++;events.push(["paragraphBreak",index,type,location]);}};}})}))};
  if(nativeBodyPreset){selection.delete=()=>{};for(const p of ctx.document.body.paragraphs.items){let builtin="Normal";Object.defineProperty(p,"styleBuiltIn",{get(){return builtin;},set(value){writes++;events.push(["restoreBodyStyle",value]);builtin=value;}});Object.defineProperty(p,"style",{set(value){writes++;events.push(["applyBodyStyle",value]);builtin="Other";}});}ctx.document.body.insertOoxml=(xml,location)=>{writes++;events.push(["importStyle",xml,location]);pending.add(()=>styles.push(style("范本正文","Paragraph","宋体",12)));return {delete(){writes++;events.push("deleteImportedRange");}};};}
  if(nativeToc)ctx.document.tablesOfContents={add(range,options){writes++;events.push(["nativeTocCollection",options]);bodyRange.insertField("Start","TOC",'\\o "1-3"',true);}};
  const page=tracked({topMargin:72,bottomMargin:72,leftMargin:72,rightMargin:72,headerDistance:36,footerDistance:36,pageWidth:612,pageHeight:792,differentFirstPageHeaderFooter:false});for(const section of ctx.document.sections.items)section.pageSetup=page;
  if(createDocument)ctx.application={createDocument(...args){assert.equal(args.length,0);writes++;events.push("createDocument");return {open(){writes++;events.push("open");}};}};
  if(nativeHtml){ctx.document.body.load=()=>{};ctx.document.body.text="old";ctx.document.body.insertHtml=(html,location)=>{writes++;events.push(["html",html,location]);};if(!missingClear)ctx.document.body.clear=()=>{writes++;events.push("clear");};}
  if(!missingStyles)ctx.document.getStyles=()=>collection;
  const window={__OFFICE_ACTIONS__:{util:{okResult:result=>({ok:true,result}),errResult:(code,error)=>({ok:false,code,error}),mapAlignment:x=>({left:"Left",center:"Centered",right:"Right",justify:"Justified"}[x])},hosts:{},isSetSupported:()=>supported}};
  vm.runInNewContext(source,{window,fetch:async(url,options)=>{events.push(["render",url,JSON.parse(options.body)]);return {ok:true,json:async()=>({ok:true,result:{pages:[{index:1,base64:"native-page"}]}})};},Word:{InsertLocation:{replace:"Replace",end:"End",start:"Start",before:"Before",after:"After"},run:async f=>{runs++;return f(ctx);}},Office:{context:{platform:"PC"}}});
  return {host:window.__OFFICE_ACTIONS__.hosts.Word,writes:()=>writes,runs:()=>runs,events,styles,ctx};
}
const patches={styles:[{name:"正文",font:{name:"宋体",size:12},paragraphFormat:{firstLineIndent:24,lineSpacing:18}},{name:"标题 1",font:{name:"黑体",size:14}},{name:"标题",font:{name:"黑体",size:18}}]};
test("global semantic style definitions preserve absent fields and return loaded actual values",async()=>{
  const f=fixture();const out=await f.host.actions.word_set_styles(patches);assert.equal(out.ok,true);assert.equal(out.result.global,true);assert.equal(out.result.updated,3);
  const [normal,heading,title]=out.result.styles;assert.equal(normal.font.name,"宋体");assert.equal(normal.font.size,12);assert.equal(normal.font.nameFarEast,"宋体");assert.equal(heading.font.nameFarEast,"黑体");assert.equal(normal.paragraphFormat.firstLineIndent,24);assert.equal(normal.paragraphFormat.lineSpacing,18);assert.equal(normal.font.color,"#000000");assert.equal(heading.font.name,"黑体");assert.equal(heading.font.size,14);assert.equal(title.font.size,18);
});
test("whole style set preflights before any mutation",async()=>{
  for(const bad of [{name:"Missing",font:{size:14}},{name:"字符",font:{size:14}},{name:"标题 1",font:{size:NaN}},{name:"标题 1",paragraphFormat:{lineSpacingMultiple:1.5}},{name:"正文",font:{size:14}}]){
    const f=fixture();const out=await f.host.actions.word_set_styles({styles:[patches.styles[0],bad]});assert.equal(out.ok,false);assert.equal(f.writes(),0);
  }
  const f=fixture({failRead:true});assert.equal((await f.host.actions.word_set_styles(patches)).code,"execution");assert.equal(f.writes(),0);
});
test("missing styles requirement or function never writes",async()=>{
  for(const config of [{supported:false},{missingStyles:true},{missingStyleFont:true},{missingFarEast:true}]){const f=fixture(config);assert.equal((await f.host.actions.word_set_styles(patches)).code,"requirement");assert.equal(f.writes(),0);}
});
test("localized style discovery and preview are read-only",async()=>{
  const f=fixture();const all=await f.host.actions.word_read_styles({});assert.equal(all.ok,true);assert.deepEqual(plain(all.result.styles.map(s=>s.name)),["正文","标题 1","标题"]);
  const pre=await f.host.previewReaders.word_set_styles(patches);assert.equal(pre.result.styles[0].font.size,11);assert.equal(f.writes(),0);
});
test("TOC and Page insert actual native fields and return synced code/type/result",async()=>{
  for(const type of ["TOC","Page","NumPages"]){const f=fixture();const pre=await f.host.previewReaders.word_insert_field({type,target:"body",mode:"prepend"});assert.equal(pre.ok,true);assert.equal(f.writes(),0);
    const out=await f.host.actions.word_insert_field({type,target:"body",mode:"prepend"});assert.equal(out.ok,true);assert.equal(out.result.nativeField,true);assert.equal(out.result.field.type,type);assert.match(out.result.field.code,new RegExp(type,"i"));assert.equal(f.events[0][0],"Start");assert.equal(f.events[0][3],true);
    const read=await f.host.actions.word_read_fields({target:"body"});assert.equal(read.result.fields.length,1);assert.equal(read.result.fields[0].type,type);
  }
});
test("invalid field and missing API refuse without mutation",async()=>{
  for(const args of [{type:"MacroButton"},{type:"TOC",switches:"bad\nline"},{type:"Page",target:"body",mode:"after"}]){const f=fixture();assert.equal((await f.host.actions.word_insert_field(args)).code,"bad_args");assert.equal(f.writes(),0);}
  for(const config of [{supported:false},{missingFields:true}]){const f=fixture(config);assert.equal((await f.host.actions.word_insert_field({type:"Page"})).code,"requirement");assert.equal(f.writes(),0);}
});
test("header/footer native PAGE preflights all section APIs then returns field readback",async()=>{
  const args={kind:"footer",section:"all",field:{type:"Page"}};const f=fixture();const out=await f.host.actions.word_set_header_footer(args);assert.equal(out.ok,true);assert.equal(out.result.sections.length,2);assert.equal(out.result.sections[0].field.type,"Page");assert.equal(out.result.sections[0].field.text,"7");
  const missing=fixture({missingSecondFooter:true});assert.equal((await missing.host.actions.word_set_header_footer(args)).code,"requirement");assert.equal(missing.writes(),0);
  const conflict=fixture();assert.equal((await conflict.host.actions.word_set_header_footer({...args,text:"fake"})).code,"bad_args");assert.equal(conflict.writes(),0);
});

test("header/footer field preview is read-only and post-write preview reports native PAGE code",async()=>{
  const f=fixture(),args={kind:"footer",section:"all",field:{type:"Page"}};
  const pre=await f.host.previewReaders.word_set_header_footer(args);assert.equal(pre.ok,true);assert.equal(f.writes(),0);assert.equal(pre.result.items[0].footerFields.length,0);
  await f.host.actions.word_set_header_footer(args);const after=await f.host.previewReaders.word_set_header_footer(args);assert.equal(after.ok,true);assert.equal(after.result.items[0].footerFields[0].type,"Page");assert.match(after.result.items[0].footerFields[0].code,/PAGE/);
  const low=fixture({supported:false});assert.equal((await low.host.actions.word_set_header_footer(args)).code,"requirement");assert.equal(low.writes(),0);
});

test("complex body HTML replacement clears then appends through native APIs and preflights missing clear",async()=>{
 const args={html:"<p>Example</p>",target:"body",mode:"replace"};const f=fixture({nativeHtml:true});const out=await f.host.actions.word_insert_html(args);assert.equal(out.ok,true);assert.deepEqual(plain(f.events),["clear",["html","<p>Example</p>","End"]]);
 const missing=fixture({nativeHtml:true,missingClear:true});assert.equal((await missing.host.actions.word_insert_html(args)).code,"requirement");assert.equal(missing.writes(),0);
});

test("blank creation uses native createDocument with zero payload and rejects uploads/missing API",async()=>{
 const f=fixture({createDocument:true});const out=await f.host.actions.word_create_document({});assert.equal(out.ok,true);assert.equal(out.result.blank,true);assert.equal(out.result.requiresNewInstance,true);assert.deepEqual(plain(f.events),["createDocument","open"]);
 for(const config of [{supported:false},{createDocument:false}]){const missing=fixture(config);assert.equal((await missing.host.actions.word_create_document({})).code,"requirement");assert.equal(missing.writes(),0);}
 const upload=fixture({createDocument:true});assert.equal((await upload.host.actions.word_create_document({base64:"ZIP"})).code,"bad_args");assert.equal(upload.writes(),0);
});
test("native page dimensions and first-page footer flag validate then load actual readback",async()=>{
 const f=fixture();const out=await f.host.actions.word_set_page_setup({sections:[0],pageWidth:595.3,pageHeight:841.9,differentFirstPageHeaderFooter:true});assert.equal(out.ok,true);assert.equal(out.result.pageSetup[0].values.pageWidth,595.3);assert.equal(out.result.pageSetup[0].values.pageHeight,841.9);assert.equal(out.result.pageSetup[0].values.differentFirstPageHeaderFooter,true);
 for(const args of [{pageWidth:0},{pageHeight:-1},{differentFirstPageHeaderFooter:"true"}]){const bad=fixture();assert.equal((await bad.host.actions.word_set_page_setup(args)).code,"bad_args");assert.equal(bad.writes(),0);}
});

test("page break binds exact paragraph text/index rather than an earlier substring",async()=>{
 const args={type:"page",location:"beforeSelection",paragraph:{index:1,expectedText:"目录"}};const f=fixture();assert.equal((await f.host.actions.word_insert_break(args)).ok,true);assert.deepEqual(plain(f.events),[["paragraphBreak",1,"Page","Before"]]);
 const changed=fixture();assert.equal((await changed.host.actions.word_insert_break({...args,paragraph:{index:0,expectedText:"目录"}})).code,"bad_args");assert.equal(changed.writes(),0);const missing=fixture({missingBreakRange:true});assert.equal((await missing.host.actions.word_insert_break(args)).code,"requirement");assert.equal(missing.writes(),0);
});

test("bounded native body preset imports one reusable style and retains it before anchor cleanup",async()=>{const f=fixture({nativeBodyPreset:true}),args={styles:[{name:"范本正文",preset:"body-1.5"}]};const preview=await f.host.previewReaders.word_set_styles(args);assert.equal(preview.ok,true);assert.equal(f.writes(),0);const out=await f.host.actions.word_set_styles(args);assert.equal(out.ok,true);assert.equal(out.result.created,true);assert.equal(out.result.styles[0].font.nameFarEast,"宋体");assert.equal(out.result.styles[0].font.size,12);assert.equal(out.result.styles[0].paragraphFormat.firstLineIndent,24);assert.match(f.events[0][1],/w:line="360" w:lineRule="auto"/);const parts=[...f.events[0][1].match(/<w:body>([\s\S]*?)<\/w:body>/)[1].matchAll(/<w:p(?:\/>|>[\s\S]*?<\/w:p>)/g)].map(m=>m[0]);const anchor=parts.findIndex(p=>p.includes("DSH_NATIVE_BODY_STYLE_ANCHOR"));assert.ok(anchor>0 && anchor<parts.length-1,"styled import paragraph must have both surrounding paragraph boundaries");assert.equal(f.events[0][2],"End");assert.deepEqual(plain(f.events.slice(1)),[["applyBodyStyle","范本正文"],"deleteImportedRange",["restoreBodyStyle","Normal"]]);assert.equal(f.events.includes("lineSpacing"),false);});
test("body preset refuses arbitrary OOXML or alternate names before entering Office",async()=>{for(const raw of [{name:"Other",preset:"body-1.5"},{name:"范本正文",preset:"body-1.5",ooxml:"arbitrary"}]){const f=fixture();assert.equal((await f.host.actions.word_set_styles({styles:[raw]})).code,"bad_args");assert.equal(f.runs(),0);assert.equal(f.writes(),0);}});
test("native page rendering requires explicit bounded count and sends only current native body text",async()=>{for(const count of [undefined,0,11,1.5,"5"]){const f=fixture();assert.equal((await f.host.actions.word_render_pages({expectedPageCount:count})).code,"bad_args");assert.equal(f.runs(),0);}const f=fixture();const out=await f.host.actions.word_render_pages({expectedPageCount:5});assert.equal(out.ok,true);assert.equal(out.result.pages[0].index,1);assert.deepEqual(plain(f.events),[["render","/office/native-word-pages",{bodyText:"old",expectedPageCount:5}]]);assert.equal(f.writes(),0);});

test("replacing an identical native footer PAGE field reuses it without reinsertion or explicit update",async()=>{const f=fixture(),args={kind:"footer",section:0,field:{type:"Page"}};assert.equal((await f.host.actions.word_set_header_footer(args)).ok,true);f.ctx.document.sections.items[0].getFooter("Primary").text="7";const before=f.writes(),events=f.events.length;const out=await f.host.actions.word_set_header_footer(args);assert.equal(out.ok,true);assert.equal(out.result.sections[0].field.type,"Page");assert.equal(f.writes(),before);assert.equal(f.events.length,events);});

test("whole paragraph field anchor rejects stale text before writes",async()=>{const f=fixture();const out=await f.host.actions.word_insert_field({type:"TOC",paragraph:{index:1,expectedText:"changed"}});assert.equal(out.code,"bad_args");assert.equal(f.writes(),0);});
test("native desktop TOC clears only anchored paragraph content and reads actual generated field cache",async()=>{const f=fixture({nativeToc:true});const out=await f.host.actions.word_insert_field({type:"TOC",mode:"replace",paragraph:{index:1,expectedText:"目录"}});assert.equal(out.ok,true);assert.equal(out.result.nativeTocCollection,true);assert.equal(out.result.field.type,"TOC");assert.match(out.result.field.text,/一、示例/);assert.ok(f.events.some(e=>e[0]==="paragraphRange"&&e[2]==="Content"));assert.ok(f.events.some(e=>e[0]==="rangeText"&&e[1]===""&&e[2]==="Replace"));const add=f.events.find(e=>e[0]==="nativeTocCollection");assert.equal(add[1].upperHeadingLevel,1);assert.equal(add[1].lowerHeadingLevel,3);const invalid=fixture({nativeToc:true});assert.equal((await invalid.host.actions.word_insert_field({type:"TOC",switches:"custom",paragraph:{index:1,expectedText:"目录"}})).code,"bad_args");assert.equal(invalid.writes(),0);});
