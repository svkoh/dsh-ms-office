// Business presentation recipes. All charts and tables are editable Office shapes/text.
const W = 960, H = 540;
export const pptThemes = Object.freeze({
  'ppt-clean-blue-deck': Object.freeze({ name:'Warm ivory / matte copper', background:'#FAF9F6', text:'#24282D', accent:'#836344', muted:'#5E6267', surface:'#EFEEEA', secondary:'#8E8C88', grid:'#D7D4CE', risk:'#8A4031' }),
  'ppt-modern-accent': Object.freeze({ name:'Cool white / mist blue', background:'#F8FAFC', text:'#1D2733', accent:'#44657D', muted:'#596674', surface:'#EDF1F4', secondary:'#8797A4', grid:'#D3DBE2', risk:'#8A4031' }),
});
const baseTheme = pptThemes['ppt-clean-blue-deck'];
const C = { ...baseTheme, tint:baseTheme.surface, white:baseTheme.background };
export function textUnits(value) { return [...String(value)].reduce((n,c)=>n+(c==='\n'?0:/[MW@%]/.test(c)?0.85:/[\x20-\x7e]/.test(c)?0.62:1),0); }
function wrap(value,width,size) {
  const out=[];
  for(const paragraph of String(value).split('\n')) {
    let line='';
    const raw=paragraph.match(/[A-Za-z][A-Za-z0-9_-]*|[−-]?\d+(?:[.,]\d+)*(?:%|[A-Za-z]+)?|[^\n]/gu)||[''];
    const tokens=[];
    let opening='';
    for(const token of raw){
      if(/^[（【《「『“‘]$/u.test(token)){opening+=token;continue;}
      if(/^[。！？；：、，,.!?;:）】》」』”’]$/u.test(token)&&tokens.length&&!opening){tokens[tokens.length-1]+=token;continue;}
      tokens.push(opening+token);opening='';
    }
    if(opening){if(tokens.length)tokens[tokens.length-1]+=opening;else tokens.push(opening);}
    for(const token of tokens) {
      if(line&&textUnits(line+token)*size>width-6){out.push(line);line='';}
      line+=token;
    }
    out.push(line);
  }
  return out.join('\n');
}
function txt(role,text,x,y,w,h,size=18,color=C.text,bold=false,fill) { return {type:'text',role,text:wrap(text,w-4,size),x,y,w,h,font:'Microsoft YaHei',size,color,bold,...(fill?{fill}:{}),lineWidth:0,autoFit:false,align:'Left',valign:'Top',margins:{left:2,right:2,top:0,bottom:0}}; }
function rect(role,x,y,w,h,fill=C.tint,lineColor=C.grid) { return {type:'shape',role,shapeType:'Rectangle',text:'',x,y,w,h,font:'Microsoft YaHei',size:14,color:C.text,fill,lineColor,lineWidth:0.5,autoFit:false,align:'Left',valign:'Top',margins:{left:0,right:0,top:0,bottom:0}}; }
function line(role,x,y,w,h=0,color=C.grid) {return {type:'line',role,x,y,w,h,color,lineWidth:1,connectorType:'Straight',dashStyle:'Solid'};}
function table(elements,headers,rows,x,y,widths,rowHeight=48) {
  [headers,...rows].forEach((row,r)=>{let left=x;row.forEach((cell,c)=>{const width=widths[c];elements.push(rect('table-grid',left,y+r*rowHeight,width,rowHeight,r===0?C.tint:C.white));elements.push(txt(r===0?'table-header':'table-cell',cell,left+7,y+r*rowHeight+7,width-14,rowHeight-12,14,C.text,r===0));left+=width;});});
}
function notes(elements,items,x,y,w,step=85) {items.forEach((item,i)=>{elements.push(txt('body',item,x,y+i*step,w,step-10,18));});}
function kpis(elements,items,accent) { items.forEach((k,i)=>{let x=42+i*294;elements.push(txt('kpi-label',k.label,x,150,276,25,16,C.muted));elements.push(txt('kpi-value',k.value,x,182,276,43,30,accent,true));elements.push(txt('kpi-gap',k.gap,x,235,276,48,16));elements.push(line('divider',x,292,276));}); }
function chart(elements,s,accent,{x=100,y=182,w=475,h=205}={}) {
  const max=s.max,base=y+h;elements.push(txt('chart-unit',`单位：${s.unit}`,x,y-31,w,22,14,C.muted));
  for(let t=0;t<=4;t++){const v=max*t/4,yy=base-h*t/4;elements.push(line('chart-grid',x,yy,w));elements.push(txt('chart-axis',Number(v.toFixed(1)).toString(),x-55,yy-9,49,21,14,C.muted));}
  const slot=w/s.categories.length,series=s.series;
  s.categories.forEach((cat,i)=>{elements.push(txt('chart-category',cat,x+i*slot,base+12,slot-5,29,14));series.forEach((ser,j)=>{const bw=Math.min(46,slot/(series.length+1)-5),bx=x+i*slot+9+j*(bw+5),value=ser.values[i],bh=value/max*h;elements.push(rect('chart-mark',bx,base-bh,bw,Math.max(bh,0.5),j===0?accent:'#A7BCC8',j===0?accent:'#A7BCC8'));elements.push(txt('chart-value',String(value),bx-8,base-bh-25,bw+18,22,14,C.text,true));});});
  series.forEach((ser,j)=>{elements.push(rect('chart-legend',x+j*200,base+50,12,12,j===0?accent:'#A7BCC8'));elements.push(txt('chart-label',ser.name,x+19+j*200,base+45,180,24,14));});
}
function validateInput(preset,input) {
  const size=input.nativeSize;if(!size||![size.width,size.height].every(Number.isFinite)||size.width<960||size.height<540||size.width>10000||size.height>10000||Math.abs(size.width/size.height-16/9)>0.001)throw new Error('Business PPT minimumSize is 960x540 native points; 720x405 must be refused, never shrink fonts');
  if(!Array.isArray(input.slides)||input.slides.length!==8)throw new Error('Complete business recipe requires exactly 8 slides');
  const types=preset.id==='ppt-clean-blue-deck'?['kpi','trend','comparison','matrix','actions','timeline','risks','decision']:['evidence','kpi','matrix','flow','roi','timeline','risks','decision'];
  input.slides.forEach((s,i)=>{if(s.kind!==types[i])throw new Error(`Slide ${i+1} requires business layout ${types[i]}`);for(const key of ['title','conclusion','source'])if(typeof s[key]!=='string'||!s[key].trim())throw new Error(`Slide ${i+1} needs ${key}`);
    if(textUnits(s.title)>28)throw new Error('Title exceeds 28 full-width characters');
    if(s.chart){const c=s.chart;if(!c.unit||!Array.isArray(c.categories)||c.categories.length<3||c.categories.length>6||!Number.isFinite(c.max)||c.max<=0||!Array.isArray(c.series)||c.series.length<1||c.series.length>2)throw new Error('Chart requires unit, 3..6 categories, max and 1..2 series');for(const ser of c.series)if(!ser.name||!Array.isArray(ser.values)||ser.values.length!==c.categories.length||ser.values.some(v=>!Number.isFinite(v)||v<0||v>c.max))throw new Error('Invalid quantitative chart values');}
    if(s.table){const t=s.table;if(!Array.isArray(t.headers)||!Array.isArray(t.rows)||!Array.isArray(t.widths)||t.headers.length!==t.widths.length||t.widths.reduce((a,b)=>a+b,0)!==876||t.rows.some(r=>r.length!==t.headers.length)||t.rows.length<2||t.rows.length>5)throw new Error('Editable table needs 2..5 rectangular rows and 876pt total columns');}
  });
  return types;
}
function compilePage(s,index,input,theme) {
  const accent=theme.accent;
  const e=[txt('title',s.title,42,36,876,43,30,C.text,true),txt('conclusion',s.conclusion,42,90,876,48,18,accent,true)];
  if(s.kind==='kpi') {if(s.kpis?.length!==3)throw new Error('Three KPIs with value and target gap required');kpis(e,s.kpis,accent);if(s.table)table(e,s.table.headers,s.table.rows,42,315,s.table.widths,43);else notes(e,s.notes||[],42,315,876,49);}
  else if(['trend','comparison','evidence','roi'].includes(s.kind)){if(!s.chart)throw new Error('Quantitative chart required');chart(e,s.chart,accent);notes(e,s.notes||[],635,165,283,100);}
  else if(s.kind==='flow'){if(s.stages?.length!==4)throw new Error('Four workflow stages required');s.stages.forEach((stage,i)=>{const x=42+i*224;e.push(rect('flow-node',x,165,204,144,C.tint));e.push(txt('flow-title',`${i+1}. ${stage.title}`,x+12,177,180,30,18,accent,true));e.push(txt('body',stage.detail,x+12,219,180,80,18));if(i<3)e.push({...txt('flow-arrow','→',x+205,210,20,28,14,accent),margins:{left:0,right:0,top:0,bottom:0}});});notes(e,s.notes||[],42,326,876,48);}
  else if(s.kind==='timeline'){if(s.milestones?.length!==4)throw new Error('Four accountable milestones required');e.push(line('timeline-axis',64,181,672,0,accent));s.milestones.forEach((m,i)=>{const x=42+i*224;e.push(rect('milestone',x+16,175,12,12,accent,accent));e.push(txt('timeline-date',m.when,x,149,204,24,16,accent,true));e.push(txt('timeline-title',m.title,x,205,204,30,18,C.text,true));e.push(txt('body',`${m.owner}\n${m.exit}`,x,248,204,95,18));});notes(e,s.notes||[],42,365,876,47);}
  else if(s.kind==='decision'){e.push(txt('decision-request',s.request,42,156,876,58,22,accent,true));table(e,s.table.headers,s.table.rows,42,230,s.table.widths,52);notes(e,s.notes||[],42,405,876,43);}
  else {if(!s.table)throw new Error('Evidence/action/risk matrix table required');table(e,s.table.headers,s.table.rows,42,155,s.table.widths,s.table.rowHeight||53);if(s.notes?.length)notes(e,s.notes,42,155+(s.table.rows.length+1)*(s.table.rowHeight||53)+8,876,43);}
  e.push(txt('source',`来源：${s.source}`,42,463,876,18,12,C.muted));e.push(txt('footer',`${input.displayLabel?`${input.displayLabel} · `:''}${index+1}/8`,42,486,876,18,12,C.muted));for(const element of e){
    if(element.type==='text'){
      element.color=['source','footer','kpi-label','chart-unit','chart-axis'].includes(element.role)?theme.muted:['flow-arrow','timeline-date'].includes(element.role)?theme.accent:theme.text;
    } else if(element.type==='line'){
      element.color=element.role==='timeline-axis'?theme.accent:theme.grid;
    } else {
      element.color=theme.text;
      if(['chart-mark','chart-legend'].includes(element.role)){
        element.fill=element.fill===accent?theme.accent:theme.secondary;
        element.lineColor=element.fill;
      } else if(element.role==='milestone'){
        element.fill=theme.accent;element.lineColor=theme.accent;
      } else {
        element.fill=element.fill===C.tint?theme.surface:theme.background;
        element.lineColor=theme.grid;
      }
    }
  }
  e.unshift({...rect('background',0,0,W,H,theme.background,theme.background),lineWidth:0});
  return e;
}
export function buildPptSteps(preset,input,add,plan) {
  const types=validateInput(preset,input),theme=pptThemes[preset.id];
  if(!theme)throw new Error('Unknown business presentation theme');
  plan.pptRecipe={version:2,theme:structuredClone(theme),minimumSize:{width:960,height:540},businessType:preset.id==='ppt-clean-blue-deck'?'operating-review':'project-proposal',layouts:types,businessFacts:structuredClone(input.businessFacts),pages:structuredClone(input.slides),chartImplementation:'editable-shapes',tableImplementation:'editable-cell-shapes',designReferences:input.designReferences||[]};
  plan.preconditions.push('Read nativeSize from ppt_read_presentation_info for this exact instance and stop on mismatch. Minimum native size 960x540pt; never shrink fonts.','Append exactly eight new slides; preserve existing slides and resolve only own added[0..7].id fields.');
  plan.manualChecks.push('Render all eight newly added slides in native PowerPoint; inspect clipping, overlaps, glyphs, chart axes and table legibility.','Charts are editable geometry and labels; tables are editable cell shapes/text. Neither is a native chart or raster screenshot.');
  add('read-native-size','ppt_read_presentation_info',{});add('add-slides','ppt_add_slides',{count:8,clearPlaceholders:true});
  const bind=(i,array=false)=>({path:array?`slideIds.${i}`:'slideId',from:'add-slides',resultPath:['added',i,'id']});
  input.slides.forEach((s,i)=>add(`build-slide-${i}`,'ppt_add_elements',{slideId:null,coordinateSpace:{width:W,height:H},targetSize:{...input.nativeSize},elements:compilePage(s,i,{...input,displayLabel:plan.displayLabel||''},theme)},[bind(i)]));
  const ids=Array(8).fill(null),bindings=ids.map((_,i)=>bind(i,true));add('inspect-new-slides','ppt_inspect_slides',{slideIds:[...ids],withStyles:true},bindings);add('render-new-slides','ppt_render_slides',{slideIds:[...ids],width:W,height:H},bindings);
  return plan;
}
function contrast(a,b) {const lum=h=>{const rgb=[1,3,5].map(i=>parseInt(h.slice(i,i+2),16)/255).map(v=>v<=.04045?v/12.92:((v+.055)/1.055)**2.4);return .2126*rgb[0]+.7152*rgb[1]+.0722*rgb[2];};const x=lum(a),y=lum(b);return(Math.max(x,y)+.05)/(Math.min(x,y)+.05);}
export function checkPptPlan(plan,{error,warning=()=>{}}) {
  const pages=plan.steps.filter(s=>s.action==='ppt_add_elements'),recipe=plan.pptRecipe,theme=pptThemes[plan.presetId];
  if(!recipe||recipe.version!==2){error('Complete business recipe metadata missing');return;}
  if(!theme||JSON.stringify(recipe.theme)!==JSON.stringify(theme)){error('Exact business theme tokens required');return;}
  try { validateInput({id:plan.presetId},{nativeSize:pages[0]?.args.targetSize,slides:recipe.pages}); } catch(caught) { error(caught.message); }
  if(recipe.minimumSize?.width!==960||recipe.minimumSize?.height!==540)error('Explicit minimumSize 960x540 required');
  if(pages.length!==8||plan.steps.filter(s=>s.action==='ppt_add_slides').length!==1||plan.steps.find(s=>s.action==='ppt_add_slides')?.args.count!==8)error('Exactly eight appended business slides required');
  if(plan.steps[0]?.action!=='ppt_read_presentation_info'||plan.steps.at(-2)?.action!=='ppt_inspect_slides'||plan.steps.at(-1)?.action!=='ppt_render_slides')error('Read info before append and inspect/render only new slides required');
  let chartPages=0,tablePages=0,quantPages=0;
  const layouts=new Set();pages.forEach((step,i)=>{const a=step.args,s=recipe.pages?.[i],es=a.elements||[];if(!s){error('Missing business page content');return;}layouts.add(s.kind);
    if(a.targetSize?.width<960||a.targetSize?.height<540||Math.abs(a.targetSize?.width/a.targetSize?.height-16/9)>.001)error('Native minimumSize 960x540 required');
    const text=es.filter(e=>e.text).map(e=>e.text).join(' ');
    if(es.filter(e=>['body','table-cell','kpi-gap','chart-label','decision-request'].includes(e.role)).reduce((n,e)=>n+textUnits(e.text||''),0)<55)error(`Slide ${i+1}: insufficient meaningful business information density`);
    if(!es.some(e=>e.role==='conclusion'&&textUnits(e.text)>12))error(`Slide ${i+1}: evidence-based conclusion missing`);
    if(!es.some(e=>e.role==='source'&&e.text.includes(s.source)))error(`Slide ${i+1}: visible source missing`);
    const expectedFooter=`${plan.displayLabel?`${plan.displayLabel} · `:''}${i+1}/8`;
    if(!es.some(e=>e.role==='footer'&&e.text===expectedFooter))error(`Slide ${i+1}: demo/page/source label missing or differs from displayLabel`);
    if((text.match(/\d+(?:\.\d+)?/g)||[]).length>=6)quantPages++;
    if(s.chart){chartPages++;for(const role of ['chart-axis','chart-category','chart-value','chart-unit','chart-mark','chart-label'])if(!es.some(e=>e.role===role))error(`Slide ${i+1}: editable chart lacks ${role}`);
      for(const ser of s.chart.series)for(const value of ser.values)if(!es.some(e=>e.role==='chart-value'&&e.text===String(value)))error('Chart labels diverge from quantitative source data');
      if(!es.some(e=>e.role==='chart-unit'&&e.text.includes(s.chart.unit)))error('Visible chart unit differs from data unit');
      for(const category of s.chart.categories)if(!es.some(e=>e.role==='chart-category'&&e.text.replaceAll('\n','')===category))error('Chart category labels missing or changed');
      const marks=es.filter(e=>e.role==='chart-mark'),values=s.chart.categories.flatMap((_,ci)=>s.chart.series.map(ser=>ser.values[ci]));
      if(marks.length!==values.length||marks.some((mark,mi)=>Math.abs(mark.h-Math.max(values[mi]/s.chart.max*205,.5))>.01))error('Editable data marks do not match quantitative scale');
    }
    if(s.table){for(const row of s.table.rows)for(const cell of row)if(!es.some(e=>e.role==='table-cell'&&e.text.replaceAll('\n','')===String(cell).replaceAll('\n','')))error('Visible business table diverges from source content');}
    if(es.some(e=>e.role==='table-cell'))tablePages++;
    if(es[0]?.role!=='background'||es.filter(e=>e.role==='background').length!==1)error('One full-slide background must be the first element');
    for(const e of es){
      if(e.role==='background'){
        if(e.type!=='shape'||e.shapeType!=='Rectangle'||e.x!==0||e.y!==0||e.w!==960||e.h!==540||e.fill!==theme.background||e.text!==''||e.lineWidth!==0)error('Background must be the exact full-slide theme rectangle');
        continue;
      }
      if(![e.x,e.y,e.w,e.h].every(Number.isFinite)||e.x<36||e.y<36||e.x+e.w>924||e.y+e.h>504||e.w<0||e.h<0)error(`Slide ${i+1}: geometry outside 36pt safe bounds`);
      if(!['text','shape','line'].includes(e.type))error('Unverified native element type');if(e.type==='line')continue;
      if(e.type==='shape'&&e.shapeType!=='Rectangle')error('Unverified geometric shape');if(e.autoFit!==false)error('Automatic font shrink prohibited');
      if(e.text){const min=e.role==='title'?28:['source','footer'].includes(e.role)?12:['table-cell','table-header','chart-axis','chart-category','chart-value','chart-unit','chart-label'].includes(e.role)?14:['kpi-label','kpi-gap','timeline-date'].includes(e.role)?16:e.role==='flow-arrow'?14:18;
        if(!['Microsoft YaHei','Arial'].includes(e.font)||!Number.isFinite(e.size)||e.size<min)error(`Slide ${i+1}: role-specific font minimum violated`);
        if(e.role==='title'&&e.size>32)error('Title must be 28–32pt');
        const lines=e.text.split('\n'),sx=a.targetSize.width/960,sy=a.targetSize.height/540,m=e.margins||{};
        if(lines.some(t=>textUnits(t)*e.size>(e.w*sx-(m.left||0)-(m.right||0))+.1)||lines.length*e.size*1.25>e.h*sy-(m.top||0)-(m.bottom||0)+.1)error(`Slide ${i+1}: text budget exceeded (${e.role}: ${e.text.slice(0,24)})`);
        const underlay=es.slice(0,es.indexOf(e)).filter(shape=>shape.type==='shape'&&shape.fill&&shape.x<=e.x&&shape.y<=e.y&&shape.x+shape.w>=e.x+e.w&&shape.y+shape.h>=e.y+e.h).at(-1);
        const background=e.fill||underlay?.fill||theme.background;
        if(!/^#[\da-f]{6}$/i.test(e.color)||!/^#[\da-f]{6}$/i.test(background)||contrast(e.color,background)<4.5)error(`Slide ${i+1}: text contrast below 4.5`);
      }
    }
  });
  const facts=recipe.businessFacts, semanticPages=recipe.pages;
  const sum=xs=>Array.isArray(xs)&&xs.every(Number.isFinite)?xs.reduce((a,b)=>a+b,0):NaN;
  const equal=(actual,expected,label)=>{if(!Number.isFinite(actual)||!Number.isFinite(expected)||Math.abs(actual-expected)>.051)error(`Business reconciliation failed: ${label}`);};
  if(!facts||typeof facts!=='object')error('Auditable business facts missing');
  else if(recipe.businessType==='operating-review'){
    equal(sum(semanticPages?.[1]?.chart?.series?.[0]?.values),facts.revenue,'six-month revenue total');
    equal(sum(semanticPages?.[2]?.chart?.series?.[0]?.values),facts.revenue,'group revenue total');
    equal(sum(facts.groupGrossProfit),facts.grossProfit,'group gross profit total');
    equal(facts.grossProfitTarget+sum(facts.bridge),facts.grossProfit,'target-to-actual gross profit bridge');
    equal(facts.revenue*facts.grossMargin/100,facts.grossProfit,'actual gross margin');
    equal(facts.revenueTarget*facts.targetMargin/100,facts.grossProfitTarget,'target gross margin');
    equal((facts.revenue-facts.revenueTarget)*facts.targetMargin/100,facts.bridge?.[0],'incremental revenue gross profit');
  } else if(recipe.businessType==='project-proposal'){
    equal(sum(semanticPages?.[0]?.chart?.series?.[0]?.values),facts.monthlyVolume*facts.baselineRate,'baseline rework count');
    equal(facts.monthlyVolume*(facts.baselineRate-facts.targetRate)*facts.reworkMinutes/60,facts.savedHours,'released monthly hours');
    equal(facts.savedHours*facts.hourlyCost*12*facts.realization/10000,facts.annualGross,'annual realized gross savings');
    equal(facts.annualGross-facts.annualMaintenance,facts.annualNet,'annual net savings');
    if(!Array.isArray(facts.scenarioRealization)||facts.scenarioRealization.length!==3)error('Three explicit ROI realization assumptions required');
    else facts.scenarioRealization.forEach((rate,i)=>equal(facts.savedHours*facts.hourlyCost*12*rate/10000-facts.annualMaintenance,semanticPages?.[4]?.chart?.series?.[0]?.values?.[i],`ROI sensitivity ${i+1}`));
  } else error('Unknown complete business presentation type');
  if(layouts.size<6)error('At least six purposeful business layouts required');if(chartPages<2)error('At least two editable quantitative charts required');if(tablePages<2)error('At least two editable business tables required');if(quantPages<3)error('At least three quantitative pages required');
  const last=recipe.pages?.at(-1);if(last?.kind!=='decision'||!last.request||!last.table||!pages.at(-1)?.args.elements.some(e=>e.role==='decision-request'))error('Explicit approval request, owner, amount and release gates required');
  warning('Static business checks cannot establish actual Office rendering quality; native export and visual acceptance remain required.');
}
