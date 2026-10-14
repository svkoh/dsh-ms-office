import {test} from 'node:test';
import assert from 'node:assert/strict';
import {countNativeListParagraphs,buildAtlasPlan,buildHtml,isOwnedTemplate,resolveStyleNames,buildParagraphTargets,DEMO} from '../scripts/create-word-style-atlas.mjs';

test('Word atlas has five logical pages and consistent shared demo totals',()=>{
 const plan=buildAtlasPlan();assert.equal(plan.pages.length,5);assert.equal(plan.logicalPages,5);assert.equal(DEMO.budget.reduce((a,b)=>a+b),300);assert.equal(DEMO.actual.reduce((a,b)=>a+b),290);assert.deepEqual(plan.rows.at(-1),['合计','300','290','10','96.7%']);
 const html=buildHtml(plan);assert.match(html,/<ul><li>/);assert.match(html,/<ol><li>/);assert.doesNotMatch(html,/<table|<img|TOC.*fldChar/);assert.equal(plan.pages[0].paragraphs[0].role,'Title');assert.equal(plan.pages.filter(p=>p.paragraphs.some(t=>t.role==='Heading1')).length,3);
});
test('write ownership guard refuses unrelated nonempty documents',()=>{
 assert.equal(isOwnedTemplate(''),true);assert.equal(isOwnedTemplate('会议纪要：预算300万元，实际290万元'),false);assert.equal(isOwnedTemplate('Office 文档基础样式范本 DEMO 300 290 目录 三、图片与页眉页脚'),true);
});
test('global style resolution requires localized semantic definitions',()=>{
 const styles=['正文','标题','标题 1','标题 2','标题 3','题注'].map(name=>({name}));assert.deepEqual(resolveStyleNames(styles),{Normal:'正文',Title:'标题',Heading1:'标题 1',Heading2:'标题 2',Heading3:'标题 3',Caption:'题注'});assert.throws(()=>resolveStyleNames(styles.slice(0,-1)),/Caption/);
});
test('format targets bind exact live indexes/text and handle repeated cells without reuse',()=>{
 const read={paragraphs:[{index:4,text:'按计划\r'},{index:8,text:'按计划'},{index:13,text:'300'}]},specs=[{text:'按计划',font:{bold:true}},{text:'按计划',firstLineIndent:0},{text:'300',alignment:'right'}];
 assert.deepEqual(buildParagraphTargets(read,specs),[{index:4,expectedText:'按计划\r',font:{bold:true}},{index:8,expectedText:'按计划',firstLineIndent:0},{index:13,expectedText:'300',alignment:'right'}]);assert.throws(()=>buildParagraphTargets(read,[...specs,{text:'按计划'}]),/Exact paragraph not found/);
});

test("native list evidence follows inherited numbering and rejects glyphs, disabled or dangling IDs",()=>{
 const numbering='<w:numbering><w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num></w:numbering>';const styles='<w:style w:styleId="Base"><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style><w:style w:styleId="Child"><w:basedOn w:val="Base"/></w:style><w:style w:styleId="Loop"><w:basedOn w:val="Loop"/></w:style>';
 const paragraph=(style,id,text="Example")=>'<w:p><w:pPr>'+(style?'<w:pStyle w:val="'+style+'"/>':'')+(id?'<w:numPr><w:numId w:val="'+id+'"/></w:numPr>':'')+'</w:pPr><w:r><w:t>'+text+'</w:t></w:r></w:p>';
 assert.equal(countNativeListParagraphs(paragraph("Child")+paragraph(null,"7")+paragraph("Child","0")+paragraph(null,"404")+paragraph("Loop")+paragraph(null,null,"• fake bullet"),styles,numbering),2);
});
