import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {loadActions} from './helpers/actions.mjs';

test('unknown action suggests existing canonical name without executing an alias',async()=>{
  const {win}=loadActions();
  const out=await win.__EXECUTE__('word_read_table_borders',{});
  assert.equal(out.code,'unknown_action');
  assert.match(out.error,/read_table_borders/);
  assert.match(out.error,/GET \/office\/capabilities/);
});

test('dispatcher preserves field verification warnings when after-state verification also fails',async()=>{
  const {win,A}=loadActions(); let reads=0;
  win.__CONFIRM_MODE__='auto';
  A.hosts.Word.actions.read_document=async()=>({ok:true,result:{text:'actual'}});
  win.__USER_ACTIONS__={contract_probe:{hosts:['Word'],destructive:true,
    impl:async()=>({ok:true,result:{verified:false,verificationWarnings:[{field:'font.color',status:'mismatch'}]}}),
    preview:async()=>++reads===1?{ok:true,result:{text:'before'}}:{ok:false,error:'read failed'},
  }};
  const out=await win.__EXECUTE__('contract_probe',{});
  assert.equal(out.ok,true); assert.equal(out.result.verified,false);
  assert.equal(out.result.verificationWarnings.length,2);
  assert.equal(out.result.verificationWarnings[0].field,'font.color');
  assert.equal(out.result.verificationWarnings[1].stage,'afterState');
});

for(const lang of ['zh','en']) {
  test(`${lang} capability metadata exposes Word result paths and repair constraints`,()=>{
    const {capabilities:c}=JSON.parse(fs.readFileSync(new URL(`../i18n/${lang}.json`,import.meta.url),'utf8'));
    for(const name of ['read_tables','read_table_borders']) assert.deepEqual(c[name].args,{});
    for(const name of ['read_tables','read_table_borders','word_format_paragraphs','word_read_styles','word_set_styles','word_render_pages','get_environment']) assert.equal(typeof c[name].returns,'object',name);
    assert.match(c.word_format_paragraphs.args.targets,/style/);
    assert.match(c.word_set_styles.args.styles,/type/);
    assert.match(c.word_set_styles.returns.result,/mismatch/);
    assert.match(c.get_environment.returns.result,/probe.nativeDiagnostics.pagination/);
    assert.doesNotMatch(c.set_paragraph_format.args.lineSpacing,/18\s*=>\s*1\.5/);
  });
}
