import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { startBridge, cleanupStateDir } from './helpers/server.mjs';

function isolatedState(t) {
  const state=fs.mkdtempSync(path.join(os.tmpdir(),'office-confirm-mode-')),bridges=[];
  t.after(async()=>{for(const bridge of bridges){await bridge.stop();cleanupStateDir(bridge.stateDir);}cleanupStateDir(state);});
  return {state,async start(mode=''){const bridge=await startBridge({env:{DSH_OFFICE_STATE_DIR:state,OFFICE_CONFIRM_MODE:mode}});bridges.push(bridge);return bridge;}};
}
const postMode=(bridge,mode)=>bridge.req('/office/config',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({confirmMode:mode})});

test('confirmation choices survive real bridge restarts without overwriting other plugin settings',async t=>{
  const isolated=isolatedState(t),settings=path.join(isolated.state,'settings.json'),file=path.join(isolated.state,'confirm-mode.json');
  fs.writeFileSync(settings,JSON.stringify({ownedByCompanion:true,port:12345}));
  let bridge=await isolated.start();
  assert.equal((await bridge.req('/office/config')).body.confirmMode,'ask');
  assert.equal((await postMode(bridge,'auto')).body.confirmMode,'auto');
  assert.deepEqual(JSON.parse(fs.readFileSync(file,'utf8')),{confirmMode:'auto'});
  await bridge.stop();bridge=await isolated.start();
  assert.equal((await bridge.req('/office/config')).body.confirmMode,'auto');
  const heartbeat=await bridge.req('/office/heartbeat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({host:'Word',instanceId:'mode-test'})});
  assert.equal(heartbeat.body.confirmMode,'auto');
  assert.equal((await postMode(bridge,'ask')).body.confirmMode,'ask');
  await bridge.stop();bridge=await isolated.start();
  assert.equal((await bridge.req('/office/config')).body.confirmMode,'ask');
  assert.deepEqual(JSON.parse(fs.readFileSync(settings,'utf8')),{ownedByCompanion:true,port:12345});
  assert.deepEqual(fs.readdirSync(isolated.state).filter(name=>name.startsWith('.confirm-mode.')),[]);
});

test('valid explicit environment wins at startup; invalid environment does not erase a persisted choice',async t=>{
  const isolated=isolatedState(t),file=path.join(isolated.state,'confirm-mode.json');
  for(const [saved,environment,expected] of [['auto','ask','ask'],['ask','auto','auto'],['auto','invalid','auto'],['auto','','auto']]){
    fs.writeFileSync(file,JSON.stringify({confirmMode:saved}));const bridge=await isolated.start(environment);
    assert.equal((await bridge.req('/office/config')).body.confirmMode,expected);
    assert.equal(JSON.parse(fs.readFileSync(file,'utf8')).confirmMode,saved);await bridge.stop();
  }
});

test('invalid or malformed stored configuration safely defaults to ask',async t=>{
  const isolated=isolatedState(t),file=path.join(isolated.state,'confirm-mode.json');
  for(const saved of ['{','null',JSON.stringify({confirmMode:'invalid'}),JSON.stringify({unrelated:true})]){
    fs.writeFileSync(file,saved);const bridge=await isolated.start();
    assert.equal((await bridge.req('/office/config')).body.confirmMode,'ask');await bridge.stop();
  }
});

test('failed atomic persistence returns HTTP failure and keeps the previous in-memory mode',async t=>{
  const isolated=isolatedState(t);fs.mkdirSync(path.join(isolated.state,'confirm-mode.json'));
  const bridge=await isolated.start(),result=await postMode(bridge,'auto');
  assert.equal(result.status,500);assert.equal(result.body.ok,false);assert.equal(result.body.confirmMode,'ask');
  assert.equal((await bridge.req('/office/config')).body.confirmMode,'ask');
  assert.deepEqual(fs.readdirSync(isolated.state).filter(name=>name.startsWith('.confirm-mode.')),[]);
});

test('invalid POST leaves the persisted mode and active policy unchanged',async t=>{
  const isolated=isolatedState(t),bridge=await isolated.start();await postMode(bridge,'auto');
  const file=path.join(isolated.state,'confirm-mode.json'),saved=fs.readFileSync(file,'utf8');
  assert.equal((await postMode(bridge,'invalid')).status,400);
  assert.equal((await bridge.req('/office/config')).body.confirmMode,'auto');assert.equal(fs.readFileSync(file,'utf8'),saved);
});

function paneHarness(handler) {
  const requests=[],elements=new Map(),buttons=['ask','auto'].map(mode=>({dataset:{mode},attributes:{},disabled:false,setAttribute(k,v){this.attributes[k]=v;}}));
  const element=id=>{if(!elements.has(id))elements.set(id,{className:'',textContent:'',title:'',setAttribute(){},addEventListener(){},querySelectorAll:()=>buttons});return elements.get(id);};
  const timers=[];let reloads=0;
  let onReady;const sandbox={
    window:{__CODE_VERSION__:'same'},location:{origin:'http://127.0.0.1:39999',reload(){reloads++;}},
    document:{getElementById:element,documentElement:{},head:{appendChild(){}}},
    Office:{version:'test',HostType:{Word:'Word'},context:{host:'Word',document:{url:'',title:''}},onReady(callback){onReady=callback;}},
    fetch:async(url,init={})=>{requests.push({url,init});return handler(url,init);},
    setInterval:()=>0,setTimeout:callback=>{timers.push(callback);return timers.length;},clearTimeout(){},console,
  };vm.createContext(sandbox);vm.runInContext(fs.readFileSync(new URL('../server/taskpane.js',import.meta.url),'utf8'),sandbox);
  return {sandbox,requests,buttons,timers,reloads:()=>reloads,run:code=>vm.runInContext(code,sandbox),ready:()=>onReady({host:'Word'})};
}
const response=body=>({ok:true,json:async()=>body});
const tick=()=>new Promise(resolve=>setImmediate(resolve));

test('idle heartbeat syncs valid cross-pane modes without echoing a config POST',async()=>{
  let mode='auto';const pane=paneHarness(()=>response({ok:true,codeVersion:'same',confirmMode:mode}));
  await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'auto');
  assert.equal(pane.buttons[1].attributes['aria-checked'],'true');
  mode='ask';await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
  mode='invalid';await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
  assert.ok(pane.requests.every(r=>!r.url.endsWith('/office/config')));
});

test('heartbeat never changes a command policy during a write or approval, then catches up when idle',async()=>{
  let mode='auto';const pane=paneHarness(()=>response({ok:true,codeVersion:'same',confirmMode:mode}));
  pane.run('busy=true');await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
  pane.run('busy=false;approvalActive=true');await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
  pane.run('approvalActive=false');await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'auto');
  mode='ask';pane.run('busy=true');await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'auto');
  pane.run('busy=false');await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
});

test('initial config response respects the same in-flight guard as heartbeat',async()=>{
  const pane=paneHarness(url=>response(url.includes('/office/i18n')?{pane:{},lang:'en'}:{ok:true,codeVersion:'same',confirmMode:'auto'}));
  pane.run('busy=true');pane.ready();await tick();assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
  pane.run('busy=false');await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'auto');
});

test('UI mode choice waits for server commit and prevents polling or heartbeat changes while saving',async()=>{
  let commit;const pane=paneHarness(url=>url.endsWith('/office/config')?new Promise(resolve=>{commit=resolve;}):response({ok:true,codeVersion:'same',confirmMode:'auto'}));
  const changing=pane.run("setConfirmMode('auto')");await tick();
  assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');assert.ok(pane.buttons.every(b=>b.disabled));
  await pane.run('poll()');assert.ok(!pane.requests.some(r=>r.url.includes('/office/poll')||r.url.includes('/office/actions-version')));
  await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
  commit(response({ok:true,confirmMode:'auto'}));assert.equal(await changing,true);
  assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'auto');assert.ok(pane.buttons.every(b=>!b.disabled));
  assert.equal(pane.requests.filter(r=>r.url.endsWith('/office/config')).length,1);
});

test('failed UI save keeps the old cache and selection; active commands cannot POST mode changes',async()=>{
  const pane=paneHarness(()=>({ok:false,status:500,json:async()=>({ok:false})}));
  pane.run('syncModeUI()');assert.equal(await pane.run("setConfirmMode('auto')"),false);
  assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');assert.equal(pane.buttons[0].attributes['aria-checked'],'true');
  const requests=pane.requests.length;pane.run('busy=true');assert.equal(await pane.run("setConfirmMode('auto')"),false);
  pane.run('busy=false;approvalActive=true');assert.equal(await pane.run("setConfirmMode('auto')"),false);
  assert.equal(pane.requests.length,requests);
});


test('queued heartbeat mode applies at the next idle boundary without another heartbeat',async()=>{
  const pane=paneHarness(()=>response({ok:true,codeVersion:'same',confirmMode:'auto'}));
  pane.run('busy=true');await pane.run('heartbeat()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
  pane.run('busy=false;approvalActive=true;flushIdleState()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');
  pane.run('approvalActive=false;flushIdleState()');assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'auto');
});

test('phase-aligned heartbeat queues reload until poll posts the in-flight command result',async()=>{
  let finishWrite,finishResult;const expiresAt=Date.now()+60000;
  const pane=paneHarness(url=>{
    if(url.endsWith('/office/heartbeat'))return response({ok:true,codeVersion:'new',confirmMode:'ask'});
    if(url.includes('/office/poll?'))return response({commandId:'held-write',action:'write',args:{},expiresAt});
    if(url.endsWith('/start'))return response({ok:true,commandId:'held-write',status:'executing'});
    if(url.endsWith('/office/result'))return new Promise(resolve=>{finishResult=resolve;});
    return response({ok:true});
  });
  pane.sandbox.window.__ACTIONS__={write:{destructive:true}};
  pane.sandbox.window.__EXECUTE__=()=>new Promise(resolve=>{finishWrite=resolve;});
  pane.run("applyConfirmMode('auto')");const running=pane.run('poll()');await tick();
  assert.equal(typeof finishWrite,'function');await pane.run('heartbeat()');
  assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'auto');assert.equal(pane.timers.length,0);
  finishWrite({ok:true,result:{written:true}});await tick();assert.equal(typeof finishResult,'function');
  assert.equal(pane.timers.length,0);assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'auto');
  finishResult(response({ok:true}));await running;
  assert.equal(pane.sandbox.window.__CONFIRM_MODE__,'ask');assert.equal(pane.timers.length,1);
  const requests=pane.requests.length;await pane.run('poll()');assert.equal(pane.requests.length,requests);
  pane.timers[0]();assert.equal(pane.reloads(),1);
});

test('reload callback rechecks in-flight approval and retries at the next idle boundary',async()=>{
  const pane=paneHarness(()=>response({ok:true,codeVersion:'new',confirmMode:'ask'}));
  await pane.run('heartbeat()');assert.equal(pane.timers.length,1);
  pane.run('approvalActive=true');pane.timers[0]();assert.equal(pane.reloads(),0);
  pane.run('approvalActive=false;flushIdleState()');assert.equal(pane.timers.length,2);
  pane.timers[1]();assert.equal(pane.reloads(),1);
});
