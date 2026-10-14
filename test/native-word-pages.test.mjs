import { test } from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { startBridge, cleanupStateDir } from './helpers/server.mjs';
const require=createRequire(import.meta.url);
const {renderNativeWordPages,validateNativeWordPagesRequest,canonicalBody,documentFingerprint,RENDERER}=require('../server/native-word-pages.cjs');
const request={bodyText:'标题\r\n正文\r\n',expectedPageCount:1};
function pngHeader(width=100,height=140){const bytes=Buffer.alloc(24);Buffer.from('89504e470d0a1a0a','hex').copy(bytes);bytes.write('IHDR',12,'ascii');bytes.writeUInt32BE(width,16);bytes.writeUInt32BE(height,20);return bytes.toString('base64');}
function validOutput(input=request){return{ok:true,result:{renderer:RENDERER,documentFingerprint:documentFingerprint(input.bodyText),pageCount:input.expectedPageCount,pages:Array.from({length:input.expectedPageCount},(_,i)=>({index:i+1,width:100,height:140,base64:pngHeader()}))}};}
function fakeProcess(report){
  const child=new EventEmitter();child.stdout=new EventEmitter();child.stderr=new EventEmitter();child.stdin=new EventEmitter();child.kills=0;
  child.stdin.end=bytes=>{child.input=bytes;queueMicrotask(()=>{if(report)report(child);});};
  child.kill=()=>{child.kills++;queueMicrotask(()=>child.emit('close',1));return true;};
  let invocation;return{child,get invocation(){return invocation;},spawn:(executable,args,options)=>{invocation={executable,args,options};return child;}};
}
const emit=(child,output,code=0)=>{const bytes=Buffer.from(JSON.stringify(output));child.stdout.emit('data',bytes.subarray(0,13));child.stdout.emit('data',bytes.subarray(13));child.emit('close',code);};

test('canonical fingerprint projects exact native table delimiters while preserving other whitespace and controls',()=>{
  assert.equal(canonicalBody(' A \r\n B\t\r\n\r\n'),' A \n B\t');
  assert.equal(canonicalBody('cell\r\u0007\r'),'cell\t');
  assert.equal(canonicalBody('a\r\u0007\r\u0007b\r'),'a\t\tb');
  assert.equal(canonicalBody('a\u0007b'),'a\u0007b');
  assert.equal(canonicalBody('a\r\n\u0007b\r'),'a\n\u0007b');
  assert.equal(canonicalBody('a\r\u0007 \t\r'),'a\t \t');
  for(const control of ['\u0001','\u0013','\u0014','\u0015','\f'])assert.notEqual(documentFingerprint('a'+control+'b'),documentFingerprint('ab'));
  assert.equal(documentFingerprint('中文\r\n正文\r'),documentFingerprint('中文\n正文'));
  assert.notEqual(documentFingerprint('中文 '),documentFingerprint('中文'));
});

test('native request rejects arbitrary paths/code/extra fields and enforces exact character/page bounds before spawning',async()=>{
  let spawns=0;for(const body of [null,[],{}, {...request,bodyText:null},{...request,bodyText:'x'.repeat(1000001)},...[-1,0,11,1.5,'1'].map(expectedPageCount=>({...request,expectedPageCount})),{...request,path:'other.docx'},{...request,script:'Quit()'}]){
    assert.equal(validateNativeWordPagesRequest(body).code,'bad_args');
    assert.equal((await renderNativeWordPages(body,{platform:'win32',spawn:()=>{spawns++;}})).code,'bad_args');
  }assert.equal(spawns,0);assert.equal(validateNativeWordPagesRequest({bodyText:'x'.repeat(1000000),expectedPageCount:10}),null);
});

test('non-Windows native renderer honestly reports its requirement without launching anything',async()=>{
  assert.equal((await renderNativeWordPages(request,{platform:'linux',spawn:()=>{throw Error('must not spawn');}})).code,'requirement');
});

test('fixed PS5 STA helper receives UTF-8 JSON through stdin only and returns fingerprint-bound native PNG metadata',async()=>{
  const input={bodyText:'中文 $word.Quit(); \r\n',expectedPageCount:1},process=fakeProcess(child=>emit(child,validOutput(input)));
  const output=await renderNativeWordPages(input,{platform:'win32',spawn:process.spawn});
  assert.equal(output.ok,true);assert.deepEqual(JSON.parse(process.child.input.toString('utf8')),input);
  assert.match(process.invocation.executable,/WindowsPowerShell[\\/]v1\.0[\\/]powershell\.exe$/i);
  assert.ok(process.invocation.args.includes('-STA'));assert.ok(process.invocation.args.includes('-NoProfile'));
  assert.equal(path.basename(process.invocation.args.at(-1)),'native-word-pages.ps1');
  assert.ok(!process.invocation.args.some(arg=>arg.includes(input.bodyText)));assert.equal(process.invocation.options.shell,false);assert.equal(process.child.kills,0);
});

test('wrong native fingerprint, page count, image dimensions or exit status can never produce success',async()=>{
  const corruptions=[out=>out.result.documentFingerprint='0'.repeat(64),out=>out.result.pageCount=2,out=>out.result.pages[0].width=101,out=>out.result.pages[0].base64='not-a-png'];
  for(const corrupt of corruptions){const output=validOutput();corrupt(output);const process=fakeProcess(child=>emit(child,output));assert.equal((await renderNativeWordPages(request,{platform:'win32',spawn:process.spawn})).code,'execution');}
  const process=fakeProcess(child=>emit(child,validOutput(),1));assert.equal((await renderNativeWordPages(request,{platform:'win32',spawn:process.spawn})).code,'execution');
});

test('native not-found/requirement reports remain failures rather than invented page images',async()=>{
  for(const code of ['not_found','requirement']){const process=fakeProcess(child=>emit(child,{ok:false,code,error:'Native requirement unavailable'}));assert.equal((await renderNativeWordPages(request,{platform:'win32',spawn:process.spawn})).code,code);}
});

test('watchdog terminates only the owned helper and releases the native render slot',async()=>{
  const process=fakeProcess();const pending=renderNativeWordPages(request,{platform:'win32',spawn:process.spawn,watchdogMs:25});
  assert.equal((await renderNativeWordPages(request,{platform:'win32',spawn:()=>{throw Error('no second child');}})).code,'busy');
  assert.equal((await pending).code,'timeout');assert.equal(process.child.kills,1);
  const next=fakeProcess(child=>emit(child,validOutput()));assert.equal((await renderNativeWordPages(request,{platform:'win32',spawn:next.spawn})).ok,true);
});

test('bounded native output fails and stops only its helper before accepting excessive transport data',async()=>{
  const process=fakeProcess(child=>child.stdout.emit('data',Buffer.alloc(64*1024*1024+1)));
  assert.equal((await renderNativeWordPages(request,{platform:'win32',spawn:process.spawn})).code,'execution');assert.equal(process.child.kills,1);
});

test('HTTP native route preserves Origin/Host checks, strict JSON type, bounded args and non-public helper files without executing Word',async t=>{
  const bridge=await startBridge();t.after(async()=>{await bridge.stop();cleanupStateDir(bridge.stateDir);});
  const post=(body,headers={})=>bridge.req('/office/native-word-pages',{method:'POST',headers:{'content-type':'application/json',...headers},body:JSON.stringify(body)});
  assert.equal((await post(request,{Origin:'http://evil.example'})).status,403);
  assert.equal((await post({}, {'content-type':'text/plain'})).status,415);
  assert.equal((await post({}, {'content-type':'application/json; charset=latin1'})).status,415);
  assert.equal((await post({}, {'content-type':'application/json; charset="utf-8"'})).status,400);
  assert.equal((await post(null)).status,400);assert.equal((await post({...request,file:'other.docx'})).status,400);
  assert.equal((await post({...request,expectedPageCount:11})).status,400);assert.equal((await post({...request,bodyText:'x'.repeat(1000001)})).status,400);
  assert.equal((await bridge.rawGet('/office/native-word-pages',{Host:'evil.example'})).status,403);
  for(const file of ['native-word-pages.ps1','native-word-pages.cjs'])assert.equal((await bridge.req('/'+file)).status,404);
});

test('PS helper uses documented read-only native page APIs and matches uniquely before any Page access',()=>{
  const source=fs.readFileSync(new URL('../server/native-word-pages.ps1',import.meta.url),'utf8');
  assert.match(source,/GetActiveObject\('Word\.Application'\)/);assert.match(source,/\$page\.EnhMetaFileBits/);
  assert.ok(source.indexOf('$matchingDocumentIndexes.Count -ne 1')<source.indexOf('$pages = $pane.Pages'));
  assert.match(source,/\$view\.Type -ne 3/);
  assert.doesNotMatch(source,/New-Object\s+-ComObject|Invoke-Expression|FinalReleaseComObject|\.(?:Quit|Activate|Select|Repaginate|Open|Close)\s*\(/i);
  assert.doesNotMatch(source,/\$(?:word|matchedDocument|pane|view|page)\.[A-Za-z]+\s*=/);
  assert.match(source,/Word document changed before native page access/);
  assert.match(source,/Word document changed during native page rendering/);
  assert.ok(source.indexOf('Document-Fingerprint $matchedDocument')<source.indexOf('$pages = $pane.Pages'));
  assert.ok(source.lastIndexOf('Document-Fingerprint $matchedDocument')>source.indexOf('$bits = $page.EnhMetaFileBits'));
});

test('Windows PowerShell 5 parses the helper without executing or attaching to Office',{skip:process.platform!=='win32'},()=>{
  const helper=fileURLToPath(new URL('../server/native-word-pages.ps1',import.meta.url)).replace(/'/g,"''");
  const command="$tokens=$null;$parseErrors=$null;[void][System.Management.Automation.Language.Parser]::ParseFile('"+helper+"',[ref]$tokens,[ref]$parseErrors);if($parseErrors.Count){$parseErrors | Format-List;exit 1};[Console]::Write('PS5 AST OK')";
  const executable=path.win32.join(process.env.SystemRoot||'C:\\Windows','System32','WindowsPowerShell','v1.0','powershell.exe');
  const output=spawnSync(executable,['-NoProfile','-NonInteractive','-Command',command],{encoding:'utf8',windowsHide:true,timeout:10000});
  assert.equal(output.status,0,output.stdout+' '+output.stderr);assert.match(output.stdout,/PS5 AST OK/);
});
