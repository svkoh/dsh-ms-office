import { before, test } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
const require = createRequire(import.meta.url);
const { canonicalBody, documentFingerprint } = require('../server/native-word-pages.cjs');
const windows = process.platform === 'win32';
const base = 37;
const fixture = (name, raw, intervals, projected, canonical, failure = false, rows = []) => ({name,raw,start:base,intervals:intervals.map(([start,end])=>({start:base+start,end:base+end})),projected,canonical,failure,rows:rows.map(([start,end])=>({start:base+start,end:base+end}))});
const slashText = 'literal/left\r/\r右/正文 \t\r';
const slashIndex = slashText.indexOf('\r/\r') + 1;
const tableText = 'A\r\u0007B\r\u0007\r\u0007\r/\r正文/\r';
const tableImageIndex = tableText.indexOf('\r/\r') + 1;
const controls = 'x\u0007 \u0001\u0013code\u0014res\u0015\f / \t\r\n\u0007end\r\r';
const unicode = '😀/文 / 尾\r';
const unicodeImageIndex = unicode.indexOf(' / ') + 1;
// Synthetic count-matched regression, not a capture from live Office.
const observedRows = Array.from({length:43},(_,i)=>Array.from({length:i<17?4:3},()=> 'cell'));
let observedRaw = 'paragraph\r'.repeat(53), observedProjected = observedRaw;
const observedRanges = [];
for(const cells of observedRows){
  const start=observedRaw.length;
  observedRaw+=cells.join('\r\u0007')+'\r\u0007\r\u0007';
  observedRanges.push([start,observedRaw.length]);
  observedProjected+=cells.join('\r\u0007')+'\r';
}
observedRaw+='tail\r';observedProjected+='tail\r';
const fixtures = [
  {...fixture('flattened Office.js row tails preserve exact cells, images and literal slashes',tableText,[[tableImageIndex,tableImageIndex+1]],'A\r\u0007B\t\t\r\r正文/\r','A\tB\t\t\n\n正文/',false,[[0,8]]),rowTail:'\t\t'},
  {...fixture('native character positions can differ from Range.Text UTF-16 offsets','A\r\u0007\r\u0007tail\r',[],'A\rtail\r','A\ntail',false,[[0,3]]),nativeOffsets:[0,1,3,5,6,7,8,9,10]},
  fixture('native row tails produce exactly 103 tabs and 96 paragraphs from the reported counts',observedRaw,[],observedProjected,'paragraph\n'.repeat(53)+observedRows.map(cells=>cells.join('\t')+'\n').join('')+'tail',false,observedRanges),
  {...fixture('reject nested table projection','A\r\u0007\r\u0007tail',[],null,null,true,[[0,5]]),nestingLevel:2},
  {...fixture('bound native table count','A\r\u0007\r\u0007tail',[],null,null,true,[[0,5]]),tableCount:1001},
  {...fixture('bound native row count','A\r\u0007\r\u0007tail',[],null,null,true,[[0,5]]),rowCount:10001},
  fixture('empty cells retain tabs while a bound row tail becomes a paragraph','A\r\u0007\r\u0007B\r\u0007\r\u0007tail\r',[],'A\r\u0007\r\u0007B\rtail\r','A\t\tB\ntail',false,[[0,10]]),
  fixture('adjacent rows retain distinct row breaks','A\r\u0007\r\u0007B\r\u0007\r\u0007tail\r',[],'A\rB\rtail\r','A\nB\ntail',false,[[0,5],[5,10]]),
  fixture('inline objects inside cells are removed before exact matching','A/\r\u0007B\r\u0007\r\u0007tail\r',[[1,2]],'A\r\u0007B\rtail\r','A\tB\ntail',false,[[0,9]]),
  fixture('reject row ranges outside the body','body',[],null,null,true,[[-1,4]]),
  fixture('reject unsupported row delimiter bytes','text\r\u0007tail',[],null,null,true,[[0,6]]),
  fixture('reject objects overlapping structural row tails','A\r\u0007\r\u0007tail',[[2,3]],null,null,true,[[0,5]]),
  fixture('reject duplicate row tails','A\r\u0007\r\u0007tail',[],null,null,true,[[0,5],[0,5]]),
  fixture('exclude only positioned picture slash; retain literal slashes and whitespace',slashText,[[slashIndex,slashIndex+1]],slashText.slice(0,slashIndex)+slashText.slice(slashIndex+1),'literal/left\n\n右/正文 \t'),
  fixture('map cell pairs while preserving double-tab row ends and paragraphs',tableText,[[tableImageIndex,tableImageIndex+1]],tableText.slice(0,tableImageIndex)+tableText.slice(tableImageIndex+1),'A\tB\t\t\n\n正文/'),
  fixture('retain BEL, object/field controls, FF, whitespace and CRLF-adjacent BEL',controls,[],controls,'x\u0007 \u0001\u0013code\u0014res\u0015\f / \t\n\u0007end'),
  fixture('no objects preserves all slash text','no image / slash\r',[],'no image / slash\r','no image / slash'),
  fixture('sort nonoverlapping ranges','A/B/C',[[3,4],[1,2]],'ABC','ABC'),
  fixture('allow boundary objects and zero-length retained segments','/body/',[[0,1],[5,6]],'body','body'),
  fixture('allow adjacent distinct objects','//tail',[[0,1],[1,2]],'tail','tail'),
  fixture('use native UTF-16 offsets with non-BMP text',unicode,[[unicodeImageIndex,unicodeImageIndex+1]],unicode.slice(0,unicodeImageIndex)+unicode.slice(unicodeImageIndex+1),'😀/文  尾'),
  fixture('reject overlapping object ranges','A//B',[[1,3],[2,3]],null,null,true),
  fixture('reject range before body','body',[[-1,0]],null,null,true),
  fixture('reject range after body','body',[[3,5]],null,null,true),
  fixture('reject empty object range','body',[[1,1]],null,null,true),
  fixture('reject reversed object range','body',[[2,1]],null,null,true),
  fixture('bound object count','body',Array.from({length:1001},()=>[0,1]),null,null,true),
];
test('synthetic native table fixture accounts for the reported delimiter counts without fuzzy matching',()=>{
  assert.equal(observedRaw.split('\r\u0007').length-1,189);
  assert.equal((observedRaw.match(/\r/g)||[]).length,243);
  const expected=fixtures.find(item=>item.raw===observedRaw).canonical;
  assert.equal((expected.match(/\t/g)||[]).length,103);
  assert.equal((expected.match(/\n/g)||[]).length,96);
  assert.notEqual(documentFingerprint(observedRaw),documentFingerprint(expected),'Flat table delimiter replacement cannot match this structural body');
  assert.equal(documentFingerprint(observedProjected),documentFingerprint(expected));
});
let receipts;
before(() => {
  if (!windows) return;
  const helper=fileURLToPath(new URL('../server/native-word-pages.ps1',import.meta.url)).replaceAll("'","''");
  // Only allowlisted function definitions execute: never the helper main block or COM attachment.
  const script=String.raw`
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::InputEncoding=New-Object Text.UTF8Encoding($false)
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$tokens=$null; $parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile('${helper}',[ref]$tokens,[ref]$parseErrors)
if($parseErrors.Count){throw 'Native helper has PowerShell syntax errors'}
$allowed=@('Release-OwnedCom','Canonical-Body','Body-Fingerprint','Body-Diagnostics','Native-RowTailRanges','Projected-DocumentBody','Document-Fingerprint','Fail-Native')
$definitions=@($ast.EndBlock.Statements|Where-Object{$_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -in $allowed})
if($definitions.Count -ne $allowed.Count){throw 'Missing or duplicate allowlisted projection functions'}
foreach($definition in $definitions){
  if($definition.Extent.Text -match 'GetActiveObject|FinalReleaseComObject|New-Object\s+-ComObject'){throw 'COM attachment forbidden in pure projection functions'}
  . ([scriptblock]::Create($definition.Extent.Text))
}
function New-FakeDocument([object]$Fixture){
  $items=@($Fixture.intervals|ForEach-Object{[pscustomobject]@{Type=3;Range=[pscustomobject]@{Start=[int]$_.start;End=[int]$_.end;Text='/'}}})
  $shapes=[pscustomobject]@{Count=$items.Count;Items=$items}
  $shapes|Add-Member -MemberType ScriptMethod -Name Item -Value {param([int]$Index) return $this.Items[$Index-1]}
  $offsets=@(0..$Fixture.raw.Length);$markerWidth=2
  if($Fixture.nativeOffsets){$offsets=@($Fixture.nativeOffsets);$markerWidth=1}
  $rowItems=@($Fixture.rows|ForEach-Object{
    $rowEnd=[int]$_.end; $markerItems=@()
    foreach($markerStart in @(($rowEnd-2*$markerWidth),($rowEnd-$markerWidth))){
      $localStart=$offsets[$markerStart-[int]$Fixture.start];$localEnd=$offsets[$markerStart+$markerWidth-[int]$Fixture.start]
      $markerItems+=[pscustomobject]@{Start=$markerStart;End=($markerStart+$markerWidth);Text=$Fixture.raw.Substring($localStart,$localEnd-$localStart)}
    }
    $characters=[pscustomobject]@{Count=$markerItems.Count;Items=$markerItems}
    $characters|Add-Member -MemberType ScriptMethod -Name Item -Value {param([int]$Index) return $this.Items[$Index-1]}
    [pscustomobject]@{Range=[pscustomobject]@{Start=[int]$_.start;End=$rowEnd;Characters=$characters}}
  })
  $rows=[pscustomobject]@{Count=$(if($Fixture.rowCount){[int]$Fixture.rowCount}else{$rowItems.Count});Items=$rowItems}
  $rows|Add-Member -MemberType ScriptMethod -Name Item -Value {param([int]$Index) return $this.Items[$Index-1]}
  $tableItems=@();if($rowItems.Count){$tableItems=@([pscustomobject]@{NestingLevel=$(if($Fixture.nestingLevel){[int]$Fixture.nestingLevel}else{1});Rows=$rows})}
  $tables=[pscustomobject]@{Count=$(if($Fixture.tableCount){[int]$Fixture.tableCount}else{$tableItems.Count});Items=$tableItems}
  $tables|Add-Member -MemberType ScriptMethod -Name Item -Value {param([int]$Index) return $this.Items[$Index-1]}
  $doc=[pscustomobject]@{Content=[pscustomobject]@{Start=[int]$Fixture.start;End=([int]$Fixture.start+$offsets.Count-1);Text=[string]$Fixture.raw};InlineShapes=$shapes;Tables=$tables;RawText=[string]$Fixture.raw;TextOffsets=$offsets;Offset=[int]$Fixture.start;Calls=(New-Object 'System.Collections.Generic.List[object]')}
  $doc|Add-Member -MemberType ScriptMethod -Name Range -Value {
    param([int]$Start,[int]$End)
    if($Start -lt $this.Offset -or $End -lt $Start -or $End -gt ($this.Offset+$this.TextOffsets.Count-1)){throw 'Fake document range out of bounds'}
    $this.Calls.Add(@{start=$Start;end=$End})
    $textStart=$this.TextOffsets[$Start-$this.Offset];$textEnd=$this.TextOffsets[$End-$this.Offset]
    return [pscustomobject]@{Start=$Start;End=$End;Text=$this.RawText.Substring($textStart,$textEnd-$textStart)}
  }
  return $doc
}
$fixtures=[Console]::In.ReadToEnd()|ConvertFrom-Json
$receipts=New-Object 'System.Collections.Generic.List[object]'
foreach($fixture in $fixtures){
  $doc=New-FakeDocument $fixture
  try {
    $projected=if($fixture.rowTail){Projected-DocumentBody $doc $fixture.rowTail}else{Projected-DocumentBody $doc}
    $receipts.Add(@{name=$fixture.name;ok=$true;projected=$projected;canonical=(Canonical-Body $projected);fingerprint=(Body-Fingerprint $projected);matchedFingerprint=(Document-Fingerprint (New-FakeDocument $fixture) (Body-Fingerprint $fixture.canonical));unmatchedFingerprint=(Document-Fingerprint (New-FakeDocument $fixture) ('0'*64));diagnostics=(Body-Diagnostics $projected $fixture.canonical);calls=$doc.Calls.ToArray()})
  } catch {
    $receipts.Add(@{name=$fixture.name;ok=$false;code=$_.Exception.Data['nativeCode'];error=$_.Exception.Message;calls=$doc.Calls.ToArray()})
  }
}
[Console]::Write((ConvertTo-Json -InputObject $receipts.ToArray() -Compress -Depth 8))
`;
  const executable=path.win32.join(process.env.SystemRoot||'C:\\Windows','System32','WindowsPowerShell','v1.0','powershell.exe');
  // File-backed stdio avoids named-pipe capture on confined Windows runners.
  const temp=fs.mkdtempSync(path.join(os.tmpdir(),'native-word-projection-'));
  const inputPath=path.join(temp,'input.json'),outputPath=path.join(temp,'output.json'),errorPath=path.join(temp,'error.txt');
  const descriptors=[];
  try{
    fs.writeFileSync(inputPath,JSON.stringify(fixtures));
    descriptors.push(fs.openSync(inputPath,'r'),fs.openSync(outputPath,'w'),fs.openSync(errorPath,'w'));
    const output=spawnSync(executable,['-NoProfile','-NonInteractive','-EncodedCommand',Buffer.from(script,'utf16le').toString('base64')],{stdio:descriptors,windowsHide:true,timeout:15000});
    const stdout=fs.readFileSync(outputPath,'utf8'),stderr=fs.readFileSync(errorPath,'utf8');
    assert.equal(output.status,0,output.error?.message||stderr||stdout);
    receipts=JSON.parse(stdout.replace(/^\uFEFF/,''));assert.equal(receipts.length,fixtures.length);
  }finally{
    for(const descriptor of descriptors)fs.closeSync(descriptor);
    fs.rmSync(temp,{recursive:true,force:true});
  }
});
for(const scenario of fixtures){
  test('real PowerShell projection: '+scenario.name,{skip:!windows},()=>{
    const receipt=receipts.find(item=>item.name===scenario.name);assert.ok(receipt,'Every fixture must have a PowerShell receipt');
    if(scenario.failure){assert.equal(receipt.ok,false,'Invalid ranges must never produce a matching body');assert.equal(receipt.code,'requirement',receipt.error);return;}
    assert.equal(receipt.ok,true,receipt.error);
    assert.equal(receipt.projected,scenario.projected,'Projection removes only declared object intervals');
    assert.equal(receipt.canonical,scenario.canonical,'PowerShell preserves every other character');
    assert.equal(canonicalBody(scenario.projected),scenario.canonical,'JavaScript and PowerShell canonicalization agree');
    assert.equal(receipt.diagnostics.firstMismatch,scenario.canonical.length,'Projected diagnostics compare the text actually used for matching');
    assert.equal(receipt.diagnostics.length,scenario.canonical.length);
    assert.equal(receipt.fingerprint,documentFingerprint(scenario.projected),'Both languages hash identical UTF-8 text');
    assert.equal(receipt.matchedFingerprint,documentFingerprint(scenario.projected),'Matching accepts only an exact expected fingerprint');
    assert.notEqual(receipt.unmatchedFingerprint,'0'.repeat(64),'Unrelated expected fingerprints must not match');
    const offsets=scenario.nativeOffsets||Array.from({length:scenario.raw.length+1},(_,i)=>i);
    if(!scenario.rows.length)assert.equal(receipt.calls.map(call=>scenario.raw.slice(offsets[call.start-base],offsets[call.end-base])).join(''),scenario.projected,'All retained text comes from bounded native ranges');
    for(const call of receipt.calls)assert.ok(call.start>=base&&call.end>=call.start&&call.end<=base+offsets.length-1,'Every text probe remains within the body');
  });
}
