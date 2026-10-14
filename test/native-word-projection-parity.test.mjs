import { before, test } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
const require = createRequire(import.meta.url);
const { canonicalBody, documentFingerprint } = require('../server/native-word-pages.cjs');
const windows = process.platform === 'win32';
const base = 37;
const fixture = (name, raw, intervals, projected, canonical, failure = false) => ({name,raw,start:base,intervals:intervals.map(([start,end])=>({start:base+start,end:base+end})),projected,canonical,failure});
const slashText = 'literal/left\r/\r右/正文 \t\r';
const slashIndex = slashText.indexOf('\r/\r') + 1;
const tableText = 'A\r\u0007B\r\u0007\r\u0007\r/\r正文/\r';
const tableImageIndex = tableText.indexOf('\r/\r') + 1;
const controls = 'x\u0007 \u0001\u0013code\u0014res\u0015\f / \t\r\n\u0007end\r\r';
const unicode = '😀/文 / 尾\r';
const unicodeImageIndex = unicode.indexOf(' / ') + 1;
const fixtures = [
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
$allowed=@('Release-OwnedCom','Canonical-Body','Body-Fingerprint','Projected-DocumentBody','Fail-Native')
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
  $doc=[pscustomobject]@{Content=[pscustomobject]@{Start=[int]$Fixture.start;End=([int]$Fixture.start+$Fixture.raw.Length);Text=[string]$Fixture.raw};InlineShapes=$shapes;RawText=[string]$Fixture.raw;Offset=[int]$Fixture.start;Calls=(New-Object 'System.Collections.Generic.List[object]')}
  $doc|Add-Member -MemberType ScriptMethod -Name Range -Value {
    param([int]$Start,[int]$End)
    if($Start -lt $this.Offset -or $End -lt $Start -or $End -gt ($this.Offset+$this.RawText.Length)){throw 'Fake document range out of bounds'}
    $this.Calls.Add(@{start=$Start;end=$End})
    return [pscustomobject]@{Start=$Start;End=$End;Text=$this.RawText.Substring($Start-$this.Offset,$End-$Start)}
  }
  return $doc
}
$fixtures=[Console]::In.ReadToEnd()|ConvertFrom-Json
$receipts=New-Object 'System.Collections.Generic.List[object]'
foreach($fixture in $fixtures){
  $doc=New-FakeDocument $fixture
  try {
    $projected=Projected-DocumentBody $doc
    $receipts.Add(@{name=$fixture.name;ok=$true;projected=$projected;canonical=(Canonical-Body $projected);fingerprint=(Body-Fingerprint $projected);calls=$doc.Calls.ToArray()})
  } catch {
    $receipts.Add(@{name=$fixture.name;ok=$false;code=$_.Exception.Data['nativeCode'];error=$_.Exception.Message;calls=$doc.Calls.ToArray()})
  }
}
[Console]::Write((ConvertTo-Json -InputObject $receipts.ToArray() -Compress -Depth 8))
`;
  const executable=path.win32.join(process.env.SystemRoot||'C:\\Windows','System32','WindowsPowerShell','v1.0','powershell.exe');
  const output=spawnSync(executable,['-NoProfile','-NonInteractive','-EncodedCommand',Buffer.from(script,'utf16le').toString('base64')],{input:JSON.stringify(fixtures),encoding:'utf8',windowsHide:true,timeout:15000,maxBuffer:1024*1024});
  assert.equal(output.status,0,output.error?.message||output.stderr||output.stdout);
  receipts=JSON.parse(output.stdout.replace(/^\uFEFF/,''));assert.equal(receipts.length,fixtures.length);
});
for(const scenario of fixtures){
  test('real PowerShell projection: '+scenario.name,{skip:!windows},()=>{
    const receipt=receipts.find(item=>item.name===scenario.name);assert.ok(receipt,'Every fixture must have a PowerShell receipt');
    if(scenario.failure){assert.equal(receipt.ok,false,'Invalid ranges must never produce a matching body');assert.equal(receipt.code,'requirement',receipt.error);return;}
    assert.equal(receipt.ok,true,receipt.error);
    assert.equal(receipt.projected,scenario.projected,'Projection removes only declared object intervals');
    assert.equal(receipt.canonical,scenario.canonical,'PowerShell preserves every other character');
    assert.equal(canonicalBody(scenario.projected),scenario.canonical,'JavaScript and PowerShell canonicalization agree');
    assert.equal(receipt.fingerprint,documentFingerprint(scenario.projected),'Both languages hash identical UTF-8 text');
    assert.equal(receipt.calls.map(call=>scenario.raw.slice(call.start-base,call.end-base)).join(''),scenario.projected,'All retained text comes from bounded ranges');
  });
}
