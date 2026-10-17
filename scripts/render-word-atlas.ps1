param([switch]$Worker,[string]$RunDirectory,[ValidateRange(5,120)][int]$DeadlineSeconds=30)
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$source=Join-Path $repo 'docs/style-atlas/word-base-styles.docx'
$published=Join-Path $repo 'docs/style-atlas/word-base-styles'
if($Worker){
 Add-Type -AssemblyName System.Drawing
 Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class AtlasWordPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid); }'
 $before=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
 $app=New-Object -ComObject Word.Application
 $app.Visible=$true
 $created=@(Get-Process WINWORD | Where-Object {$_.Id -notin $before})
 if($created.Count -ne 1){throw 'Dedicated Word process is ambiguous'}
 $hwnd=$created[0].MainWindowHandle
 [uint32]$ownedWordId=0
 $null=[AtlasWordPid]::GetWindowThreadProcessId([IntPtr]$hwnd,[ref]$ownedWordId)
 $beforeCount=$app.Documents.Count
 if($hwnd -eq 0 -or $ownedWordId -in $before -or $beforeCount -ne 0){throw 'Dedicated Word isolation failed'}
 @{pid=$ownedWordId;hwnd=$hwnd;startTime=(Get-Process -Id $ownedWordId).StartTime.ToString('o');beforeDocuments=$beforeCount}|ConvertTo-Json|Set-Content -Encoding UTF8 -LiteralPath (Join-Path $RunDirectory 'owned-word.json')
 $doc=$null
 try{
  $copy=Join-Path $RunDirectory 'input.docx'
  $doc=$app.Documents.Open($copy,$false,$true,$false)
  if($doc.FullName -ne $copy){throw 'Unexpected document target'}
  $doc.ActiveWindow.View.Type=3
  $pages=$doc.ComputeStatistics(2)
  $footerFields=0
  foreach($section in $doc.Sections){foreach($footer in $section.Footers){foreach($field in $footer.Range.Fields){if($field.Type -eq 33){$footerFields++}}}}
  $counts=@{toc=$doc.TablesOfContents.Count;tables=$doc.Tables.Count;images=$doc.InlineShapes.Count;PAGE=$footerFields}
  if($pages -ne 5 -or $counts.toc -ne 1 -or $counts.tables -ne 1 -or $counts.images -ne 1 -or $counts.PAGE -ne 1){throw 'Atlas page or native object counts differ from expected 5/1/1/1/1'}
  $files=@()
  for($i=1;$i -le $pages;$i++){
   $start=$doc.GoTo(1,1,$i).Start
   $end=if($i -lt $pages){$doc.GoTo(1,1,($i+1)).Start}else{$doc.Content.End}
   [byte[]]$bits=$doc.Range($start,$end).EnhMetaFileBits
   $stream=[IO.MemoryStream]::new($bits,$false)
   $metafile=[Drawing.Imaging.Metafile]::new($stream)
   $height=[int][Math]::Round(1240*$metafile.Height/$metafile.Width)
   $bitmap=[Drawing.Bitmap]::new(1240,$height)
   $graphics=[Drawing.Graphics]::FromImage($bitmap)
   try{
    $graphics.Clear([Drawing.Color]::White)
    $graphics.DrawImage($metafile,[Drawing.Rectangle]::new(0,0,1240,$height))
   $name='page-{0:D2}.png' -f $i
    $bitmap.Save((Join-Path $RunDirectory $name),[Drawing.Imaging.ImageFormat]::Png)
    $files+=$name
   }finally{$graphics.Dispose();$bitmap.Dispose();$metafile.Dispose();$stream.Dispose()}
  }
 @{id='word-base-styles';host='Word';pages=$pages;nativePageCount=$pages;counts=$counts;renderScope='body-range';renderScopeNote='Native Word page body ranges only; excludes headers, footers and page margins.';renderer='Microsoft Word Range.EnhMetaFileBits, rasterized preserving aspect ratio';previewFiles=$files;pdfAvailable=$false;pdfFailure='Native ExportAsFixedFormat exceeded30s on atlas variants and one-paragraph ASCII control (2026-10-14). This renderer does not attempt PDF export.';sourceLiveTOC=$true;sourceLivePAGE=$true;officeJsE2E=$false;dedicatedInstance=$true;openedReadOnly=$true;ownedWordPID=$ownedWordId;beforeDocuments=$beforeCount;sourceSHA256=(Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash}|ConvertTo-Json -Depth 8|Set-Content -Encoding UTF8 -LiteralPath (Join-Path $RunDirectory 'evidence.json')
 }finally{if($doc){$doc.Close(0)}}
exit 0
}
$run=Join-Path $repo ('.tmp/word-native-render/'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $run -Force
$hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
Copy-Item -LiteralPath $source -Destination (Join-Path $run 'input.docx')
$workerProcess=$null
try{
 $exe=(Get-Process -Id $PID).Path
$args=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$PSCommandPath+'"'),'-Worker','-RunDirectory',('"'+$run+'"'))
 $workerProcess=Start-Process -FilePath $exe -ArgumentList $args -PassThru -RedirectStandardOutput (Join-Path $run 'stdout.log') -RedirectStandardError (Join-Path $run 'stderr.log')
if(-not $workerProcess.WaitForExit($DeadlineSeconds*1000)){throw ('Native Word render exceeded '+$DeadlineSeconds+' seconds; diagnostics: '+$run)}
 if($null -ne $workerProcess.ExitCode -and $workerProcess.ExitCode -ne 0){throw ('Native Word worker failed; diagnostics: '+$run)}
$evidencePath=Join-Path $run 'evidence.json'
 $evidence=Get-Content -Raw -LiteralPath $evidencePath|ConvertFrom-Json
if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $hash -or $evidence.sourceSHA256 -ne $hash){throw 'Source changed during read-only rendering'}
 $null=New-Item -ItemType Directory -Path $published -Force
foreach($name in $evidence.previewFiles){Copy-Item -LiteralPath (Join-Path $run $name) -Destination (Join-Path $published $name) -Force}
 $evidence|Add-Member -NotePropertyName sourceUnchanged -NotePropertyValue $true
$evidence|ConvertTo-Json -Depth 8|Set-Content -Encoding UTF8 -LiteralPath (Join-Path $repo 'docs/style-atlas/word-base-styles.evidence.json')
 Write-Output ('Published '+$evidence.pages+' native Word body previews; PDF not attempted; source unchanged.')
}
finally{
 $identityPath=Join-Path $run 'owned-word.json'
if(Test-Path -LiteralPath $identityPath){
  $identity=Get-Content -Raw -LiteralPath $identityPath|ConvertFrom-Json
 $owned=Get-Process -Id $identity.pid -ErrorAction SilentlyContinue
  if($owned -and $owned.ProcessName -eq 'WINWORD' -and $owned.StartTime.ToString('o') -eq $identity.startTime -and $identity.beforeDocuments -eq 0){Stop-Process -Id $identity.pid -Force}
}
 if($workerProcess -and -not $workerProcess.HasExited){Stop-Process -Id $workerProcess.Id -Force}
}
