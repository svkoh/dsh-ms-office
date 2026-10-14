param(
  [ValidateSet('word-business-blue','word-plain-accessible','excel-budget-green','excel-plain-data','ppt-clean-blue-deck','ppt-modern-accent')]
  [string[]]$Preset=@('word-business-blue','word-plain-accessible','excel-budget-green','excel-plain-data','ppt-clean-blue-deck','ppt-modern-accent'),
  [ValidateRange(800,2400)][int]$Width=1600
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$storageType=[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$pdfType=[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]
$memoryType=[Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]
$readerType=[Windows.Storage.Streams.DataReader,Windows.Storage.Streams,ContentType=WindowsRuntime]
$optionsType=[Windows.Data.Pdf.PdfPageRenderOptions,Windows.Data.Pdf,ContentType=WindowsRuntime]
$taskMethods=[System.WindowsRuntimeSystemExtensions].GetMethods()
$genericAwait=$taskMethods|Where-Object {$_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'}|Select-Object -First 1
$actionAwait=$taskMethods|Where-Object {$_.Name -eq 'AsTask' -and -not $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'}|Select-Object -First 1
if(-not $genericAwait -or -not $actionAwait){throw 'Windows Runtime async bridge unavailable; use Windows PowerShell 5.1.'}
function Wait-Result($Operation,[Type]$ResultType){
  $task=$genericAwait.MakeGenericMethod($ResultType).Invoke($null,@($Operation))
  $task.Wait()
  return $task.Result
}
function Wait-Action($Operation){
  $task=$actionAwait.Invoke($null,@($Operation))
  $task.Wait()
}
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../../docs/examples'))
$report=@()
foreach($id in $Preset){
  $base=if($id -match '^(word|excel)-'){'native-'+$id}else{$id}
  $pdfFile=Join-Path $root ($base+'.pdf')
  if(-not (Test-Path -LiteralPath $pdfFile -PathType Leaf)){throw ('Generated PDF missing: '+$pdfFile)}
  $file=Wait-Result ($storageType::GetFileFromPathAsync($pdfFile)) $storageType
  $pdf=Wait-Result ($pdfType::LoadFromFileAsync($file)) $pdfType
  $directory=Join-Path $root $id
  [IO.Directory]::CreateDirectory($directory)|Out-Null
  $images=@()
  for($i=0;$i -lt $pdf.PageCount;$i++){
    $page=$pdf.GetPage([uint32]$i)
    $stream=[Activator]::CreateInstance($memoryType)
    $reader=$null
    try{
      $options=[Activator]::CreateInstance($optionsType)
      $options.DestinationWidth=[uint32]$Width
      Wait-Action ($page.RenderToStreamAsync($stream,$options))
      $reader=[Activator]::CreateInstance($readerType,@($stream.GetInputStreamAt(0)))
      $loaded=Wait-Result ($reader.LoadAsync([uint32]$stream.Size)) ([uint32])
      if($loaded -ne $stream.Size){throw 'Incomplete native PDF image stream'}
      $bytes=New-Object byte[] ([int]$stream.Size)
      $reader.ReadBytes($bytes)
      $png=Join-Path $directory ('pdf-page-{0:d2}.png' -f ($i+1))
      [IO.File]::WriteAllBytes($png,$bytes)
      $actualWidth=([int]$bytes[16] -shl 24) -bor ([int]$bytes[17] -shl 16) -bor ([int]$bytes[18] -shl 8) -bor [int]$bytes[19]
      $actualHeight=([int]$bytes[20] -shl 24) -bor ([int]$bytes[21] -shl 16) -bor ([int]$bytes[22] -shl 8) -bor [int]$bytes[23]
      $images+=@{page=$i+1;path=$png;width=$actualWidth;height=$actualHeight;requestedWidth=$Width;bytes=$bytes.Length}
    }finally{
      if($reader){$reader.Dispose()}
      $stream.Dispose()
      $page.Dispose()
    }
  }
  $report+=@{presetId=$id;pdf=$pdfFile;pageCount=$pdf.PageCount;renderer='Windows.Data.Pdf';images=$images}
  Write-Output ('Rendered '+$id+': '+$pdf.PageCount+' actual PDF pages')
}
$json=@{scope='native-pdf-rendering';renderer='Windows.Data.Pdf';files=$report}|ConvertTo-Json -Depth 10
[IO.File]::WriteAllText((Join-Path $root 'pdf-render-evidence.json'),$json,[Text.UTF8Encoding]::new($false))
