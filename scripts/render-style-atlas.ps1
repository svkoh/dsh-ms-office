param(
 [ValidateSet('word-base-styles','excel-base-styles','ppt-base-styles')]
 [string[]]$Preset=@('word-base-styles','excel-base-styles','ppt-base-styles'),
 [ValidateRange(800,2400)][int]$Width=1600
)
$ErrorActionPreference='Stop'
# PyMuPDF is workspace-local; install with:
# python -m pip install --target .tmp/office-libs PyMuPDF
# Word has no native PDF in this runtime; keep its genuine body previews unchanged.
$pdfPresets=@($Preset | Where-Object {$_ -ne 'word-base-styles'})
if('word-base-styles' -in $Preset){Write-Output 'Word: existing native body-range previews retained; PDF unavailable.'}
if($pdfPresets.Count -gt 0){
 $renderer=Join-Path $PSScriptRoot 'render-atlas-pdfs.py'
 & python $renderer --width $Width --preset $pdfPresets
 if($LASTEXITCODE -ne 0){throw ('Native PDF rasterizer failed with exit code '+$LASTEXITCODE)}
}
