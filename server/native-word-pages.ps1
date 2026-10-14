# Read-only Windows Word page renderer. Input/output are UTF-8 JSON over stdio.
# References: learn.microsoft.com/en-us/office/vba/api/word.page.enhmetafilebits
# and learn.microsoft.com/en-us/office/vba/api/word.pane.pages
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$ownedReferences = New-Object 'System.Collections.Generic.List[object]'
$output = $null
function Release-OwnedCom([object] $Reference) {
  if ($null -ne $Reference -and [Runtime.InteropServices.Marshal]::IsComObject($Reference)) {
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Reference) } catch { }
  }
}
function Canonical-Body([string] $Text) {
  $normalized = $Text.Replace(([string][char]13 + [char]7), [string][char]9).Replace(([string][char]13 + [char]10), [string][char]10).Replace([string][char]13, [string][char]10)
  return $normalized.TrimEnd([char[]]@([char]10))
}
function Body-Fingerprint([string] $Text) {
  $sha = [Security.Cryptography.SHA256]::Create()
  try {
    $bytes = [Text.Encoding]::UTF8.GetBytes((Canonical-Body $Text))
    return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
  } finally { $sha.Dispose() }
}
function Body-Diagnostics([string] $Text, [string] $Expected) {
  $actual = Canonical-Body $Text
  $wanted = Canonical-Body $Expected
  $index = 0
  while ($index -lt [Math]::Min($actual.Length, $wanted.Length) -and $actual[$index] -ceq $wanted[$index]) { $index++ }
  $tableProjection=$Text.Replace(([string][char]13+[char]7),[string][char]9)
  return @{ objectMarkers=([regex]::Matches($Text,[string][char]1)).Count; fieldBegins=([regex]::Matches($Text,[string][char]19)).Count; tableProjectedHash=(Body-Fingerprint $tableProjection); tableWithoutObjectHash=(Body-Fingerprint ($tableProjection.Replace([string][char]1,''))); length=$actual.Length; hash=(Body-Fingerprint $Text); tableCellEnds=([regex]::Matches($Text, ([string][char]13+[char]7))).Count; tabs=([regex]::Matches($Text,[string][char]9)).Count; paragraphMarks=([regex]::Matches($Text,[string][char]13)).Count; pageBreaks=([regex]::Matches($Text,[string][char]12)).Count; firstMismatch=$index; expectedCode=$(if($index -lt $wanted.Length){[int]$wanted[$index]}else{-1}); actualCode=$(if($index -lt $actual.Length){[int]$actual[$index]}else{-1}) }
}
function Projected-DocumentBody([object] $Document) {
  $content=$null; $shapes=$null
  try {
    $content=$Document.Content; $start=[int]$content.Start; $end=[int]$content.End
    $shapes=$Document.InlineShapes
    if([int]$shapes.Count -gt 1000) { Fail-Native 'requirement' 'Too many native inline object ranges for bounded matching' }
    $intervals=New-Object 'System.Collections.Generic.List[object]'
    for($i=1;$i -le [int]$shapes.Count;$i++) {
      $shape=$null; $range=$null
      try { $shape=$shapes.Item($i); $range=$shape.Range; $intervals.Add(@{start=[int]$range.Start;end=[int]$range.End}) } finally { Release-OwnedCom $range; Release-OwnedCom $shape }
    }
    $builder=New-Object Text.StringBuilder; $cursor=$start
    foreach($interval in @($intervals | Sort-Object { $_.start })) {
      if($interval.start -lt $cursor -or $interval.end -le $interval.start -or $interval.end -gt $end) { Fail-Native 'requirement' 'Native inline object ranges overlap or leave the body scope' }
      $segment=$null
      try { $segment=$Document.Range($cursor,$interval.start); [void]$builder.Append([string]$segment.Text) } finally { Release-OwnedCom $segment }
      $cursor=$interval.end
    }
    $last=$null
    try { $last=$Document.Range($cursor,$end); [void]$builder.Append([string]$last.Text) } finally { Release-OwnedCom $last }
    return $builder.ToString()
  } finally { Release-OwnedCom $shapes; Release-OwnedCom $content }
}
function Document-Fingerprint([object] $Document) {
  $content = $null
  return Body-Fingerprint (Projected-DocumentBody $Document)
}
function Fail-Native([string] $Code, [string] $Message) {
  $failure = New-Object System.InvalidOperationException($Message)
  $failure.Data['nativeCode'] = $Code
  throw $failure
}
try {
  $request = [Console]::In.ReadToEnd() | ConvertFrom-Json
  $names = @($request.PSObject.Properties.Name)
  if ($names.Count -ne 2 -or @($names | Where-Object { $_ -notin @('bodyText', 'expectedPageCount') }).Count -gt 0 -or $request.bodyText -isnot [string] -or $request.bodyText.Length -gt 1000000 -or ($request.expectedPageCount -isnot [int] -and $request.expectedPageCount -isnot [long]) -or $request.expectedPageCount -lt 1 -or $request.expectedPageCount -gt 10) {
    Fail-Native 'requirement' 'Invalid fixed native Word renderer input'
  }
  $expectedHash = Body-Fingerprint $request.bodyText
  # Attach only. Never create Word, open files, activate windows, select, or change a view.
  $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
  $ownedReferences.Add($word)
  $documents = $word.Documents
  $ownedReferences.Add($documents)
  $matchingDocumentIndexes = New-Object 'System.Collections.Generic.List[int]'
  $diagnostics = New-Object 'System.Collections.Generic.List[object]'
  for ($documentIndex = 1; $documentIndex -le [int]$documents.Count; $documentIndex++) {
    $candidate = $null
    try {
      $candidate = $documents.Item($documentIndex)
      $candidateContent=$null
      try {
        $candidateContent=$candidate.Content
        $diagnostic=Body-Diagnostics ([string]$candidateContent.Text) $request.bodyText
        $shapes=$null; $shapeRangeCodes=New-Object 'System.Collections.Generic.List[object]'
        try {
          $shapes=$candidate.InlineShapes
          $diagnostic.inlineShapeCount=[int]$shapes.Count
          $diagnostic.inlineRangesExcludedTableProjectedHash=Body-Fingerprint ((Projected-DocumentBody $candidate).Replace(([string][char]13+[char]7),[string][char]9))
          for($shapeIndex=1;$shapeIndex -le [Math]::Min([int]$shapes.Count,100);$shapeIndex++) {
            $shape=$null; $shapeRange=$null
            try { $shape=$shapes.Item($shapeIndex); $shapeRange=$shape.Range; $codes=@(([string]$shapeRange.Text).ToCharArray() | ForEach-Object {[int]$_}); $shapeRangeCodes.Add($codes) } finally { Release-OwnedCom $shapeRange; Release-OwnedCom $shape }
          }
          $diagnostic.inlineShapeRangeCodes=$shapeRangeCodes.ToArray()
        } finally { Release-OwnedCom $shapes }
        $diagnostics.Add($diagnostic)
      } finally { Release-OwnedCom $candidateContent }
      if ((Document-Fingerprint $candidate) -ceq $expectedHash) { $matchingDocumentIndexes.Add($documentIndex) }
    } finally { Release-OwnedCom $candidate }
  }
  if ($matchingDocumentIndexes.Count -eq 0) { Fail-Native 'not_found' ('No existing Word document exactly matches the Office.js body fingerprint; metadata=' + (@{expected=(Body-Diagnostics $request.bodyText $request.bodyText);candidates=$diagnostics.ToArray()} | ConvertTo-Json -Compress -Depth 8)) }
  if ($matchingDocumentIndexes.Count -ne 1) { Fail-Native 'requirement' 'More than one existing Word document matches the body fingerprint; document selection is ambiguous' }
  $matchedDocument = $documents.Item($matchingDocumentIndexes[0])
  $ownedReferences.Add($matchedDocument)
  if ((Document-Fingerprint $matchedDocument) -cne $expectedHash) { Fail-Native 'requirement' 'Word document changed before native page access' }
  $window = $matchedDocument.ActiveWindow
  $ownedReferences.Add($window)
  $panes = $window.Panes
  $ownedReferences.Add($panes)
  $pane = $panes.Item(1)
  $ownedReferences.Add($pane)
  $view = $pane.View
  $ownedReferences.Add($view)
  if ([int]$view.Type -ne 3) { Fail-Native 'requirement' 'Native Word page rendering requires the existing document to already be in Print Layout view' }
  $pages = $pane.Pages
  $ownedReferences.Add($pages)
  $pageCount = [int]$pages.Count
  if ($pageCount -ne [int]$request.expectedPageCount -or $pageCount -lt 1 -or $pageCount -gt 10) { Fail-Native 'requirement' 'Native Word page count does not match expectedPageCount (1..10)' }
  Add-Type -AssemblyName System.Drawing
  $renderedPages = New-Object 'System.Collections.Generic.List[object]'
  for ($pageIndex = 1; $pageIndex -le $pageCount; $pageIndex++) {
    $page = $null; $emfStream = $null; $metafile = $null; $bitmap = $null; $graphics = $null; $pngStream = $null
    try {
      $page = $pages.Item($pageIndex)
      $widthPt = [double]$page.Width; $heightPt = [double]$page.Height
      if ([double]::IsNaN($widthPt) -or [double]::IsInfinity($widthPt) -or [double]::IsNaN($heightPt) -or [double]::IsInfinity($heightPt) -or $widthPt -le 0 -or $heightPt -le 0) { Fail-Native 'requirement' 'Native Word page has invalid dimensions' }
      $scale = [Math]::Min(2.0, 1600.0 / $widthPt)
      $width = [int][Math]::Min(1600, [Math]::Ceiling($widthPt * $scale))
      $height = [int][Math]::Ceiling($heightPt * $scale)
      if ($height -lt 1 -or $height -gt 10000) { Fail-Native 'requirement' 'Native Word page dimensions exceed the bounded raster size' }
      [byte[]]$bits = $page.EnhMetaFileBits
      if ($bits.Length -lt 1 -or $bits.Length -gt 67108864) { Fail-Native 'requirement' 'Native Word page EMF is unavailable or exceeds its bounded size' }
      $emfStream = [IO.MemoryStream]::new($bits)
      $metafile = [Drawing.Imaging.Metafile]::new($emfStream)
      $bitmap = [Drawing.Bitmap]::new($width, $height, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
      $bitmap.SetResolution([single](72 * $scale), [single](72 * $scale))
      $graphics = [Drawing.Graphics]::FromImage($bitmap)
      $graphics.Clear([Drawing.Color]::White)
      $graphics.PageUnit = [Drawing.GraphicsUnit]::Pixel
      # Playback Word's complete native page EMF, including headers and footers.
      # No text layout, font substitution, document reconstruction, or PDF renderer.
      $graphics.DrawImage($metafile, [Drawing.RectangleF]::new(0, 0, [single]($widthPt * $scale), [single]($heightPt * $scale)))
      $pngStream = [IO.MemoryStream]::new()
      $bitmap.Save($pngStream, [Drawing.Imaging.ImageFormat]::Png)
      $renderedPages.Add([ordered]@{ index = $pageIndex; width = $width; height = $height; pageWidthPt = $widthPt; pageHeightPt = $heightPt; dpi = 72 * $scale; base64 = [Convert]::ToBase64String($pngStream.ToArray()) })
    } finally {
      if ($null -ne $graphics) { $graphics.Dispose() }
      if ($null -ne $bitmap) { $bitmap.Dispose() }
      if ($null -ne $metafile) { $metafile.Dispose() }
      if ($null -ne $pngStream) { $pngStream.Dispose() }
      if ($null -ne $emfStream) { $emfStream.Dispose() }
      Release-OwnedCom $page
    }
  }
  if ((Document-Fingerprint $matchedDocument) -cne $expectedHash -or [int]$pages.Count -ne $pageCount) { Fail-Native 'requirement' 'Word document changed during native page rendering; output discarded' }
  $output = [ordered]@{ ok = $true; result = [ordered]@{ pages = @($renderedPages.ToArray()); pageCount = $pageCount; renderer = 'Microsoft Word native Page.EnhMetaFileBits'; documentFingerprint = $expectedHash; viewType = 'PrintLayout'; nominalDpi = 144 } }
} catch {
  $code = 'requirement'
  if ($_.Exception.Data.Contains('nativeCode')) { $code = [string]$_.Exception.Data['nativeCode'] }
  $output = [ordered]@{ ok = $false; code = $code; error = 'Native Word page rendering unavailable: ' + $_.Exception.Message }
} finally {
  for ($referenceIndex = $ownedReferences.Count - 1; $referenceIndex -ge 0; $referenceIndex--) { Release-OwnedCom $ownedReferences[$referenceIndex] }
}
[Console]::Out.Write(($output | ConvertTo-Json -Depth 8 -Compress))
