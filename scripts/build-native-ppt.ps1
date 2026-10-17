# Native offline PPT acceptance. Only creates/closes its own presentation; never quits Office.
[CmdletBinding()]
param([string[]]$Preset=@('ppt-clean-blue-deck','ppt-modern-accent'),[string]$RepoRoot=(Split-Path $PSScriptRoot -Parent),[switch]$PlanOnly)
$ErrorActionPreference='Stop'
$RepoRoot=(Resolve-Path $RepoRoot).Path
$skillRoot=Join-Path $RepoRoot 'skills/office-document-styles'
$outRoot=Join-Path $RepoRoot 'docs/examples'
$plansRoot=Join-Path $outRoot 'plans'
$null=New-Item -ItemType Directory -Force -Path $outRoot,$plansRoot
$planner=Join-Path $skillRoot 'scripts/plan.mjs'
$evidencePath=Join-Path $outRoot 'ppt-evidence.json'
$evidence=[ordered]@{schemaVersion=1;generatedAt=(Get-Date).ToString('o');generator='PowerPoint native COM';scopePolicy='Only self-created presentation; no ActivePresentation, application Quit, global Visible/DisplayAlerts, clipboard, or user file open/save.';results=@();exceptions=@()}
function Write-Evidence { $evidence|ConvertTo-Json -Depth 80|Set-Content -LiteralPath $evidencePath -Encoding utf8 }
function Prop($O,[string]$N,$Default=$null){if($null -ne $O -and $null -ne $O.PSObject.Properties[$N]){return $O.$N};return $Default}
function Color([string]$Hex){
 $h=$Hex.TrimStart('#');if($h -notmatch '^[0-9a-fA-F]{6}$'){throw "Unsupported color: $Hex"}
 return [Convert]::ToInt32($h.Substring(0,2),16)+256*[Convert]::ToInt32($h.Substring(2,2),16)+65536*[Convert]::ToInt32($h.Substring(4,2),16)
}
function Snapshot($Collection){
 $items=@();for($i=1;$i -le $Collection.Count;$i++){$o=$Collection.Item($i);$items += [ordered]@{name=[string]$o.Name;fullName=[string]$o.FullName;saved=[string]$o.Saved}}
 return ,$items
}
function Output-Path([string]$Id,[string]$Ext){
 if($Id -notmatch '^ppt-[a-z0-9-]+$'){throw 'Unsafe preset id'}
 $p=[IO.Path]::GetFullPath((Join-Path $outRoot "$Id.$Ext"))
 if((Split-Path $p -Parent) -ne $outRoot){throw 'Output escapes generated directory'}
 # Only an exact generated preset artifact, never an arbitrary user document.
 if(Test-Path -LiteralPath $p){Remove-Item -LiteralPath $p -Force};return $p
}
function Add-Element($Slide,$E,[double]$Sx,[double]$Sy,$R,[int]$Index){
 $x=[single]($E.x*$Sx);$y=[single]($E.y*$Sy);$w=[single]($E.w*$Sx);$h=[single]($E.h*$Sy)
 switch($E.type){
  'text' {$shape=$Slide.Shapes.AddTextbox(1,$x,$y,$w,$h)}
  'shape' {
   $types=@{Rectangle=1;RoundedRectangle=5;Ellipse=9;Triangle=7;RightTriangle=8;Diamond=4;Chevron=52;Pentagon=51;Hexagon=10;RightArrow=33;LeftArrow=34;DownArrow=36;UpArrow=35;Parallelogram=2;Trapezoid=3;Arc=25}
   $type=Prop $E 'shapeType' 'Rectangle';if(!$types.ContainsKey($type)){throw "Unknown native shape: $type"}
   $shape=$Slide.Shapes.AddShape($types[$type],$x,$y,$w,$h)
  }
  'line' {
   $shape=$Slide.Shapes.AddLine($x,$y,($x+$w),($y+$h))
   $shape.Line.ForeColor.RGB=Color (Prop $E 'color' '#64748B');$shape.Line.Weight=[single](Prop $E 'lineWidth' 1)
   if($null -ne $E.transparency){$shape.Line.Transparency=[single]$E.transparency}
   if($E.dashStyle){$shape.Line.DashStyle=@{Solid=1;Dash=4;DashDot=5;LongDash=7;RoundDot=3;SquareDot=2}[[string]$E.dashStyle]}
  }
  default {throw "Unsupported PPT element $($E.type)"}
 }
 $shape.Name='recipe-{0}-{1}' -f $Index,$Slide.Shapes.Count
 $shape.Tags.Add('recipeRole',[string](Prop $E 'role' ''))
 if($null -ne $E.rotation){$shape.Rotation=[single]$E.rotation}
 if($E.type -eq 'line'){return}
 if($E.fill){$shape.Fill.Visible=-1;$shape.Fill.Solid();$shape.Fill.ForeColor.RGB=Color $E.fill}else{$shape.Fill.Visible=0}
 if($E.lineColor -and (Prop $E 'lineWidth' 1) -gt 0){$shape.Line.Visible=-1;$shape.Line.ForeColor.RGB=Color $E.lineColor;$shape.Line.Weight=[single](Prop $E 'lineWidth' 1)}else{$shape.Line.Visible=0}
 if($null -ne $E.fillTransparency){$shape.Fill.Transparency=[single]$E.fillTransparency}
 if($E.role -eq 'background'){
  if($Slide.Shapes.Count -ne 1 -or $E.type -ne 'shape' -or $E.shapeType -ne 'Rectangle' -or $x -ne 0 -or $y -ne 0 -or $w -ne 960 -or $h -ne 540 -or $E.text){throw 'Only the first full-page empty rectangle may be a background'}
  $R.backgrounds += [ordered]@{slide=$Index+1;fill=$E.fill;nativeRgb=[int]$shape.Fill.ForeColor.RGB;box=@($x,$y,$w,$h)};return
 }
 $tf=$shape.TextFrame;$tf.AutoSize=0;$tf.WordWrap=-1
 foreach($edge in @('Left','Right','Top','Bottom')){$tf."Margin$edge"=[single](Prop $E.margins $edge.ToLowerInvariant() 0)}
 $tf.VerticalAnchor=@{top=1;middle=3;center=3;bottom=4}[[string](Prop $E 'valign' 'Top').ToLowerInvariant()]
 $tr=$tf.TextRange;$tr.Text=[string](Prop $E 'text' '')
 $tr.Font.Name=[string](Prop $E 'font' 'Microsoft YaHei');$tr.Font.NameFarEast=[string](Prop $E 'font' 'Microsoft YaHei')
 $tr.Font.Size=[single](Prop $E 'size' 18);$tr.Font.Bold=[int]([bool](Prop $E 'bold' $false))*-1;$tr.Font.Italic=[int]([bool](Prop $E 'italic' $false))*-1
 $tr.Font.Color.RGB=Color (Prop $E 'color' '#172033')
 $tr.ParagraphFormat.Alignment=@{left=1;center=2;right=3;justify=4}[[string](Prop $E 'align' 'Left').ToLowerInvariant()]
 $tr.ParagraphFormat.SpaceBefore=0;$tr.ParagraphFormat.SpaceAfter=0
 if($tr.Text){
  $R.fonts += [ordered]@{slide=$Index+1;shape=[string]$shape.Name;role=(Prop $E 'role');font=[string]$tr.Font.Name;points=[double]$tr.Font.Size;color=$E.color;nativeColorRgb=[int]$tr.Font.Color.RGB;text=[string]$tr.Text;box=@($x,$y,$w,$h);boundWidth=[double]$tr.BoundWidth;boundHeight=[double]$tr.BoundHeight}
  if($tr.BoundHeight -gt ($h-$tf.MarginTop-$tf.MarginBottom+2) -or $tr.BoundWidth -gt ($w-$tf.MarginLeft-$tf.MarginRight+2)){throw "Native text overflow slide $($Index+1), role $($E.role), font $($tr.Font.Name) $($tr.Font.Size)pt; box=[$x,$y,$w,$h]; usable=$($w-$tf.MarginLeft-$tf.MarginRight)x$($h-$tf.MarginTop-$tf.MarginBottom); bounds=$($tr.BoundWidth)x$($tr.BoundHeight): $($tr.Text)"}
 }
}
function Build-Ppt($Plan,$R){
 $add=@($Plan.steps|Where-Object action -eq 'ppt_add_slides')
 if($add.Count -ne 1 -or $add[0].args.count -ne 8){throw 'Refusing obsolete PPT: exactly eight recipe slides required'}
 if(@($Plan.steps|Where-Object action -eq 'ppt_add_elements').Count -ne 8){throw 'Eight slide element steps required'}
 $app=New-Object -ComObject PowerPoint.Application;$before=Snapshot $app.Presentations;$deck=$null
 try{
  $deck=$app.Presentations.Add(0);$deck.PageSetup.SlideWidth=960;$deck.PageSetup.SlideHeight=540
  $boundSlides=@()
  foreach($s in $Plan.steps){$a=$s.args;switch($s.action){
   'ppt_add_slides' {for($i=0;$i -lt $a.count;$i++){$null=$deck.Slides.Add($deck.Slides.Count+1,12)}}
   'ppt_add_elements' {
    $binding=@($s.bindings|Where-Object path -eq 'slideId');if($binding.Count -ne 1){throw 'Missing unique slide binding'}
    $index=[int]$binding[0].resultPath[1];if($index -lt 0 -or $index -gt 7 -or $index -in $boundSlides){throw 'Invalid or duplicate slide binding'};$boundSlides += $index
    if($a.targetSize.width -ne 960 -or $a.targetSize.height -ne 540){throw 'Native plan size must be 960x540 points'}
    $slide=$deck.Slides.Item($index+1)
    foreach($e in $a.elements){Add-Element $slide $e (960/[double]$a.coordinateSpace.width) (540/[double]$a.coordinateSpace.height) $R $index}
   }
   {$_ -in @('ppt_read_presentation_info','ppt_inspect_slides','ppt_render_slides')} {}
   default {throw "Unsupported PPT action $($s.action)"}
  }}
  $native=Output-Path $Plan.presetId 'pptx';$pdf=Output-Path $Plan.presetId 'pdf'
  $deck.SaveAs($native,24);$deck.SaveAs($pdf,32)
  $R.pdfExportMethod='PowerPoint Presentation.SaveAs ppSaveAsPDF (32)'
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip=[IO.Compression.ZipFile]::OpenRead($native)
  try{$R.nativeSlideXmlCount=@($zip.Entries|Where-Object{$_.FullName -match '^ppt/slides/slide\d+\.xml$'}).Count;$R.nativeChartXmlCount=@($zip.Entries|Where-Object{$_.FullName -match '^ppt/charts/chart\d+\.xml$'}).Count}finally{$zip.Dispose()}
  if($R.nativeSlideXmlCount -ne 8){throw 'Native PPTX package does not contain eight slides'}
  $R.slideCount=[int]$deck.Slides.Count;$R.nativeChartCount=0;$R.editableShapeCount=0;$R.slides=@()
  foreach($slide in $deck.Slides){
   $png=Join-Path $outRoot ('{0}-slide-{1:D2}.png' -f $Plan.presetId,$slide.SlideIndex)
   $slide.Export($png,'PNG',1600,900);$R.editableShapeCount += [int]$slide.Shapes.Count
   $bytes=[IO.File]::ReadAllBytes($png);$pw=[int]$bytes[16]*16777216+[int]$bytes[17]*65536+[int]$bytes[18]*256+[int]$bytes[19];$ph=[int]$bytes[20]*16777216+[int]$bytes[21]*65536+[int]$bytes[22]*256+[int]$bytes[23]
   if($pw -ne 1600 -or $ph -ne 900){throw "Unexpected native PNG dimensions: ${pw}x${ph}"}
   $R.slides += [ordered]@{slide=[int]$slide.SlideIndex;shapeCount=[int]$slide.Shapes.Count;png=$png;pngWidth=$pw;pngHeight=$ph}
  }
  $R.saved=[string]$deck.Saved
  $R.manualItems += 'Charts/table diagrams are editable text/shape/line compositions, not native Chart. Native slide PNG/PDF; no browser rendering. PDF rasterization delegated separately.'
 }finally{
  if($deck){$deck.Saved=-1;$deck.Close()}
  $after=Snapshot $app.Presentations;$same=($before|ConvertTo-Json -Depth 8 -Compress) -ceq ($after|ConvertTo-Json -Depth 8 -Compress)
  $R.objectScope=[ordered]@{existingBefore=$before;existingAfter=$after;unchanged=$same;onlyCreatedFileClosed=$true;applicationQuitCalled=$false}
  if(!$same){throw 'Existing presentation collection changed; inspect scope evidence'}
 }
}
try{
 foreach($id in $Preset){
  if($id -notmatch '^ppt-[a-z0-9-]+$'){throw 'Only safe PPT presets accepted'}
  $inputFile=Join-Path $skillRoot "resources/examples/$id.json"
  $planJson=& node $planner build $id $inputFile 'OFFLINE-OWNED-NATIVE-EXAMPLE'
  if($LASTEXITCODE -ne 0){throw "Planner failed for $id"};$plan=($planJson -join "`n")|ConvertFrom-Json
  if($plan.presetId -cne $id -or $plan.host -cne 'PowerPoint'){throw 'Plan identity mismatch'}
  ($planJson -join "`n")|Set-Content -LiteralPath (Join-Path $plansRoot "$id.json") -Encoding utf8
  $result=[ordered]@{presetId=$id;host='PowerPoint';status='building';theme=$plan.pptRecipe.theme;backgrounds=@();fonts=@();files=@();manualItems=@();planManualChecks=$plan.manualChecks;planPath=(Join-Path $plansRoot "$id.json")}
  $evidence.results += $result;Write-Evidence
  if($PlanOnly){$result.status='planned';continue}
  Write-Host "Building native PowerPoint: $id";Build-Ppt $plan $result
  $result.files=@(Get-ChildItem -LiteralPath $outRoot -File|Where-Object{$_.Name -like "$id.*" -or $_.Name -like "$id-slide-*.png"}|ForEach-Object{[ordered]@{path=$_.FullName;bytes=$_.Length}})
  $result.status='exported-native; visual review required';Write-Evidence
 };Write-Evidence
}catch{
 $evidence.exceptions += [ordered]@{message=$_.Exception.Message;detail=[string]$_;scriptStackTrace=$_.ScriptStackTrace;timestamp=(Get-Date).ToString('o')}
 if($result){$result.status='failed';$failedId=$result.presetId;$result.files=@(Get-ChildItem -LiteralPath $outRoot -File|Where-Object{$_.Name -like "$failedId.*" -or $_.Name -like "$failedId-slide-*.png"}|ForEach-Object{[ordered]@{path=$_.FullName;bytes=$_.Length}})};Write-Evidence;throw
}
