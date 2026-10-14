# Native QA renderer: translates the actual generated plan into owned COM documents only.
# Never uses ActiveDocument/ActiveWorkbook, changes global Office options, saves user files, or quits Office.
[CmdletBinding()]
param(
  [string[]]$PresetIds = @('word-business-blue','word-plain-accessible','excel-budget-green','excel-plain-data'),
  [string]$OutputDirectory,
  [switch]$Build
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo 'docs/examples' }
$planScript = Join-Path $PSScriptRoot 'plan.mjs'
$missing = [Type]::Missing
function Color-Value([string]$hex) {
  $h = $hex.TrimStart('#')
  return [Convert]::ToInt32($h.Substring(0,2),16) + 256*[Convert]::ToInt32($h.Substring(2,2),16) + 65536*[Convert]::ToInt32($h.Substring(4,2),16)
}
function Has($obj,[string]$name) { return $null -ne $obj -and $null -ne $obj.PSObject.Properties[$name] }
function Target-Range($wb,[string]$address) {
  $bang = $address.LastIndexOf('!')
  if ($bang -lt 1) { throw "Explicit worksheet required: $address" }
  $sheetName = $address.Substring(0,$bang).Trim("'")
  return ,($wb.Worksheets.Item($sheetName).Range($address.Substring($bang+1)))
}
function Apply-ExcelFormat($r,$a) {
  if (Has $a 'font') {$r.Font.Name=[string]$a.font}
  if (Has $a 'fontName') {$r.Font.Name=[string]$a.fontName}
  if (Has $a 'size') {$r.Font.Size=[double]$a.size}
  if (Has $a 'fontSize') {$r.Font.Size=[double]$a.fontSize}
  if (Has $a 'bold') {$r.Font.Bold=[bool]$a.bold}
  if (Has $a 'fill') {$r.Interior.Color=Color-Value $a.fill}
  if (Has $a 'color') {$r.Font.Color=Color-Value $a.color}
  if (Has $a 'numberFormat') {$r.NumberFormat=[string]$a.numberFormat}
  if (Has $a 'wrapText') {$r.WrapText=[bool]$a.wrapText}
  if (Has $a 'rowHeight') {$r.RowHeight=[double]$a.rowHeight}
  if (Has $a 'columnWidth') {
    # Excel COM ColumnWidth is in characters; recipe widths are points. Calibrate against actual Width.
    foreach($column in $r.Columns) {
      $column.ColumnWidth=10
      $column.ColumnWidth=[Math]::Min(255,[Math]::Max(1,10*[double]$a.columnWidth/[double]$column.Width))
      for($pass=0;$pass -lt 2;$pass++) { $column.ColumnWidth=[Math]::Min(255,[Math]::Max(1,[double]$column.ColumnWidth*[double]$a.columnWidth/[double]$column.Width)) }
    }
  }
  $align = if(Has $a 'horizontalAlignment') {$a.horizontalAlignment} elseif(Has $a 'align') {$a.align} else {$null}
  if($align) { $r.HorizontalAlignment=@{Left=-4131;Center=-4108;Right=-4152;Justify=-4130}[[string]$align] }
  if(Has $a 'verticalAlignment') {$r.VerticalAlignment=@{Top=-4160;Center=-4108;Bottom=-4107}[[string]$a.verticalAlignment]}
  if((Has $a 'autoFitRows') -and $a.autoFitRows) {$null=$r.Rows.AutoFit()}
}
function Set-NativeWordPageFooter($doc) {
  $fontName=[string][char]0x5B8B+[string][char]0x4F53
  foreach($section in $doc.Sections) {
    $footer=$section.Footers.Item(1)
    $footer.Range.Text=''
    $anchor=$footer.Range.Duplicate
    $anchor.Collapse(1)
    $field=$footer.Range.Fields.Add($anchor,33,$missing,$true) # wdFieldPage; dynamic PAGE, never static text
    $footer.Range.Font.Name=$fontName
    $footer.Range.Font.NameFarEast=$fontName
    $footer.Range.Font.Size=10.5
    $footer.Range.Font.Bold=0
    $footer.Range.Font.Color=0
    $footer.Range.ParagraphFormat.Alignment=1
    $footer.Range.ParagraphFormat.FirstLineIndent=0
    $footer.Range.ParagraphFormat.SpaceBefore=0
    $footer.Range.ParagraphFormat.SpaceAfter=0
    $null=$field.Update()
  }
}
function Build-Word($app,$plan,[string]$base) {
  $doc=$null
  $warnings=[Collections.Generic.List[string]]::new()
  try {
    $doc=$app.Documents.Add($missing,$false,0,$false)
    $doc.PageSetup.PaperSize=7 # wdPaperA4
    foreach($edge in @('TopMargin','BottomMargin','LeftMargin','RightMargin')) {$doc.PageSetup.$edge=25*72/25.4}
    $doc.PageSetup.TextColumns.SetCount(1)
    $styles=@{Normal=-1;Title=-63;Heading1=-2;Heading2=-3}
    $steps=0
    foreach($step in $plan.steps) {
      $a=$step.args
      switch($step.action) {
        'insert_paragraph' {
          if($a.location -ne 'start') {throw 'Native Word QA expects start insertion only'}
          $range=$doc.Range(0,0)
          $range.InsertBefore([string]$a.text+"`r")
          $steps++
        }
        'word_format_paragraphs' {
          foreach($target in $a.targets) {
            $para=$doc.Paragraphs.Item([int]$target.index+1)
            $actual=$para.Range.Text.TrimEnd([char]13,[char]7)
            if($actual -cne [string]$target.expectedText) {throw "Paragraph mismatch at $($target.index)"}
            if(-not $styles.ContainsKey([string]$target.style)) {throw "Unsupported semantic style $($target.style)"}
            $para.Range.Style=$styles[[string]$target.style]
            $f=$target.font
            $para.Range.Font.Name=[string]$f.name
            $para.Range.Font.NameFarEast=[string]$f.name
            $para.Range.Font.Size=[double]$f.size
            $para.Range.Font.Bold=[int][bool]$f.bold
            $para.Range.Font.Color=Color-Value $f.color
            $fmt=$para.Format
            $fmt.Alignment=@{left=0;center=1;right=2;justify=3}[[string]$target.alignment]
            $fmt.FirstLineIndent=[double]$target.firstLineIndent
            $fmt.LeftIndent=[double]$target.leftIndent
            $fmt.RightIndent=[double]$target.rightIndent
            if([double]$target.lineSpacing -ne 18) {throw 'Expected Office.js 18 => native 1.5-line paragraph rule'}
            $fmt.LineSpacingRule=1 # wdLineSpace1pt5 (not an exact 18-point rule)
            $fmt.SpaceBefore=[double]$target.spaceBefore
            $fmt.SpaceAfter=[double]$target.spaceAfter
            $fmt.WidowControl=-1
            if($target.style -in @('Heading1','Heading2','Title')) {$fmt.KeepWithNext=-1}
          }
          $steps++
        }
        {$_ -in @('read_document','word_list_paragraphs')} { }
        default {throw "Unmapped Word plan action: $($step.action)"}
      }
    }
    Set-NativeWordPageFooter $doc
    $doc.Repaginate()
    $paragraphs=@(foreach($p in $doc.Paragraphs) {
      if($p.Range.Text.TrimEnd([char]13,[char]7).Length) {
        [ordered]@{text=$p.Range.Text.TrimEnd([char]13,[char]7);styleIdOrName=[string]$p.Range.Style.NameLocal;font=[string]$p.Range.Font.Name;eastAsiaFont=[string]$p.Range.Font.NameFarEast;size=[double]$p.Range.Font.Size;bold=[int]$p.Range.Font.Bold;alignment=[int]$p.Format.Alignment;firstLineIndent=[double]$p.Format.FirstLineIndent;lineSpacingRule=[int]$p.Format.LineSpacingRule;spaceBefore=[double]$p.Format.SpaceBefore;spaceAfter=[double]$p.Format.SpaceAfter}
      }
    })
    $doc.SaveAs2($base+'.docx',16)
    $doc.ExportAsFixedFormat($base+'.pdf',17)
    return [ordered]@{host='Word';executedWriteSteps=$steps;pages=[int]$doc.ComputeStatistics(2);paragraphCount=$paragraphs.Count;paragraphs=$paragraphs;warnings=@($warnings);manualQaOverrides=@('Native COM A4, 25mm margins, single column; these are not proof that the plugin applied page setup.','Office.js lineSpacing=18 mapped to Word COM LineSpacingRule=wdLineSpace1pt5.','Native footer PAGE manual QA: centered dynamic PAGE field, SimSun 10.5pt; not a plugin automatic field capability.');document=$base+'.docx';pdf=$base+'.pdf'}
  } finally {if($null -ne $doc) {$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}
}
function Set-NativeExcelDetailPrintHeader($ws) {
  $title=([string]$ws.Range('A1').Value2).Replace('&','&&')
  $metadata=([string]$ws.Range('A2').Value2 -split '\|')
  $subtitle=(($metadata|Select-Object -First 2|ForEach-Object {$_.Trim()}) -join ' | ')+' | DEMO'
  $subtitle=$subtitle.Replace('&','&&')
  if($title.Length -gt 80 -or $subtitle.Length -gt 80) {throw 'Native print header exceeds the reviewed 80-character line budget'}
  $ws.PageSetup.LeftHeader='&"Microsoft YaHei,Regular"&12'+$title+"`n"+'&9'+$subtitle
  $ws.PageSetup.TopMargin=60
  $ws.PageSetup.HeaderMargin=12
}
function Build-Excel($app,$plan,[string]$base) {
  $wb=$null
  $warnings=[Collections.Generic.List[string]]::new()
  $manual=[Collections.Generic.List[string]]::new()
  try {
    $wb=$app.Workbooks.Add(-4167) # xlWBATWorksheet: exactly one owned blank sheet
    $initial=@()
    $reuseInitial=$true
    $writeSteps=0
    foreach($step in $plan.steps) {
      $a=$step.args
      switch($step.action) {
        'add_sheet' {
          if($reuseInitial) {$ws=$wb.Worksheets.Item(1);$reuseInitial=$false}
          else {$ws=$wb.Worksheets.Add($missing,$wb.Worksheets.Item($wb.Worksheets.Count))}
          $ws.Name=[string]$a.name
          $writeSteps++
        }
        'write_range' {
          $r=Target-Range $wb $a.address
          if(Has $a 'formulas') {$matrix=$a.formulas} else {$matrix=$a.values}
          if($r.Rows.Count -ne $matrix.Count -or $r.Columns.Count -ne $matrix[0].Count) {throw "Matrix/range mismatch: $($a.address)"}
          for($row=0;$row -lt $matrix.Count;$row++) {for($col=0;$col -lt $matrix[$row].Count;$col++) {
            $value=$matrix[$row][$col]
            if(Has $a 'formulas') {$r.Cells.Item($row+1,$col+1).Formula=[string]$value}
            elseif($value -is [string]) {$r.Cells.Item($row+1,$col+1).Value2=[string]$value}
            else {$r.Cells.Item($row+1,$col+1).Value2=[double]$value}
          }}
          $writeSteps++
        }
        'format_range' {Apply-ExcelFormat (Target-Range $wb $a.address) $a;$writeSteps++}
        'create_table' {
          $r=Target-Range $wb $a.address
          $t=$r.Worksheet.ListObjects.Add(1,$r,$missing,1)
          $t.Name=[string]$a.name
          $t.TableStyle=[string]$a.style
          if(Has $a 'showBandedRows') {$t.ShowTableStyleRowStripes=[bool]$a.showBandedRows}
          if(Has $a 'showFilterButton') {$t.ShowAutoFilter=[bool]$a.showFilterButton}
          $writeSteps++
        }
        'excel_set_freeze_panes' {
          $ws=$wb.Worksheets.Item([string]$a.sheet)
          # Freeze is view-specific. Only the newly created workbook window is activated.
          $owned=$wb.Windows.Item(1)
          $owned.Activate()
          $ws.Activate()
          $owned.FreezePanes=$false
          if($a.mode -eq 'rows') {$owned.SplitRow=[int]$a.count;$owned.SplitColumn=0}
          else {throw "Unsupported freeze mode $($a.mode)"}
          $owned.FreezePanes=$true
          $writeSteps++
        }
        'add_conditional_format' {
          $r=Target-Range $wb $a.address
          if($a.type -eq 'CellValue') {
            $op=@{LessThan=6;GreaterThan=5;EqualTo=3;LessThanOrEqualTo=8;GreaterThanOrEqualTo=7}[[string]$a.rule.operator]
            if(-not $op) {throw "Unsupported conditional operator $($a.rule.operator)"}
            $cf=$r.FormatConditions.Add(1,$op,[string]$a.rule.formula1)
          } elseif($a.type -eq 'ContainsText') {
            $needle=([string]$a.rule.text).Replace('"','""')
            $local=$r.Cells.Item(1,1).Address($false,$false)
            $cf=$r.FormatConditions.Add(2,$missing,"=ISNUMBER(SEARCH(`"$needle`",$local))")
          } else {throw "Unsupported conditional type $($a.type)"}
          if(Has $a.format 'fillColor') {$cf.Interior.Color=Color-Value $a.format.fillColor}
          if(Has $a.format 'fontColor') {$cf.Font.Color=Color-Value $a.format.fontColor}
          if(Has $a.format 'bold') {$cf.Font.Bold=[bool]$a.format.bold}
          $writeSteps++
        }
        'insert_chart' {
          $r=Target-Range $wb $a.dataRange
          $ws=$r.Worksheet
          $start=$ws.Range([string]$a.position.startCell)
          $end=$ws.Range([string]$a.position.endCell)
          $obj=$ws.ChartObjects().Add([double]$start.Left,[double]$start.Top,[double]($end.Left+$end.Width-$start.Left),[double]($end.Top+$end.Height-$start.Top))
          $chart=$obj.Chart
          $chart.ChartType=51
          $null=$chart.SetSourceData($r)
          $chart.HasTitle=$true
          $chart.ChartTitle.Text=[string]$a.title
          $chart.ChartTitle.Font.Size=12
          $chart.ChartArea.Font.Name=[string]$plan.preset.tokens.font
          $chart.ChartArea.Font.Size=10
          $writeSteps++
        }
        {$_ -in @('list_sheets','read_range','excel_render_range')} { }
        default {throw "Unmapped Excel plan action: $($step.action)"}
      }
    }
    foreach($window in $wb.Windows) {$window.Visible=$false}
    $manual.Add('Owned workbook starts with exactly one blank sheet, reused for the first add_sheet plan step; three physical business sheets, no sheet deletion or global DisplayAlerts change.')
    $wb.Worksheets.Item([string]$plan.business.sheets.detail).Calculate()
    $wb.Worksheets.Item([string]$plan.business.sheets.overview).Calculate()
    $wb.Worksheets.Item([string]$plan.business.sheets.guide).Calculate()
    $sheets=@(foreach($ws in $wb.Worksheets) {
      if($ws.Visible -ne -1) {continue}
      $role=if($ws.Name -eq $plan.business.sheets.detail) {'detail'} elseif($ws.Name -eq $plan.business.sheets.overview) {'overview'} else {'guide'}
      $range=$ws.UsedRange
      $verify=@($plan.steps|Where-Object {$_.id -eq "verify-$role"})[0]
      if($null -eq $verify -or [string]$verify.args.address -notmatch '!([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$') {throw "Missing bounded verification range for $role"}
      $last=$Matches[3]
      $bottom=[int]$Matches[4]
      if($role -eq 'overview') {$bottom=[Math]::Min($bottom,[int]$range.Row+[int]$range.Rows.Count-1)}
      if($role -eq 'guide') {$last='B'}
      $ps=$ws.PageSetup
      $ps.PaperSize=9
      $ps.Orientation=2
      $ps.LeftMargin=18;$ps.RightMargin=18;$ps.TopMargin=18;$ps.BottomMargin=18
      $startRow=if($role -eq 'detail') {5} else {1}
      $ps.PrintArea="`$A`$$startRow`:`$$last`$$bottom"
      $ps.Zoom=90
      $ps.PrintTitleRows=if($role -eq 'detail') {'$5:$5'} else {''}
      if($role -eq 'detail') {Set-NativeExcelDetailPrintHeader $ws}
      $ps.PrintTitleColumns=if($role -eq 'detail') {'$A:$B'} else {''}
      $ps.CenterHorizontally=$false
      $ps.LeftFooter='DEMO / sample data'
      $ps.RightFooter='Page &P / &N'
      $errors=@(foreach($c in $range.Cells) {if([string]$c.Text -match '^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|SPILL!)$') {[ordered]@{cell=$c.Address();text=[string]$c.Text;formula=[string]$c.Formula}}})
      if($errors.Count) {throw "Formula errors on $($ws.Name): $($errors|ConvertTo-Json -Compress)"}
      [ordered]@{name=[string]$ws.Name;role=$role;usedRange=[string]$range.Address();tableCount=[int]$ws.ListObjects.Count;chartCount=[int]$ws.ChartObjects().Count;conditionalFormatCount=[int]$range.FormatConditions.Count;printArea=[string]$ps.PrintArea;printZoom=90;printTitleRows=[string]$ps.PrintTitleRows;printTitleColumns=[string]$ps.PrintTitleColumns;leftHeader=[string]$ps.LeftHeader;topMarginPt=[double]$ps.TopMargin;headerMarginPt=[double]$ps.HeaderMargin;landscape=$true;formulaErrors=$errors}
    })
    $overview=$wb.Worksheets.Item([string]$plan.business.sheets.overview)
    $kpis=@(foreach($row in 7..11) {$c=$overview.Cells.Item($row,2);[ordered]@{cell="B$row";value=$c.Value2;formula=[string]$c.Formula;display=[string]$c.Text}})
    $expected=if($plan.business.kind -eq 'budget') {@(482000,480800,1200,(480800/482000),$null)} else {@(16,137,96,(96/137),3)}
    foreach($i in 0..4) {if($null -ne $expected[$i] -and [Math]::Abs([double]$kpis[$i].value-[double]$expected[$i]) -gt 0.000001) {$warnings.Add("KPI $($kpis[$i].cell) actual=$($kpis[$i].value) differs from current demonstration acceptance value $($expected[$i]); inspect changed source input.")}}
    $manual.Add('Native QA only: A4 landscape, 18pt side/bottom margins, detail top margin 60pt/header margin 12pt, 90% print zoom; detail prints from row 5, repeats row 5 and columns A:B, and places complete title plus short period/unit/DEMO text in a two-line print header. Screen cells A1:A4 are unchanged. Wide detail columns intentionally span horizontal pages instead of shrinking body text below readability.')
    # Save a visible workbook window so the delivered file does not open as an invisible workbook.
    $wb.Windows.Item(1).Visible=$true
    $wb.SaveAs($base+'.xlsx',51)
    $wb.Windows.Item(1).Visible=$false
    $wb.ExportAsFixedFormat(0,$base+'.pdf',0,$true,$false)
    return [ordered]@{host='Excel';executedWriteSteps=$writeSteps;visibleSheetCount=$sheets.Count;physicalSheetCount=[int]$wb.Worksheets.Count;hiddenInitialSheets=$initial;sheets=$sheets;kpis=$kpis;warnings=@($warnings);manualQaOverrides=@($manual);document=$base+'.xlsx';pdf=$base+'.pdf'}
  } finally {if($null -ne $wb) {$wb.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}}
}
if(-not $Build) {Write-Output 'Script ready. Pass -Build after source plans are ready. No Office applications were accessed.';return}
$null=New-Item -ItemType Directory -Path $OutputDirectory -Force
$word=$null;$excel=$null
try {
  foreach($id in $PresetIds) {
    if($id -notin @('word-business-blue','word-plain-accessible','excel-budget-green','excel-plain-data')) {throw "Unsupported native QA preset: $id"}
    $inputFile=Join-Path $PSScriptRoot "../resources/examples/$id.json"
    $raw=(& node $planScript build $id $inputFile native-qa-instance) -join "`n"
    if($LASTEXITCODE -ne 0) {throw "Plan generator failed: $id"}
    $plan=$raw|ConvertFrom-Json
    $base=Join-Path $OutputDirectory "native-$id"
    foreach($ext in @('.docx','.xlsx','.pdf','.evidence.json')) {if(Test-Path ($base+$ext)) {throw "Refusing overwrite: $base$ext"}}
    Write-Output "Rendering $id ($($plan.steps.Count) actual plan steps)"
    if($plan.host -eq 'Word') {if($null -eq $word) {$word=New-Object -ComObject Word.Application};$evidence=Build-Word $word $plan $base}
    elseif($plan.host -eq 'Excel') {
      if($null -eq $excel) {$excel=New-Object -ComObject Excel.Application}
      $nativeOutput=@(Build-Excel $excel $plan $base)
      $records=@($nativeOutput|Where-Object {$_ -is [Collections.IDictionary] -and $_.Contains('host')})
      if($records.Count -ne 1) {throw "Expected one native Excel evidence record; got $($records.Count)"}
      $evidence=$records[0]
    }
    else {throw "Unsupported host $($plan.host)"}
    $evidence.presetId=$id
    $evidence.createdAt=[DateTimeOffset]::Now.ToString('o')
    $evidence.validationMethod='Actual generated plan translated to native Office COM; not Office.js/bridge end-to-end validation.'
    $evidence.sourcePlan=$plan
    $evidence.visualReview='PDF exported; page-by-page PNG inspection pending. Export success does not prove absence of clipping.'
    $evidence|ConvertTo-Json -Depth 100|Set-Content ($base+'.evidence.json') -Encoding utf8
    Write-Output "Created $base + PDF + evidence"
  }
} finally {
  # Release automation references only. Never Quit shared Office processes.
  if($null -ne $word) {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
  if($null -ne $excel) {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
