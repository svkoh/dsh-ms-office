param([ValidateSet('All','Word','Excel','PowerPoint')][string]$Only='All')
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../docs/style-atlas'))
[IO.Directory]::CreateDirectory($root)|Out-Null
[IO.Directory]::CreateDirectory((Join-Path $root 'assets'))|Out-Null
function RGB([string]$hex){$hex=$hex.TrimStart('#');return [Convert]::ToInt32($hex.Substring(0,2),16)+([Convert]::ToInt32($hex.Substring(2,2),16)-shl 8)+([Convert]::ToInt32($hex.Substring(4,2),16)-shl 16)}
function App([string]$name){return New-Object -ComObject $name}
function Snapshot($collection){$a=@();foreach($item in $collection){$a+=@{name=[string]$item.Name;saved=[bool]$item.Saved}};return $a}
function Evidence($id,$data){[IO.File]::WriteAllText((Join-Path $root ($id+'.evidence.json')),($data|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))}
$categories=@('研发','运营','营销','支持');$budgets=@(120,80,60,40);$actual=@(114,84,54,38)
Add-Type -AssemblyName System.Drawing
$image=Join-Path $root 'assets/original-workflow.png'
$bmp=New-Object Drawing.Bitmap 1200,400;$g=[Drawing.Graphics]::FromImage($bmp);$g.Clear([Drawing.Color]::FromArgb(250,249,246));$g.SmoothingMode=4
$f=New-Object Drawing.Font 'Microsoft YaHei',32;$small=New-Object Drawing.Font 'Microsoft YaHei',18;$ink=New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(36,40,45));$panel=New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(239,238,234));$pen=New-Object Drawing.Pen ([Drawing.Color]::FromArgb(131,99,68)),5
$labels=@('规划','制作','复核');$notes=@('先定义结构与口径','应用元素和版式','读回内容并检查效果')
for($i=0;$i -lt 3;$i++){$x=45+$i*395;$g.FillRectangle($panel,$x,75,320,250);$g.DrawString($labels[$i],$f,$ink,$x+38,125);$g.DrawString($notes[$i],$small,$ink,$x+25,220);if($i -lt 2){$g.DrawLine($pen,$x+330,200,$x+380,200)}}
$bmp.Save($image,[Drawing.Imaging.ImageFormat]::Png);$g.Dispose();$bmp.Dispose();$f.Dispose();$small.Dispose();$ink.Dispose();$panel.Dispose();$pen.Dispose()
if($Only -in @('All','Word')){
 $app=App 'Word.Application';$before=Snapshot $app.Documents;$script:doc=$null
 try{
  $script:doc=$app.Documents.Add();$doc=$script:doc;$doc.PageSetup.PaperSize=7;foreach($side in @('TopMargin','BottomMargin','LeftMargin','RightMargin')){$doc.PageSetup.$side=70.87};$doc.PageSetup.DifferentFirstPageHeaderFooter=-1
  foreach($code in @(-1,-2,-3,-4,-63,-35)){$st=$doc.Styles.Item($code);$st.Font.Name='宋体';$st.Font.NameFarEast='宋体';$st.Font.Size=12;$st.Font.Color=RGB '24282D';$st.ParagraphFormat.SpaceBefore=0;$st.ParagraphFormat.SpaceAfter=0;$st.ParagraphFormat.FirstLineIndent=0;$st.ParagraphFormat.LineSpacingRule=1}
  $doc.Styles.Item(-1).ParagraphFormat.FirstLineIndent=24
  foreach($code in @(-2,-3,-4)){$st=$doc.Styles.Item($code);$st.Font.Bold=-1;$st.ParagraphFormat.KeepWithNext=-1;$st.ParagraphFormat.SpaceBefore=12;$st.ParagraphFormat.SpaceAfter=6}
  $doc.Styles.Item(-2).Font.Name='黑体';$doc.Styles.Item(-2).Font.NameFarEast='黑体';$doc.Styles.Item(-2).Font.Size=14;$doc.Styles.Item(-63).Font.Name='黑体';$doc.Styles.Item(-63).Font.NameFarEast='黑体';$doc.Styles.Item(-63).Font.Size=18;$doc.Styles.Item(-35).Font.Size=10.5
  function P([string]$text,[int]$style=-1,[int]$alignment=3){$pos=$script:doc.Content.End-1;$r=$script:doc.Range($pos,$pos);$r.InsertAfter($text+"`r");$r=$script:doc.Range($pos,$pos+$text.Length);$r.Style=$style;$r.ListFormat.RemoveNumbers();$r.ParagraphFormat.Alignment=$alignment;return ,$r}
  function Page{$pos=$script:doc.Content.End-1;$r=$script:doc.Range($pos,$pos);$r.InsertBreak(7)}
  $r=P 'Office 文档基础样式范本' -63 1;$r.Font.Size=28;$r.ParagraphFormat.SpaceBefore=120;$r.ParagraphFormat.SpaceAfter=24
  $r=P '一份文件 · 常用元素 · 可复用样式' -1 1;$r.ParagraphFormat.FirstLineIndent=0;$r.Font.Size=16
  foreach($text in @('封面 / 目录 / 章节 / 段落 / 表格 / 图片 / 页眉页脚','适用：一般企业办公文档','资料性质：DEMO / 示例数据','版本：2026年10月')){$r=P $text -1 1;$r.ParagraphFormat.FirstLineIndent=0;$r.ParagraphFormat.SpaceBefore=12}
  Page;$null=P '目录' -63 0;$pos=$doc.Content.End-1;$toc=$doc.TablesOfContents.Add($doc.Range($pos,$pos));Page
  $null=P '一、章节、段落与列表' -2 0;$null=P '（一）标题层级与正文' -3 0;$null=P '1. 三级标题示例' -4 0
  $null=P '正文使用宋体12pt、真实1.5倍行距、首行缩进24pt和两端对齐。标题通过Word全局样式定义，目录读取实际标题层级。新内容应根据含义选用标题或正文样式，避免仅用加粗模拟章节。'
  $null=P '本范本以同一组演示数据说明文字、表格和图片的组织方式：预算合计300万元，实际支出290万元，余额10万元，执行率96.7%。这些数字仅用于样式展示，正式使用时须替换并核验。'
  $null=P '（二）对齐方式按用途选择' -3 0
  foreach($entry in @(@('左对齐：短说明、标签和列表。',0),@('居中：封面标题与独立图注。',1),@('右对齐：日期或签署位置。',2),@('两端对齐：连续正文保持左右边界整齐；段落最后一行自然结束，不应手工插入空格凑齐。',3))){$r=P $entry[0] -1 $entry[1];$r.ParagraphFormat.FirstLineIndent=0}
  $null=P '（三）真正的项目符号与编号' -3 0
  foreach($text in @('项目符号用于并列事项，保持同级缩进一致。','重要信息同时使用文字说明，颜色仅作辅助。')){$r=P $text -1 0;$r.ListFormat.ApplyBulletDefault()}
  foreach($text in @('先确定内容与数据口径。','再应用对应元素样式。')){$r=P $text -1 0;$r.ListFormat.ApplyNumberDefault()}
  Page;$null=P '二、表格与数值对齐' -2 0;$null=P '（一）简洁表头、边框与汇总' -3 0
  $null=P '表格使用真实Word表格对象。表头加粗并设为跨页重复，金额右对齐，文字左对齐；保留清楚的单位、汇总和示例标签。'
  $pos=$doc.Content.End-1;$t=$doc.Tables.Add($doc.Range($pos,$pos),6,5);$t.Range.Font.Name='宋体';$t.Range.Font.Size=11;$t.Range.ParagraphFormat.FirstLineIndent=0;$t.Range.ParagraphFormat.SpaceAfter=3;$t.Range.ParagraphFormat.LineSpacingRule=0;$t.Borders.Enable=1;$t.PreferredWidthType=3;$t.PreferredWidth=450
  $rows=@(@('类别','预算（万元）','实际（万元）','余额（万元）','说明'),@('研发','120','114','6','按计划执行'),@('运营','80','84','-4','复核超支原因'),@('营销','60','54','6','保留后续安排'),@('支持','40','38','2','控制零星支出'),@('合计','300','290','10','执行率96.7%'))
  for($i=0;$i -lt 6;$i++){for($j=0;$j -lt 5;$j++){$c=$t.Cell($i+1,$j+1);$c.Range.Text=$rows[$i][$j];$c.Range.ParagraphFormat.Alignment=if($j -in 1,2,3){2}else{0}}};$t.Rows.Item(1).HeadingFormat=-1;$t.Rows.Item(1).Range.Font.Bold=-1;$t.Rows.Item(1).Shading.BackgroundPatternColor=RGB 'EFEEEA';$t.Rows.Item(6).Range.Font.Bold=-1
  $null=P '表1　预算与实际对比（DEMO，单位：万元）' -35 1;$null=P '（二）强调与颜色' -3 0
  $r=P '默认以黑色正文传达信息；强调使用加粗，辅助标签可用哑铜色。运营类别超支4万元，需要结合原因与动作复核。' -1 3;$r.Words.Item(1).Font.Bold=-1
  $r=P '辅助标签示例：待复核 / 示例数据 / 统计口径' -1 0;$r.Font.Color=RGB '836344';$r.ParagraphFormat.FirstLineIndent=0
  Page;$null=P '三、图片、图注与页眉页脚' -2 0;$null=P '（一）图片嵌入与宽高比' -3 0
  $null=P '下图是本项目原创PNG，作为图片元素嵌入文档。保持原始宽高比，添加替代文字，并用图注说明用途；图像不能替代可编辑正文。'
  $pos=$doc.Content.End-1;$pic=$doc.InlineShapes.AddPicture($image,$false,$true,$doc.Range($pos,$pos));$pic.LockAspectRatio=-1;$pic.Width=440;$pic.AlternativeText='规划、制作和复核三个阶段的原创流程插图';$null=P '' -1 0;$null=P '图1　从内容规划到效果复核（原创示意图）' -35 1
  $null=P '（二）页眉、页脚与自动字段' -3 0;$null=P '封面不显示页眉页脚；后续页面保留统一页眉和居中动态PAGE页码。目录为真实Word目录字段，可在修改标题后更新。当前范本中的字段与样式已存在，桥接重建时应分别核对对应API能力。'
  $null=P '（三）复用方式' -3 0;$null=P '先另存工作副本，再替换演示内容。保留标题层级、表格结构和图片比例；按正文、表格、图注等元素选择样式。完成后更新目录，逐页检查标题孤行、表格跨页、图片边界与页码。'
  $head=$doc.Sections.Item(1).Headers.Item(1).Range;$head.Text='Office 基础样式范本 · DEMO';$head.Font.Name='宋体';$head.Font.Size=10.5;$head.ParagraphFormat.FirstLineIndent=0
  $foot=$doc.Sections.Item(1).Footers.Item(1).Range;$foot.ParagraphFormat.Alignment=1;$foot.ParagraphFormat.FirstLineIndent=0;$null=$foot.Fields.Add($foot,-1,'PAGE',$true);$foot.Font.Name='宋体';$foot.Font.Size=10.5
  $doc.Repaginate();$toc.Update();$doc.Repaginate();$toc.UpdatePageNumbers();$doc.SaveAs2((Join-Path $root 'word-base-styles.docx'),16);$doc.ExportAsFixedFormat((Join-Path $root 'word-base-styles.pdf'),17)
  $pages=$doc.ComputeStatistics(2);$counts=@{tables=$doc.Tables.Count;images=$doc.InlineShapes.Count;toc=$doc.TablesOfContents.Count;sections=$doc.Sections.Count;pages=$pages;fields=$doc.Fields.Count};$doc.Close(0);$script:doc=$null;$after=Snapshot $app.Documents
  Evidence 'word-base-styles' @{id='word-base-styles';host='Word';scope='Native COM sample; not Office.js E2E';pdfPageCount=$pages;counts=$counts;coverage=@('cover','toc','heading-1-2-3','paragraph','native-lists','alignment','native-table','color','image','caption','header','footer','PAGE');objectScope=@{before=$before;after=$after;unchanged=(($before|ConvertTo-Json -Compress) -eq ($after|ConvertTo-Json -Compress))}}
  Write-Output ('Word atlas saved: '+$pages+' pages')
 }finally{if($script:doc){$script:doc.Close(0);$script:doc=$null}}
}
if($Only -in @('All','Excel')){
 $app=App 'Excel.Application';$before=Snapshot $app.Workbooks;$book=$null
 try{
  $book=$app.Workbooks.Add(-4167);$basic=$book.Worksheets.Item(1);$basic.Name='基础样式';$chartSheet=$book.Worksheets.Add();$chartSheet.Name='数据与图表';$chartSheet.Move([Type]::Missing,$basic)
  foreach($s in @($basic,$chartSheet)){$s.Range('A1:H32').Font.Name='Microsoft YaHei';$s.Range('A1:H32').Font.Size=11;$s.Range('A1:H32').RowHeight=18;$s.Range('A1:H32').ColumnWidth=12;$s.Cells.Item(2,1).Value2='DEMO / 示例数据 · 单位：万元';$s.Range('A1:H1').Font.Size=18;$s.Range('A1:H1').Font.Bold=$true;$s.Range('A1:H1').RowHeight=34}
  $basic.Cells.Item(1,1).Value2='Excel 基础样式：文字、表格与自动布局';$chartSheet.Cells.Item(1,1).Value2='Excel 图表样式：数据与原生图表'
  $categories=@('研发','运营','营销','支持');$budgets=@(120,80,60,40);$actual=@(114,84,54,38);$headers=@('编号','费用类别','预算（万元）','实际（万元）','余额（万元）','备注')
  foreach($s in @($basic,$chartSheet)){
   for($j=0;$j -lt 6;$j++){$s.Cells.Item(5,$j+1).Value2=$headers[$j]}
   for($i=0;$i -lt 4;$i++){$row=$i+6;$s.Cells.Item($row,1).Value2=('S{0:d2}' -f ($i+1));$s.Cells.Item($row,2).Value2=$categories[$i];$s.Cells.Item($row,3).Value2=[double]$budgets[$i];$s.Cells.Item($row,4).Value2=[double]$actual[$i];$s.Cells.Item($row,5).Formula=('=C{0}-D{0}' -f $row);$s.Cells.Item($row,6).Value2=if($i -eq 1){'实际高于预算，需复核原因并记录后续动作。'}else{'此列展示自动换行；设定列宽后自动调整行高。'}}
   $null=$s.ListObjects.Add(1,$s.Range('A5:F9'),[Type]::Missing,1);$s.ListObjects.Item(1).TableStyle='TableStyleLight1';$s.ListObjects.Item(1).ShowTotals=$true;$s.Range('A5:F5').Font.Bold=$true;$s.Range('A5:F5').Interior.Color=RGB 'EFEEEA';$s.Range('A5:F5').RowHeight=28;$s.Range('C6:E10').NumberFormat='#,##0.00';$s.Range('C6:E10').HorizontalAlignment=-4152;$s.Range('A5:E10').Columns.AutoFit()|Out-Null;$s.Range('F1:F32').ColumnWidth=31;$s.Range('F6:F9').WrapText=$true;$s.Range('A6:F9').Rows.AutoFit()|Out-Null
   $s.Cells.Item(10,2).Value2='合计';foreach($col in @('C','D','E')){$s.Range($col+'10').Formula=('=SUM({0}6:{0}9)' -f $col)};$s.Range('B10:F10').Font.Bold=$true;$s.Range('B10:F10').Borders.Item(8).LineStyle=1
   $cf=$s.Range('E6:E9').FormatConditions.Add(1,6,'0');$cf.Interior.Color=RGB 'F3E4E1';$cf.Font.Color=RGB '8A4031';$s.Calculate()
   $s.PageSetup.PaperSize=9;$s.PageSetup.Orientation=2;$s.PageSetup.Zoom=100;foreach($side in @('TopMargin','BottomMargin','LeftMargin','RightMargin')){$s.PageSetup.$side=24};$s.PageSetup.LeftFooter='DEMO / 示例数据';$s.PageSetup.RightFooter='&P / &N'
  }
  $basic.Cells.Item(14,2).Value2='对齐：按信息类型设置';$basic.Range('B14:F14').Font.Bold=$true;$basic.Cells.Item(15,2).Value2='左对齐';$basic.Cells.Item(15,3).Value2='居中';$basic.Cells.Item(15,4).Value2='右对齐';$basic.Range('B15').HorizontalAlignment=-4131;$basic.Range('C15').HorizontalAlignment=-4108;$basic.Range('D15').HorizontalAlignment=-4152
  $names=@('整数','千分位','百分比','日期');for($i=0;$i -lt 4;$i++){$basic.Cells.Item(18,$i+2).Value2=$names[$i]};$basic.Range('B18:E18').Font.Bold=$true;$basic.Range('B19').Value2=[double]16;$basic.Range('B19').NumberFormat='0';$basic.Range('C19').Value2=[double]12198;$basic.Range('C19').NumberFormat='#,##0';$basic.Range('D19').Formula='=IFERROR(290/300,0)';$basic.Range('D19').NumberFormat='0.0%';$basic.Range('E19').Value2=[double]46309;$basic.Range('E19').NumberFormat='yyyy-mm-dd';$basic.Range('E1:E32').ColumnWidth=13
  $basic.Cells.Item(23,2).Value2='复用顺序';$basic.Range('B23:F23').Font.Bold=$true;$basic.Cells.Item(24,2).Value2='先设置数据格式，再对数据范围自动列宽。';$basic.Cells.Item(25,2).Value2='备注列保留适当宽度，再换行并自动行高。';$basic.Cells.Item(26,2).Value2='金额右对齐，状态同时用文字与条件格式。';$basic.PageSetup.PrintArea='$A$1:$F$26'
  $window=$book.Windows.Item(1);$basic.Activate();$window.SplitRow=5;$window.FreezePanes=$true
  $chart=$chartSheet.ChartObjects().Add(0,[double]$chartSheet.Range('A12').Top,650,240);$chart.Chart.ChartType=51;$chart.Chart.SetSourceData($chartSheet.Range('B5:D9'));$chart.Chart.HasTitle=$true;$chart.Chart.ChartTitle.Text='预算与实际对比（万元）';$chart.Chart.ChartTitle.Font.Size=14;$chart.Chart.ChartArea.Font.Name='Microsoft YaHei';$chart.Chart.ChartArea.Font.Size=10;$chart.Chart.Legend.Position=-4107;$chart.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB=RGB '8E8C88';$chart.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB=RGB '44657D';$chart.Chart.Axes(2).TickLabels.NumberFormat='0';$chart.Chart.ApplyDataLabels();$chartSheet.PageSetup.PrintArea='$A$1:$H$26'
  $stage=Join-Path ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.tmp'))) ('excel-atlas-'+[guid]::NewGuid().ToString('N'));[IO.Directory]::CreateDirectory($stage)|Out-Null;Write-Output 'Excel: saving unique owned staging files';$book.SaveAs((Join-Path $stage 'atlas.xlsx'),51);$book.ExportAsFixedFormat(0,(Join-Path $stage 'atlas.pdf'));$totals=@($chartSheet.Range('C10').Value2,$chartSheet.Range('D10').Value2,$chartSheet.Range('E10').Value2);if(($totals -join ',') -ne '300,290,10'){throw 'Excel totals mismatch'};$counts=@{sheets=$book.Worksheets.Count;tables=2;charts=$chartSheet.ChartObjects().Count;totals=$totals;rows=4};$book.Close($false);$book=$null;[IO.File]::Copy((Join-Path $stage 'atlas.xlsx'),(Join-Path $root 'excel-base-styles.xlsx'),$true);[IO.File]::Copy((Join-Path $stage 'atlas.pdf'),(Join-Path $root 'excel-base-styles.pdf'),$true);$after=Snapshot $app.Workbooks
  Evidence 'excel-base-styles' @{id='excel-base-styles';host='Excel';scope='Native COM sample; not Office.js E2E';counts=$counts;sheets=@('基础样式','数据与图表');coverage=@('bold-title','auto-columns','wrap-text','auto-rows','alignment','number-formats','borders','native-table','filter','freeze','conditional-format','formulas','native-chart');objectScope=@{before=$before;after=$after;unchanged=(($before|ConvertTo-Json -Compress) -eq ($after|ConvertTo-Json -Compress))}}
  Write-Output 'Excel atlas saved: two sheets, one native chart'
 }finally{if($book){$book.Close($false)}}
}
if($Only -in @('All','PowerPoint')){
 $app=App 'PowerPoint.Application';$before=Snapshot $app.Presentations;$deck=$null
 try{
  $deck=$app.Presentations.Add(0);$deck.PageSetup.SlideWidth=960;$deck.PageSetup.SlideHeight=540;$script:slide=$null
  function Shape([int]$type,[double]$x,[double]$y,[double]$w,[double]$h,[string]$fill='EFEEEA'){$s=$script:slide.Shapes.AddShape($type,$x,$y,$w,$h);$s.Fill.ForeColor.RGB=RGB $fill;$s.Line.Visible=0;return $s}
  function Text([string]$text,[double]$x,[double]$y,[double]$w,[double]$h,[int]$size=18,[bool]$bold=$false,[string]$color='24282D'){$s=$script:slide.Shapes.AddTextbox(1,$x,$y,$w,$h);$s.Fill.Visible=0;$s.Line.Visible=0;$s.TextFrame.AutoSize=0;$s.TextFrame.WordWrap=-1;$s.TextFrame.MarginLeft=2;$s.TextFrame.MarginRight=2;$s.TextFrame.MarginTop=2;$s.TextFrame.MarginBottom=2;$r=$s.TextFrame.TextRange;$r.Text=$text;$r.Font.Name='Microsoft YaHei';$r.Font.Size=$size;$r.Font.Bold=if($bold){-1}else{0};$r.Font.Color.RGB=RGB $color;$r.ParagraphFormat.SpaceBefore=0;$r.ParagraphFormat.SpaceAfter=0;return $s}
  function Slide([string]$title,[string]$sub){$script:slide=$deck.Slides.Add($deck.Slides.Count+1,12);$null=Shape 1 0 0 960 540 'FAF9F6';$null=Text $title 42 34 876 48 30 $true;$null=Text $sub 44 96 870 50 18;$null=Text ('Office 基础样式 · DEMO / 示例数据 · '+$deck.Slides.Count) 44 480 870 24 12 $false '5E6267'}
  Slide '演示文稿基础样式范本' '一份文件覆盖常用元素与空间布局；文字、图形与表格保留编辑结构，图片单独嵌入。'
  $null=Text "01 文字层级与关键数字`n02 基础图形与原生表格`n03 数据图与图文布局" 65 185 420 180 24 $true;$null=Text "04 流程与时间线`n05 配色、间距与对齐`n06 结论与行动" 520 185 380 180 24 $true
  Slide '文字层级：标题、正文与关键数字' '标题30pt，正文18pt，标签16pt，表格及图表14pt，来源12pt。'
  $values=@('300','290','10');$labels=@('预算合计 / 万元','实际支出 / 万元','可用余额 / 万元');for($i=0;$i -lt 3;$i++){$x=55+$i*295;$null=Shape 1 $x 180 265 130;$null=Text $labels[$i] ($x+12) 192 240 30 16;$null=Text $values[$i] ($x+12) 235 240 60 36 $true};$null=Text '执行率96.7%。数字、单位、比较口径应一起出现；强调优先使用大小与字重。' 60 350 840 70 18
  Slide '基础图形：填充、描边与方向' '用可编辑的原生形状表达结构，避免将整页截图作为模板。'
  $types=@(1,5,9,4,33);$names=@('矩形','圆角矩形','圆形','菱形','箭头');for($i=0;$i -lt 5;$i++){$x=55+$i*177;$s=Shape $types[$i] $x 195 145 115;if($i -eq 2){$s.Width=115;$s.Left=$x+15;$s.Fill.Visible=0;$s.Line.Visible=-1;$s.Line.ForeColor.RGB=RGB '836344';$s.Line.Weight=2};if($i -eq 4){$s.Fill.ForeColor.RGB=RGB '836344'};$null=Text $names[$i] $x 325 145 32 16 $true};$null=Text '统一顶部与水平间距。浅底容纳正文，强调色用于方向与数据。' 60 400 840 42 18
  Slide '原生表格：表头、金额与汇总' '真正的PowerPoint表格对象；可直接编辑单元格、行列与对齐方式。'
  $tab=$slide.Shapes.AddTable(6,5,55,174,850,258);$matrix=@(@('类别','预算 / 万元','实际 / 万元','余额 / 万元','状态'),@('研发','120','114','6','按计划'),@('运营','80','84','-4','待复核'),@('营销','60','54','6','按计划'),@('支持','40','38','2','按计划'),@('合计','300','290','10','96.7%'))
  for($i=1;$i -le 6;$i++){for($j=1;$j -le 5;$j++){$cell=$tab.Table.Cell($i,$j).Shape;$cell.TextFrame.TextRange.Text=$matrix[$i-1][$j-1];$cell.TextFrame.TextRange.Font.Name='Microsoft YaHei';$cell.TextFrame.TextRange.Font.Size=14;$cell.TextFrame.TextRange.Font.Color.RGB=RGB '24282D';$cell.TextFrame.MarginLeft=8;$cell.TextFrame.MarginRight=8;$cell.TextFrame.TextRange.ParagraphFormat.Alignment=if($j -in 2,3,4){3}else{1};$cell.Fill.ForeColor.RGB=RGB $(if($i -in 1,6){'EFEEEA'}else{'FAF9F6'});if($i -in 1,6){$cell.TextFrame.TextRange.Font.Bold=-1}}}
  Slide '数据图：比较关系与可追溯标签' '此页为可编辑形状数据图；使用共同数据，单位为万元。'
  $baseline=405;$scale=1.65;for($i=0;$i -lt 4;$i++){$x=85+$i*172;$h1=$budgets[$i]*$scale;$h2=$actual[$i]*$scale;$null=Shape 1 $x ($baseline-$h1) 43 $h1 '8E8C88';$null=Shape 1 ($x+49) ($baseline-$h2) 43 $h2 '836344';$null=Text ([string]$budgets[$i]) $x ($baseline-$h1-30) 45 26 14;$null=Text ([string]$actual[$i]) ($x+49) ($baseline-$h2-30) 45 26 14;$null=Text $categories[$i] ($x-5) 413 115 30 16};$null=Text '中灰：预算    哑铜：实际    来源：本范本演示数据' 60 449 840 24 12 $false '5E6267'
  Slide '图文双栏：图片比例与文字留白' '图片为项目原创PNG，真正嵌入文件；配合简短说明与图注。'
  $pic=$slide.Shapes.AddPicture($image,0,-1,55,190,545,181.67);$pic.AlternativeText='规划、制作、复核三个阶段的原创插图';$null=Text "图像保持宽高比。`n`n右侧文字解释用途，避免重复图内标签。`n`n图片不能代替需要编辑的数据或表格。" 650 175 250 250 18;$null=Text '图1  内容规划与效果复核（原创示意图）' 55 390 550 30 14
  Slide '流程布局：节点与真实连接线' '四个节点等宽、对齐；箭头表达方向，不靠堆叠长段落。'
  $titles=@('明确目标','组织内容','应用样式','核验结果');$body=@('确认读者与用途','整理真实数据','选择对应元素','读回并查看页面');for($i=0;$i -lt 4;$i++){$x=55+$i*220;$null=Shape 5 $x 195 190 155;$null=Text $titles[$i] ($x+12) 215 166 35 20 $true;$null=Text $body[$i] ($x+12) 270 166 62 18;if($i -lt 3){$line=$slide.Shapes.AddConnector(1,$x+192,270,$x+218,270);$line.Line.ForeColor.RGB=RGB '836344';$line.Line.EndArrowheadStyle=3}};$null=Text '每个节点仅承载一个动作；条件、责任与例外使用补充文字说明。' 55 400 850 44 18
  Slide '时间线与三列布局' '同一网格容纳阶段、责任与产出；以统一基线组织阅读顺序。'
  $line=$slide.Shapes.AddConnector(1,180,196,780,196);$line.Line.ForeColor.RGB=RGB '836344';$line.Line.Weight=2
  $stages=@('准备阶段','实施阶段','复核阶段');$details=@("确定结构`n核对数据口径","应用元素`n保持样式一致","更新目录/公式`n检查原生输出");for($i=0;$i -lt 3;$i++){$x=55+$i*300;$null=Shape 9 ($x+115) 186 20 20 '836344';$label=Text $stages[$i] $x 225 250 40 20 $true;$label.TextFrame.TextRange.ParagraphFormat.Alignment=2;$null=Shape 1 $x 278 250 135;$null=Text $details[$i] ($x+12) 295 226 100 18}
  Slide '配色角色与对齐网格' '一套克制配色，通过底色、文字层级和间距形成稳定秩序。'
  $swatches=@('FAF9F6','24282D','836344','EFEEEA','5E6267','D7D4CE');$roles=@('页面背景','标题与正文','数据与标记','卡片与表头','辅助文字','线条与分隔');for($i=0;$i -lt 6;$i++){$x=55+($i%3)*300;$y=174+[int][math]::Floor($i/3)*130;$null=Shape 1 $x $y 265 90 $swatches[$i];$tc=if($i -in 1,2,4){'FFFFFF'}else{'24282D'};$null=Text ($roles[$i]+"`n#"+$swatches[$i]) ($x+12) ($y+12) 240 65 16 $true $tc}
  Slide '结论与行动：信息收束的版式' '先表达判断，再给出责任、动作与节点；来源留在固定位置。'
  $null=Text '统一元素样式，让内容结构更容易复用。' 60 180 840 65 28 $true;$null=Shape 1 60 275 835 130;$null=Text "责任：文档编制者与复核者`n动作：替换示例、保留语义结构、检查原生效果`n完成依据：数据正确、元素可编辑、页面无遮挡" 78 290 795 100 18;$null=Text '来源：本项目原创样式范本；演示数字不代表真实业务。' 60 438 840 30 12 $false '5E6267'
  $checks=@();foreach($sl in $deck.Slides){foreach($sh in $sl.Shapes){if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){$bh=$sh.TextFrame.TextRange.BoundHeight;$bw=$sh.TextFrame.TextRange.BoundWidth;if($bh -gt $sh.Height+1 -or $bw -gt $sh.Width+1){throw ('PPT text overflow: slide '+$sl.SlideIndex+' '+$sh.TextFrame.TextRange.Text)};$checks+=@{slide=$sl.SlideIndex;size=$sh.TextFrame.TextRange.Font.Size;boundHeight=$bh;boundWidth=$bw;box=@($sh.Left,$sh.Top,$sh.Width,$sh.Height)}}}}
  $pptStage=Join-Path ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.tmp'))) ('ppt-atlas-'+[guid]::NewGuid().ToString('N'));[IO.Directory]::CreateDirectory($pptStage)|Out-Null;$deck.SaveAs((Join-Path $pptStage 'atlas.pptx'),24);$deck.SaveAs((Join-Path $pptStage 'atlas.pdf'),32);foreach($sl in $deck.Slides){$sl.Export((Join-Path $root ('ppt-base-styles-slide-{0:d2}.png' -f $sl.SlideIndex)),'PNG',1600,900)};$deck.Saved=-1;$deck.Close();$deck=$null;[IO.File]::Copy((Join-Path $pptStage 'atlas.pptx'),(Join-Path $root 'ppt-base-styles.pptx'),$true);[IO.File]::Copy((Join-Path $pptStage 'atlas.pdf'),(Join-Path $root 'ppt-base-styles.pdf'),$true);$after=Snapshot $app.Presentations
  Evidence 'ppt-base-styles' @{id='ppt-base-styles';host='PowerPoint';scope='Native COM sample; not Office.js E2E';pdfPageCount=10;counts=@{slides=10;nativeTables=1;embeddedImages=1;shapeCharts=1};textBounds=$checks;coverage=@('cover','type-hierarchy','shapes','native-table','shape-chart','image','two-column','flow','connectors','timeline','three-column','palette','grid','conclusion');objectScope=@{before=$before;after=$after;unchanged=(($before|ConvertTo-Json -Compress) -eq ($after|ConvertTo-Json -Compress))}}
  Write-Output 'PowerPoint atlas saved: ten slides'
 }finally{if($deck){$deck.Saved=-1;$deck.Close()}}
}
