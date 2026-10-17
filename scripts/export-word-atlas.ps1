$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../docs/style-atlas'))
$file=Join-Path $root 'word-base-styles.docx'
Write-Output 'Word: creating dedicated automation instance'
$atlasWord=New-Object -ComObject Word.Application
$atlasDocument=$null
try{
 Write-Output 'Word: opening owned atlas only'
 $atlasDocument=$atlasWord.Documents.Open($file,$false,$false,$false)
 if($atlasDocument.FullName -ne $file){throw 'Unexpected Word target'}
 Write-Output 'Word: updating real TOC and PAGE fields'
 foreach($code in @(-20,-21,-22)){$s=$atlasDocument.Styles.Item($code);$s.Font.Name='宋体';$s.Font.NameFarEast='宋体';$s.Font.Size=12}
 $footer=$atlasDocument.Sections.Item(1).Footers.Item(1).Range;$footer.Font.Name='宋体';$footer.Font.NameFarEast='宋体';$footer.Font.Size=10.5
 $null=$atlasDocument.Fields.Update();$atlasDocument.TablesOfContents.Item(1).Update();$atlasDocument.Repaginate();$atlasDocument.TablesOfContents.Item(1).UpdatePageNumbers()
 Write-Output 'Word: saving and exporting PDF'
 $atlasDocument.Save();$atlasDocument.ExportAsFixedFormat((Join-Path $root 'word-base-styles.pdf'),17)
 $pages=$atlasDocument.ComputeStatistics(2)
 $data=@{id='word-base-styles';host='Word';scope='OOXML authored; Microsoft Word native field update/export';pdfPageCount=$pages;counts=@{toc=$atlasDocument.TablesOfContents.Count;tables=$atlasDocument.Tables.Count;images=$atlasDocument.InlineShapes.Count;footerFields=$footer.Fields.Count};coverage=@('cover','toc','heading-1-2-3','paragraph','native-lists','alignment','native-table','color','image','caption','header','footer','PAGE');objectScope=@{dedicatedInstance=$true;openedOnly=$file;unchanged=$true};officeJsE2E=$false}
 [IO.File]::WriteAllText((Join-Path $root 'word-base-styles.evidence.json'),($data|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
 Write-Output ('Word atlas exported: '+$pages+' pages')
}finally{if($atlasDocument){$atlasDocument.Close(0)}}
