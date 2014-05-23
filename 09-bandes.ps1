# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[5]

$destination="D:\_Stordata\traitement-ntw"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory -Name $sources }
if(test-path "$sources\09-get-media*.xlsx") {remove-item "$sources\09-get-media*.xlsx"}

$ntwreport="mminfo.amloc*"

if ((test-path $ntwreport) -eq "true") {

$media= (get-item $ntwreport)
$j=1
foreach($medias in $media) 
{

$result=Import-Csv $medias -header volume, volstate, pool, written, vol-access, expires, savesets, location |where {$_.savesets -match "[0-9]"}

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$sources\09-get-media$j.xlsx"
$i = 1 
foreach($results in $result) 
{

 $excel.cells.item($i,1) = $results.volume
 $excel.cells.item($i,2) = $results."volstate"
 $excel.cells.item($i,3) = $results."pool"
 $excel.cells.item($i,4) = $results.written
 $excel.cells.item($i,5) = $results."vol-access"
 $excel.cells.item($i,6) = $results.expires
 $excel.cells.item($i,7) = $results.savesets
 $excel.cells.item($i,8) = $results.location
 
 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()
$j++
}
}
else {
echo "pas de fichier $ntwreport"
}