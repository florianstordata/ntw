# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[5]

$destination="D:\_Stordata\traitement-ntw"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\06-get-vol*.xlsx") {remove-item "$sources\06-get-vol*.xlsx"}

$ntwreport=".\mminfo.disk.output"

if ((test-path $ntwreport) -eq "true") {

$stat=ipcsv $ntwreport | where {$_.volume -notmatch '.RO'}

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$($sources)\06-get-vol.xlsx"
$i = 1 
foreach($results in $stat) 
{
 $excel.cells.item($i,1) = $results."volume"
 $excel.cells.item($i,2) = $results."pool"
 $excel.cells.item($i,3) = $results."Úcrit"
 $excel.cells.item($i,4) = $results."accÞs-vol"
 $excel.cells.item($i,5) = $results."entitÚs_de_sauvegarde"

 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()
}
else {
echo "pas de fichier $ntwreport"
}