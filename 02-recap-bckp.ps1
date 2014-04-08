# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[6]

$destination="D:\_Stordata\NTW\tigf"
$sources="$destination\sources\$server"
$ntwreport=".\STD.su.servsum.output.csv"


If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\02-get-backup*.xlsx") {remove-item "$sources\02-get-backup*.xlsx"}

if ((test-path "$ntwreport") -eq "true") {


$backup=import-csv $ntwreport  -Header "Server Name", "Total Duration (Sec)", "Total Group Runs", "Successful", "Failed", "Interrupted", "Success Ratio (%)" | where {$_.Failed -match '[0-9]'}

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$($sources)\02-get-backup.xlsx"
$i = 1 
foreach($results in $backup) 
{
 $excel.cells.item($i,1) = $results."Server Name"
 $excel.cells.item($i,2) = $results."Total Duration (Sec)" 
 $excel.cells.item($i,3) = $results."Total Group Runs"
 $excel.cells.item($i,4) = $results."Successful"
 $excel.cells.item($i,5) = $results."Failed"
 $excel.cells.item($i,6) = $results."Interrupted"
 $excel.cells.item($i,7) = $results."Success Ratio (%)"

 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()
}
else {
echo "pas de fichier $ntwreport"
}