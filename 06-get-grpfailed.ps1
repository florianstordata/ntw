# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[6]

$destination="D:\_Stordata\NTW\tigf"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\06-get-grpfailed*.xlsx") {remove-item "$sources\06-get-grpfailed*.xlsx"}

$ntwreport=".\STD.su.gs.output.csv"

if ((test-path $ntwreport) -eq "true") {


$grpfailed=import-csv $ntwreport -Header "Group Name", "Total Duration (Sec)", "Total Group Runs", "Successful", "Failed", "Interrupted", "Success Ratio (%)" | where {$_."Success Ratio (%)" -match '[0-9]'} #| where {$_."Success Ratio (%)" -ne "100"} |

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$($sources)\06-get-grpfailed.xlsx"
$i = 1 
foreach($results in $grpfailed) 
{
 $excel.cells.item($i,1) = $results."Group Name"
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