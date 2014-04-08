# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[6]

$destination="D:\_Stordata\NTW\tigf"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory -Name $sources }
if(test-path "$sources\10-get-disk.xlsx") {remove-item "$sources\10-get-disk.xlsx"}

$ntwreport=".\disk.output.csv"

if ((test-path $ntwreport) -eq "true") {


#$disk=(ipcsv .\disk.output.csv).freespace

#$disk -replace '\s+', ',' >$sources\disk.csv

$result=Import-Csv $ntwreport | where {$_.freespace -ne "$null"}


$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$sources\10-get-disk.xlsx"
$i = 1 
foreach($results in $result) 
{

 $excel.cells.item($i,1) = $results.DeviceID
 $excel.cells.item($i,2) = $results."VolumeName"
 $excel.cells.item($i,3) = ($results.Size / 1GB)
 $excel.cells.item($i,4) = ($results.FreeSpace / 1GB)
 
 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()
}
else {
echo "pas de fichier $ntwreport"
}