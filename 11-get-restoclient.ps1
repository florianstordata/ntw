# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[6]

$destination="D:\_Stordata\NTW\tigf"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\11-get-restoclient*.xlsx") {remove-item "$sources\11-get-restoclient*.xlsx"}

$ntwreport=".\STD.re.cs.output.csv"

if ((test-path $ntwreport) -eq "true") {


$restoclient=ipcsv -header "Client", "Amount of Data (B)", "Number of Files", "Total Requests", "Successful", "Failed" -path $ntwreport | where {$_.failed -match '[0-9]'}


$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$($sources)\11-get-restoclient.xlsx"
$i = 1 
foreach($results in $restoclient) 
{
 $excel.cells.item($i,1) = $results."Client"
 $excel.cells.item($i,2) = $results."Amount of Data (B)"
 $excel.cells.item($i,3) = $results."Number of Files"
 $excel.cells.item($i,4) = $results."Total Requests"
 $excel.cells.item($i,5) = $results."Successful"
 $excel.cells.item($i,6) = $results."Failed"

 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()
}
else {
echo "pas de fichier $ntwreport"
}