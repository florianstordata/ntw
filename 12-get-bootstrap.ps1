# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[5]

$destination="D:\_Stordata\traitement-ntw"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\12-get-bootstrap.xlsx") {remove-item "$sources\12-get-bootstrap.xlsx"}

$ntwreport=".\mminfo.b.output"

if ((test-path $ntwreport) -eq "true") {


$bootstrap=Import-Csv $ntwreport -header "date-time", "level", "ssid", "media-file", "record", "volume" 


$20bootstrap= $bootstrap | select-object -last 20


$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$sources\12-get-bootstrap.xlsx"
$i = 1 
foreach($1bootstrap in $20bootstrap) 
{

 $excel.cells.item($i,1) = $1bootstrap."date-time"
 $excel.cells.item($i,2) = $1bootstrap."level"
 $excel.cells.item($i,3) = $1bootstrap."ssid"
 $excel.cells.item($i,4) = $1bootstrap."media-file"
 $excel.cells.item($i,5) = $1bootstrap."record"
 $excel.cells.item($i,6) = $1bootstrap."volume"

 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()

}
else {
echo "pas de fichier $ntwreport"
}