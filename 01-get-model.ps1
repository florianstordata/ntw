    # on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[5]

$destination="D:\_Stordata\traitement-ntw"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\01-get-model*.xlsx") {remove-item "$sources\01-get-model*.xlsx"}

$ntwreport=".\model.output"

if ((test-path $ntwreport) -eq "true") {

$table=@(Get-Content $ntwreport)

echo "hostname;$($table[0])">$sources\model.csv
echo "model;$($table[1])">>$sources\model.csv
echo "os;$($table[2])">>$sources\model.csv
echo "last-boot;$($table[5])">>$sources\model.csv

$hostid=get-content nsradmin.hostid.output | where {$_ -match "host id"} | select-object -First 1
$hostid -replace ':',';' >>model.csv
$build=get-content nsradmin.hostid.output | where {$_ -match "build"} | select-object -First 1
$vers=$build.Split("NetWorker")

echo "build;$($vers[9])">>$sources\model.csv


$result=import-csv -header index, valeur  -delimiter ";" $sources\model.csv

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$($sources)\01-get-model.xlsx"
$i = 1 
foreach($results in $result) 
{
$bash=$results.valeur.Split("|")

$results.valeur=$bash[0
]
 $excel.cells.item($i,1) = $results.index
 $excel.cells.item($i,2) = $results.valeur
 

 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()

Remove-Item $sources\model.csv
}
else {
echo "pas de fichier $ntwreport"
}