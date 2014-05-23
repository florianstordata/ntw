# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[5]

$destination="D:\_Stordata\traitement-ntw"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }

if(test-path "$sources\01b-get-licence.xlsx") {remove-item "$sources\01b-get-licence.xlsx"}

$ntwreport="nsradmin.hostid.output"

if ((test-path $ntwreport) -eq "true") {


$size=get-content $ntwreport  | where {$_ -match "TB License"}
$size=$size -replace 'TB',' TB'
$size=$size -replace '\s+', ';'
$size > $sources\size.txt
$sizes=Import-Csv $sources\size.txt -Delimiter ";" -header vide, nom, networker, source, capacity, taille, unit, license


$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$sources\01b-get-licence.xlsx"
$i = 1 
foreach($results in $sizes) 
{

 $excel.cells.item($i,1) = $results.capacity
 $excel.cells.item($i,2) = $results.taille
 $excel.cells.item($i,3) = $results.unit
 

 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()

Remove-Item $sources\size.txt
}
else {
echo "pas de fichier $ntwreport"
}