# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[6]

$destination="D:\_Stordata\NTW\tigf"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\13-get-index*.xlsx") {remove-item "$sources\13-get-index*.xlsx"}

$ntwreport=".\nsrls.output"

if ((test-path $ntwreport) -eq "true") {

$content=get-content $ntwreport | where {$_ -notlike "*actuellement*" -and $_ -notlike "*currently*" -and $_ -notlike $null}

$content -replace "Program Files", "Program_Files" > $sources\nsrls.txt

$file = Get-Content $sources\nsrls.txt
$containsWord = $file | %{$_ -match "enregistrement"}
If($containsWord -contains $true)
{
$ntw=Import-Csv $sources\\nsrls.txt -header "index", "point", "nbrecord", "enregistrements", "necessitant", "taille", "unite" -delimiter " " | Sort-Object {[int] $_.nbrecord} -Descending
}
else {
$ntw=Import-Csv $sources\\nsrls.txt -header "index", "nbrecord", "enregistrements", "necessitant", "taille", "unite" -delimiter " " | Sort-Object {[int] $_.nbrecord} -Descending

}


#$ntw | select-object -First 4 | Export-csv $sources\13-index-first.txt -NoTypeInformation

#$ntw | Export-csv index-all.csv -NoTypeInformation

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$sources\13-get-index.xlsx"
$i = 1 
foreach($ntwid in $ntw) 
{

 $excel.cells.item($i,1) = $ntwid.index
 $excel.cells.item($i,2) = $ntwid."nbrecord"
 $excel.cells.item($i,3) = $ntwid."taille"
 $excel.cells.item($i,4) = $ntwid."unite"
 
 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()

Remove-Item $sources\nsrls.txt}

else {
echo "pas de fichier $ntwreport"
}