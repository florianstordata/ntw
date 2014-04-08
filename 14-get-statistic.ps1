# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[6]

$destination="D:\_Stordata\NTW\tigf"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\14-get-statistiques.txt") {remove-item "$sources\14-get-statistiques.txt"}

$ntwreport=".\mminfo.ax.output"

if ((test-path $ntwreport) -eq "true") {

Copy-Item $ntwreport -Destination "$sources\14-get-statistiques.txt"
}
else {
echo "pas de fichier $ntwreport"
}