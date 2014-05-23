# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[5]

$destination="D:\_Stordata\traitement-ntw"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
if(test-path "$sources\04-get-clean*.txt") {remove-item "$sources\04-get-clean*.txt"}

$ntwreport=".\nsrjb.output"

if ((test-path $ntwreport) -eq "true") {


Get-Content $ntwreport | where {$_ -match "jukebox" -or $_ -match "clean*"} > $sources\04-get-clean.txt

}
else {
echo "pas de fichier $ntwreport"
}