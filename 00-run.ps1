

$asupv2="C:\ASUPV2"
$7z="$asupv2\7za.exe"

$destination="D:\_Stordata\traitement-ntw"
$sources="$destination\sources"
if(test-path $sources) {remove-item $sources -Recurse}

#recuperation du repertoire de scripts 
$scripts="D:\_Stordata\_Scripts\Ntw"
# $scripts=Read-Host "specifier le repertoire contenant les scripts"
# if ($scripts -eq "") {
# echo "aucun repertoire specifié"
# Read-Host "press enter key to exit"
# return
# }
# else {
# echo "le repertoire : $scripts, sera utilisé "
# }
#definition du repertoire de traitement

$traitement="D:\_Stordata\traitement-ntw\traitements"
# $traitement=Read-Host "specifier le repertoire de traitements"
# if ($traitement -eq "") {
# echo "aucun repertoire specifié"
# Read-Host "press enter key to exit"
# return
# }
# else {
# echo "le repertoire : $traitement, sera utilisé "
# }

#recuperation de la date avant l'execution du scripts pour connaitre le temps de traitements
$Date = Get-Date


do {

$archives=Get-Item -Path $traitement\* -Include *.zip

foreach ($archive in $archives )
{
$base=$archive.basename
& "$7z" x "$archive" -o"$traitement\$base" -aoa
Remove-Item "$archive"
}}
while ($flagArchive=(Test-Path -Path $traitement\* -Include *.zip))



$item=(get-item $traitement\*).Name

foreach($dir in $item){
#echo $scripts
$date1= get-date
echo "."
echo ".."
echo "debut du traitement du repertoire : $dir"
pushd "$traitement\$dir"

echo "traitement du script 01-get-model.ps1        (1/15)"
& $scripts\01-get-model.ps1

echo "traitement du script 01b-get-licence.ps1     (1b/15)"
& $scripts\01b-get-licence.ps1

echo "traitement du script 02-recap-bckp.ps1       (2/15)"
& $scripts\02-recap-bckp.ps1

echo "traitement du script 03-recap-resto.ps1      (3/15)"
& $scripts\03-recap-resto.ps1

echo "traitement du script 04-get-clean.ps1        (4/15)"
& $scripts\04-get-clean.ps1

echo "traitement du script 05-get-stat.ps1         (5/15)"
& $scripts\05-get-stat.ps1

echo "traitement du script 06-get-grpfailed.ps1    (6/15)"
& $scripts\06-get-grpfailed.ps1

echo "traitement du script 07-get-ssfailed.ps1     (7/15)"
& $scripts\07-get-ssfailed.ps1

echo "traitement du script 08-get-robot.ps1        (8/15)"
& $scripts\08-get-robot.ps1

echo "traitement du script 09-bandes.ps1           (9/15)"
& $scripts\09-bandes.ps1

echo "traitement du script 10-get-disk.ps1        (10/15)"
& $scripts\10-get-disk.ps1

echo "traitement du script 11-get-restoclient.ps1 (11/15)"
& $scripts\11-get-restoclient.ps1

echo "traitement du script 12-get-bootstrap.ps1   (12/15)"
& $scripts\12-get-bootstrap.ps1

echo "traitement du script 13-get-index.ps1       (13/15)"
& $scripts\13-get-index.ps1

echo "traitement du script 14-get-statistic.ps1   (14/15)"
& $scripts\14-get-statistic.ps1

echo "traitement du script 15-get-volume.ps1      (15/15)"
& $scripts\15-get-volume.ps1


$elapsed1=[math]::round(((Get-Date) - $Date1).TotalMinutes,2)


echo "Fin du traitement du repertoire : $dir en $elapsed1"
echo ".."
echo "."

}

$elapsed=[math]::round(((Get-Date) - $Date).TotalMinutes,2)
echo "This report took $elapsed minutes to run all scripts."
Read-Host "press Enter key to continue"