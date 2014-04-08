
$destination="D:\_Stordata\NTW\tigf"
$sources="$destination\sources"
if(test-path $sources) {remove-item $sources -Recurse}

#recuperation du repertoire de scripts 
$scripts="D:\_Stordata\NTW\Scripts"
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

$traitement="D:\_Stordata\NTW\tigf\traitements"
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

$item=(get-item $traitement\*).Name

foreach($dir in $item){
#echo $scripts
$date1= get-date
echo "."
echo ".."
echo "debut du traitement du repertoire : $dir"
pushd "$traitement\$dir"

echo "traitement du script 01-get-model.ps1        (1/14)"
& $scripts\01-get-model.ps1

echo "traitement du script 01b-get-licence.ps1     (1b/14)"
& $scripts\01b-get-licence.ps1

echo "traitement du script 02-recap-bckp.ps1       (2/14)"
& $scripts\02-recap-bckp.ps1

echo "traitement du script 03-recap-resto.ps1      (3/14)"
& $scripts\03-recap-resto.ps1

echo "traitement du script 04-get-clean.ps1        (4/14)"
& $scripts\04-get-clean.ps1

echo "traitement du script 05-get-stat.ps1         (5/14)"
& $scripts\05-get-stat.ps1

echo "traitement du script 06-get-grpfailed.ps1    (6/14)"
& $scripts\06-get-grpfailed.ps1

echo "traitement du script 07-get-ssfailed.ps1     (7/14)"
& $scripts\07-get-ssfailed.ps1

echo "traitement du script 08-get-robot.ps1        (8/14)"
& $scripts\08-get-robot.ps1

echo "traitement du script 09-bandes.ps1           (9/14)"
& $scripts\09-bandes.ps1

echo "traitement du script 10-get-disk.ps1        (10/14)"
& $scripts\10-get-disk.ps1

echo "traitement du script 11-get-restoclient.ps1 (11/14)"
& $scripts\11-get-restoclient.ps1

echo "traitement du script 12-get-bootstrap.ps1   (12/14)"
& $scripts\12-get-bootstrap.ps1

echo "traitement du script 13-get-index.ps1       (13/14)"
& $scripts\13-get-index.ps1

echo "traitement du script 13-get-index.ps1       (14/14)"
& $scripts\14-get-statistic.ps1

$elapsed1=[math]::round(((Get-Date) - $Date1).TotalMinutes,2)


echo "Fin du traitement du repertoire : $dir en $elapsed1"
echo ".."
echo "."

}

$elapsed=[math]::round(((Get-Date) - $Date).TotalMinutes,2)
echo "This report took $elapsed minutes to run all scripts."
Read-Host "press Enter key to continue"