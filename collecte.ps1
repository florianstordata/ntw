# SERVICE MANAGES - Copyright SOTRDATA 2014
# Script de Collecte d'infos 
# client : 
#
#creation d'un repertoire contenant les fichiers

$date=Get-Date -Format ddMMyyyhhmmss
$rep="$($env:computername)_$date"
New-Item -ItemType Directory $rep


# \***** VARIABLES DU CLIENT
$zip=".\7z\x64\7z.exe"
$user="user"
$passwd="passwd"
$nsrpath="E:\Legato\nsr"
$nmcpath="E:\Legato\Management\GST\bin"
$report="E:\Legato\Management\GST\bin\gstclreport.bat"
$log="$nsrpath\stordata\$rep\collecte.log"


$expediteur="client@domain.out"
$smtp="smtp.domain.local"
$destinataire="mail@domain.local"
# VARIABLES DU CLIENT *****\


ECHO "Debut de la collecte" >> $log
Get-date -format g >> $log

# *****  NOM DU FICHIER 
# creation d'un nom de fichier unique base sur la date et l'heure
# T prend comme valeur l'heure sous ce format : 20:20:33.10 hh:mm:ss.dd ou hh:mm:ss,dd suivant les versions
$date=Get-Date -Format ddMMyyyhhmmss
$FICH="$($env:computername)_$date.zip"


# ***** COLLECTE DES INFOS 
# les fichiers ont tous l'extension .output

# infos sur les disques
Get-WmiObject Win32_LogicalDisk | Export-Csv -path $rep\disk.output.csv -NoTypeInformation

# nom de la machine
$env:computername > $rep\model.output

# constructeur et modele
$model=Get-WmiObject -Class Win32_ComputerSystem
echo "$($model.Manufacturer) $($model.model)" >> $rep\model.output

# type d'OS
(Get-WmiObject Win32_OperatingSystem).name >> $rep\model.output

# date de dernier demarrage
echo "dernier reboot" >> $rep\model.output
$reboot = Get-WmiObject -Class Win32_OperatingSystem
$reboot.ConvertToDateTime($reboot.LastBootUpTime) >> $rep\model.output
echo "  " >> $rep\model.output

# quantite de memoire
Get-WmiObject -Class Win32_CacheMemory -Namespace root/cimv2 -ComputerName . | Export-Csv -path $rep\memory.output.csv -NoTypeInformation


Get-WmiObject Win32_Processor | Export-Csv -path $rep\proc.output.csv -NoTypeInformation

# repertoires d'installation de NW et NMC 
echo "repertoire de NW : $nsrpath" >> $rep\model.output
echo "repertoire de NMC : $nmcpath" >> $rep\model.output

# systeminfo > systeminfo.output
# propositions 
# Get-WmiObject CIM_TemperatureSensor
#  SystemType    ThermalState  TotalPhysicalMemory


# recuperation de la totalite de la configuration
nsradmin -i total.input > $rep\nsradmin.total.output

# recuperation de la version de NW, du hostid et des licences
nsradmin -i hostid.input > $rep\nsradmin.hostid.output

# liste des repertoires de la base des indexes avec leurs tailles
nsrls | Select-String records > $rep\nsrls.output

mminfo -aX > $rep\mminfo.ax.output

mminfo -B "-xc," > $rep\mminfo.b.output

ECHO "creation des reports" >> $log

# reports sur les status des sauvegardes
& $report -u $user -P $passwd -r STD.su.ssdc.failed -f $rep\STD.su.ssdc.failed.output -x csv

& $report -u $user -P $passwd -r STD.su.gs -f $rep\STD.su.gs.output -x csv

& $report -u $user -P $passwd -r STD.su.servsum -f $rep\STD.su.servsum.output -x csv

# reports sur les statistics des sauvegardes
& $report -u $user -P $passwd -r STD.sa.gs.full -f $rep\STD.sa.gs.full.output -x csv

& $report -u $user -P $passwd -r STD.sa.ms -f $rep\STD.sa.ms.output -x csv

# reports sur les restaurations
& $report -u $user -P $passwd -r STD.re.cs -f $rep\STD.re.cs.output -x csv

& $report -u $user -P $passwd -r STD.re.servsum -f $rep\STD.re.servsum.output -x csv

# reports sur les clones
& $report -u $user -P $passwd -r STD.cl.cd -f $rep\STD.cl.cd.output -x csv

# ***** MOYENS DE SAUVEGARDE
# En cas de DataDomain, décommenter la ligne apres avoir creer le report
# reports sur les data domain
# & $report -u $user -P $passwd -r STD.dd.ds -f $rep\STD.dd.ds.output -x csv

# En cas de robotique
# Si il y a un seul robot, utiliser le nsrjb sans preciser le nom du robot en cas de changement de nom
$robot1="HP@10.0.0"
nsrjb -v > $rep\nsrjb.output
inquire > $rep\inquire.output

mminfo -a -q "location=$robot1" -r "volume,state,pool,written,volaccess,volretent,savesets,location" "-xc," > $rep\mminfo.amloc.output



# # ************ ENVOI DES INFOS PAR MAIL
# # 

ECHO "creation du zip et envoi du mail" >> $log
If (Test-Path $FICH) { remove-item $FICH }
& $zip a "$FICH" ".\$rep\*"

send-mailmessage -to "$destinataire" -from "$expediteur" -SmtpServer "$smtp" -subject "Networker services manag�s"  -Attachments $FICH
# send-mailmessage -to "$destinataire" -from "$expediteur" -SmtpServer "$smtp" -subject "Networker services managés" 

# Gestion des fichiers de donnees, on supprime ceux qui sont plus vieux que 22 jours
get-childitem *.zip | ?{!$_.PSIsContainer -and ($_.CreationTime -lt (get-Date).adddays(-22))}|remove-item



ECHO "FIN de la collecte" >> $log
Get-date -format g >> $log
ECHO "********************" >> $log