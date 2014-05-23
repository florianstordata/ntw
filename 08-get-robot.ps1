# on recupere le repertoire courant
$rep = (Get-Location).path
$server=$rep.split("\\""_")[5]

$destination="D:\_Stordata\traitement-ntw"
$sources="$destination\sources\$server"

If (-not (Test-Path $sources)) { New-Item -ItemType Directory $sources }
If (-not (Test-Path $sources\junk)) { New-Item -ItemType Directory $sources\junk }

if(test-path "$sources\08-get-robots*.xlsx") {remove-item "$sources\08-get-robots*.xlsx"}

if ((Get-content  .\nsradmin.total.output | where {$_ -match "type: NSR Jukebox"})) {
# split du fichier nsradmin total
$file = (GC .\nsradmin.total.output)
$i=1
$newfile="$sources\junk\hidden.log"
ForEach ($line in $file) {
  If ($line -match "type: NSR") {
   
  $newfile = "$sources\junk\$($line.Split(':')[1])$i.log"
  $i++
  # echo $newfile
  }

  Else {
    $line | Out-File -Append $newfile
    }
     }


$juke = Get-Item $sources\junk\*jukebox*.log
$j=1
foreach($jukes in $juke) 
{

$filtemp="$sources\temp.csv"
[String] $S = [io.file]::ReadAllText($jukes)
$S2=$s -replace "\\\r\n",""
$s2 | set-content "$sources\temp.txt"

$temp=Get-Content $sources\temp.txt
$temp=$temp -replace '\s+', ' '
$temp > $filtemp


# contructeur
$names=import-csv $filtemp -delimiter ":" -header type, valeur | where {$_.type -eq "name"}
$name = $names.valeur.Split('@')[0]

# modele
$serials=import-csv $filtemp -delimiter ":" -header type, valeur | where {$_.type -eq "hardware id"}
$model = $serials.valeur.Split(' ')[1]

#serial number
$serial = $serials.valeur.Split(' ')[4]

echo "tabname:$name" >> $filtemp
echo "tabmodel:$model" >> $filtemp
echo "tabserial:$serial" >> $filtemp



# number of Slots 
$info=import-csv $filtemp -delimiter ":" -header type, valeur | where {
$_.type -eq "tabname" -or
$_.type -eq "tabmodel" -or
$_.type -eq "tabserial" -or
$_.type -eq "physical slots" -or 
$_.type -eq "number drives" -or 
$_.type -eq "name" -or
$_.type -eq "load sleep" -or
$_.type -eq "unload sleep" -or
$_.type -eq "eject sleep" -or
$_.type -eq "idle device timeout" -or
$_.type -eq "auto clean" 
}



$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$sources\08-get-robots$j.xlsx"
$i = 1 
foreach($results in $info) 
{

 $excel.cells.item($i,1) = $results.type
 $excel.cells.item($i,2) = $results.valeur
 
 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()
$j++
}

remove-item $sources\temp.csv
Remove-Item $sources\temp.txt}

else { echo "No Jukebox Available" }
