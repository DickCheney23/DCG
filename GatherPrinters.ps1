$strFilter = "(objectCategory=printQueue)"

$objDomain = New-Object System.DirectoryServices.DirectoryEntry

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.Filter = $strFilter

$colResults = $objSearcher.FindAll()

$objPrinters = @()
foreach ($objResult in $colResults)
 {$objItem = $objResult.Properties;
 $objptr = New-Object system.object
 $objptr | Add-Member -type noteproperty -name Name -value $objItem.printername
 $objptr | Add-Member -type noteproperty -name Location -value $objItem.location
 $objptr | Add-Member -type noteproperty -name Driver -value $objItem.drivername
 $objptr | Add-Member -type noteproperty -name Server -value $objItem.servername
 $objptr | Add-Member -type noteproperty -name Description -value $objItem.description
 $objPrinters += $objptr
 }

$objPrinters
#Export the data to a CSV file
$objPrinters | export-csv Printers.csv -NTI
$objPrinters | ft name, driver, server, description


[pscustomobject]@{
    ExtraInfo = (@(1,3,5,6) | Out-String).Trim()
} | Export-Csv -notype Random.csv

$Top5 = Get-Process -Computer $Item | Sort Handles -descending |Select -expand Handles | |Select -First 5 
 $new = [pscustomobject]@{ Top5 = (@($Top5) -join ',')
 }

[pscustomobject]@{
    First = 'Boe'
    Last = 'Prox'
    ExtraInfo = (@(1,3,5,6) | Out-String).Trim()
    State = 'NE'
    IPs = (@('111.222.11.22','55.12.89.125','125.48.2.1','145.23.15.89','123.12.1.0') | Out-String).Trim()
} | Export-Csv -notype Random.csv
 
$Printers = [pscustomobject]@{
	objPrinters = (@($objPrinters) -join ',')
	}

$OutArray |Select-Object Name,@{Expression={$_.objPrinters.Name -join ';'}} #|Export-Csv test.csv
$OutArray

$Printers

$NumColsToExport = 4


#Create array to hold the data to be sent to the CSV file
$Printers=@()
$pNames=@("Name", "Driver", "Server", "Description")
ForEach ($Row in $objPrinters){
    $obj = New-Object PSObject
    For ($i=0;$i -lt $NumColsToExport; $i++){
              $obj | Add-Member -MemberType NoteProperty -Name $pNames[$i] -Value $Row[$i]
      Echo $i
      Echo $Row[$i]
      }
   $Printers+=$obj
   $obj=$Null
}

#$Printers

#Export the data to a CSV file
$Printers | export-csv Printers.csv -NTI

$Printers | ft name, driver, server, description