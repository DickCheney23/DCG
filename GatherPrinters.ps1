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

#Show the array
$objPrinters

#Export the data to a CSV file
$objPrinters | export-csv Printers.csv -NTI

#View the data in a table
$objPrinters | ft name, driver, server, description
