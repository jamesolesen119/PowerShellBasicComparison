#load intue.migration file

$FilePath = 'C:\Users\214jolesen.DVC-LOCAL\Documents\intune.migration - user and pc tracking.csv' #ensure this file exists in the location provided
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$wb = $xl.Workbooks.Open($FilePath)

#get data from column 6 of intune.migration
$intune_data = $wb.Worksheets['intune.migration - user and pc '].UsedRange.Rows.Columns[6].Value2

#close intune.migration file
$wb.close()

#open DevicesWithInventory file
$FilePath = 'C:\Users\214jolesen.DVC-LOCAL\Documents\DevicesWithInventory_703b5062-9111-4442-b8d3-d6dd1085a52e.csv' #ensure this file exists in the location provided
$wb = $xl.Workbooks.Open($FilePath)

#collect data from column 2 of DevicesWithInventory file
$database = $wb.Worksheets['DevicesWithInventory_703b5062-9'].UsedRange.Rows.Columns[2].Value2 #default name entered. Ensure that it has not been renamed.

#now that the data is collected, close DevicesWithInventory file
$wb.close()

#for each element in data1, search data2 to locate it
#if the element in data1 is not in data2, write it to the exceptions list.

$exceptionList = $null

for(($i = 0) ; ($i -lt $intune_data.Count) ; ($i++)) {

    #read through each element in database to compare
    $found = $false

    for(($x = 0) ; ($x -lt $database.Count) ; ($x++)){
        if ($intune_data[6, $i] -eq $database[2, $x]) {
            $found = $true
            break
        }# end if       
    }#end inner for loop
 
    if ($found -ne $true) {
        $exceptionList += $intune_data[6, $i] + " `n"
    }
}#end outer for loop
#write the exception list to a file
$exceptionList | Out-File -Append C:\Users\214jolesen.DVC-LOCAL\Documents\Missing.txt