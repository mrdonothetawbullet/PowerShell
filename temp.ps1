$file1 = 'C:\Users\eric\Documents\Book1.xlsx' # source's fullpath
$file2 = 'C:\Users\eric\Documents\Book2.xlsx' # destination's fullpath
$xl = new-object -c excel.application
$xl.displayAlerts = $false # don't prompt the user
$wb2 = $xl.workbooks.open($file1, $null, $true) # open source, readonly
$wb1 = $xl.workbooks.open($file2) # open target
$sh1_wb1 = $wb1.sheets.item(2) # second sheet in destination workbook
$sheetToCopy = $wb2.sheets.item('Sheet3') # source sheet to copy
$sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook
$wb2.close($false) # close source workbook w/o saving
$wb1.close($true) # close and save destination workbook
$xl.quit()
