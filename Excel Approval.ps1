Clear-Host
Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
Remove-Variable * -ErrorAction SilentlyContinue; $Error.Clear();
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$datetime = Get-Date -Format G
$dt = (Get-Date).ToString("ddMMyyyy_HHmmss")
Import-Module -Name slpslib

########
#BATCH3#
########
#region Function
Function Cleanup
{
    Write-Host "Cleaning up files/directory : " -NoNewline
    Try
    {
        Get-ChildItem -Path "$ScriptDir" -Recurse -Exclude "format it.ps1" | Remove-Item -Recurse -Force
        Write-Host "Success" -ForegroundColor Green
    }

    Catch 
    {
        Write-Warning ($_); Continue
    }
}

Function Cleanup2
{
    Write-Host "Cleaning up : " -NoNewline
    Try
    {
        Get-ChildItem -Path "$ScriptDir" -Recurse -Exclude "format it.ps1","Thursday","Saturday","Sunday","Batch3.xlsx" | Remove-Item -Recurse -Force
        Write-Host "Success" -ForegroundColor Green
    }

    Catch 
    {
        Write-Warning ($_); Continue
    }
}
#endregion Function

#region Cleanup
Cleanup
#endregion Cleanup

#region Copy MPL to current directory
$Mpl = "\\kulbscvmfs03\IT\30 Knowledgebase\30 Infrastructure\10 Knowledge Base\10 Server\14 DCM Patching\Patching Team\01 Master Patching List\MasterPatchingList_v0.9.xlsx"
Write-Host "Coying MPL to current directory : " -NoNewline
Try
{
    Copy-Item -Path $Mpl -Destination "$ScriptDir\" -Force -ErrorAction Stop | Out-Null
    Write-Host "Success" -ForegroundColor Green
}

Catch 
{
    Write-Warning ($_); Continue
}

Finally
{
    $Error.Clear()
}
#endregion Copy MPL to current directory

#region Create Temp_Batch3.csv, Thursday, Saturday and Sunday directory
Write-Host "Creating Temp_Batch3.csv, Thursday, Saturday and Sunday directory : " -NoNewline
Start-Sleep 2
Try
{
    $ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent -ErrorAction Stop
    New-SLDocument -WorkbookName Temp_Batch3 -WorksheetName Overwritten -Path "$ScriptDir" -Force -ErrorAction Stop
    New-Item -ItemType Directory -Path "$ScriptDir\Thursday\" -Force -ErrorAction Stop | Out-Null
    New-Item -ItemType Directory -Path "$ScriptDir\Saturday\" -Force -ErrorAction Stop | Out-Null
    New-Item -ItemType Directory -Path "$ScriptDir\Sunday\" -Force -ErrorAction Stop | Out-Null
    Write-Host "Success" -ForegroundColor Green

}

Catch
{
    Write-Warning ($_)
    Continue
}

Finally
{
    $Error.Clear()
}
#endregion Create Temp_Batch3.csv, Thursday, Saturday and Sunday directory

#region Copy PatchingList sheet from MPL to Temp_Batch3.xlsx
Write-Host "Copying PatchingList sheet from MPL to Temp_Batch3.xlsx : " -NoNewline
Start-Sleep 2
Try
{
    $File1 = "$ScriptDir\MasterPatchingList_v0.9.xlsx"
    $File2 = "$ScriptDir\Temp_Batch3.xlsx"
    $xl = new-object -c excel.application -ErrorAction Stop
    $xl.displayAlerts = $false # don't prompt the user
    $wb2 = $xl.workbooks.open($file1, $null, $true) # open source, readonly
    $wb1 = $xl.workbooks.open($file2)  # open target
    $sh1_wb1 = $wb1.sheets.item(1)  # second sheet in destination workbook
    $sheetToCopy = $wb2.sheets.item('PatchingList') # source sheet to copy
    $sheetToCopy.copy($sh1_wb1)  # copy source sheet to destination workbook
    $wb2.close($false)  # close source workbook w/o saving
    $wb1.close($true)  # close and save destination workbook
    $xl.quit()
    Write-Host "Success" -ForegroundColor Green
}

Catch
{
    Write-Warning ($_)
    Continue
}

Finally
{
    $Error.Clear()
}
#spps -n excel 
#endregion Copy PatchingList sheet from MPL to Temp_Batch3.xlsx

#region Temp_Batch3.xlsx to Temp_Batch3.csv with delimiter
Write-Host "Converting Temp_Batch3.xlsx to Temp_Batch3.csv with delimiter : " -NoNewline
Start-Sleep 2
Try
{
    $Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $Excel.visible = $false
    $Excel.DisplayAlerts = $False
    $file = "$ScriptDir\Temp_Batch3.xlsx"
    $csv = "$ScriptDir\Batch3.csv"
    $csvfinal = "$ScriptDir\New_Batch3.csv"
    $workfile = $Excel.Workbooks.open($file)
    $Sheet = $workfile.Worksheets.Item(1)
    $Sheet.Activate()
    $Sheet.SaveAs($csv,[Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSVWindows)
    $workfile.Close()
    sleep 3
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Import-Csv $csv | Export-Csv $csvfinal -Delimiter ";" -NoTypeInformation -ErrorAction Stop | Out-Null
    Write-Host "Success" -ForegroundColor Green
}

Catch
{
    Write-Warning ($_)
    Continue
}

Finally
{
    $Error.Clear()
}
#endregion Convert Temp_Batch3.xlsx to Temp_Batch3.csv with delimiter

#THURSDAY
#region Filter the Batch-Thursday column
Write-Host "Filtering the Batch-Thursday column : " -NoNewline
Start-Sleep 2
$File1 = "C:\Users\suhail_asrulsani-ops\Desktop\Excel\Batch3.csv"
$filtered = Import-Csv -Path $File1 | Where-Object { $_."New Batch" -eq "Batch3-Thurs" }
Try { New-Item -ItemType "file" -Path "$ScriptDir\Thursday.csv" -Force -ErrorAction Stop | Out-Null } Catch{}
$file2 = "$ScriptDir\Temp_Thursday.csv"
$filtered | Export-Csv -Path $File2 -NoTypeInformation
Write-Host "Success" -ForegroundColor Green
#endregion Filter the Batch-Thursday column

#region Remove unnecessary columns
Write-Host "Remove unnecessary columns : " -NoNewline
Start-Sleep 2
Import-Csv -Path "$ScriptDir\Temp_Thursday.csv" | Select-Object "ServerName","IP Address","Application","Application Owner","Domain","Site","Region","DCM Group","Remark","Ad-hoc Remark","Site Approver","Application Approval","Site Approval","Raise CR","Notification Email to Averis IT","Raise Service Request","Blackout","DCM Patch","Scheduled reboot","CyberArk","Installed","Rebooted","Second DCM Patch","Second Reboot","Healh Check","DB services Start/Stop","Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Thursday.csv" -NoTypeInformation
Write-Host "Success" -ForegroundColor Green
#endregion Remove unnecessary columns

#region Convert Thursday.csv to Thursday.xlsx
Write-Host "Converting Thursday.csv to Thursday.xlsx : " -NoNewline
Start-Sleep 2
Start-Process -FilePath 'C:\Program Files\Microsoft Office\Office16\excelcnv.exe' -ArgumentList "-nme -oice ""$ScriptDir\Thursday.csv"" ""$ScriptDir\Thursday.xlsx""" | Out-Null
Write-Host "Success" -ForegroundColor Green
#endregion Convert to .xlsx


#region Format table Thursday.xlsx
Write-Host "Formating table Thursday.xlsx : " -NoNewline
Start-Sleep 2
$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true

$wb = $xl.Workbooks.Open("$ScriptDir\Thursday.xlsx")
$ws = $wb.Sheets.Item(1)

$rows = $ws.UsedRange.Rows.Count

$row = foreach ( $col in "A") {
  $xl.WorksheetFunction.CountIf($ws.Range($col + "1:" + $col + $rows), "<>") - 1
}

$totalrow = $row + 1

$wb.Close()
$xl.Quit()

Start-Sleep 2
$doc = Get-SLDocument "$ScriptDir\Thursday.xlsx"
Start-Sleep 2
Set-SLTableStyle -WorkBookInstance $doc -WorksheetName Thursday -TableStyle Medium13 -StartRowIndex 1 -StartColumnIndex 1 -EndRowIndex $totalrow -EndColumnIndex 27 | 
#Set-SLBorder -Range A1:AA$totalrow -BorderStyle Thin -BorderColor Black -CellBorder | 
Set-SLAutoFitColumn -StartColumnName A -ENDColumnName AA |
Set-SLColumnWidth -ColumnName I -ColumnWidth 70 |
Save-SLDocument
Write-Host "Success" -ForegroundColor Green
#endregion Format table Thursday.xlsx

#SATURDAY
#region Filter the Batch-Saturday column
Write-Host "Filtering the Batch-Saturday column : " -NoNewline
Start-Sleep 2
$File1 = "C:\Users\suhail_asrulsani-ops\Desktop\Excel\Batch3.csv"
$filtered = Import-Csv -Path $File1 | Where-Object { $_."New Batch" -eq "Batch3-Sat" }
Try { New-Item -ItemType "file" -Path "$ScriptDir\Saturday.csv" -Force -ErrorAction Stop | Out-Null } Catch{}
$file2 = "$ScriptDir\Temp_Saturday.csv"
$filtered | Export-Csv -Path $File2 -NoTypeInformation
Write-Host "Success" -ForegroundColor Green
#endregion Filter the Batch-Saturday column

#region Remove unnecessary columns
Write-Host "Remove unnecessary columns : " -NoNewline
Start-Sleep 2
Import-Csv -Path "$ScriptDir\Temp_Saturday.csv" | Select-Object "ServerName","IP Address","Application","Application Owner","Domain","Site","Region","DCM Group","Remark","Ad-hoc Remark","Site Approver","Application Approval","Site Approval","Raise CR","Notification Email to Averis IT","Raise Service Request","Blackout","DCM Patch","Scheduled reboot","CyberArk","Installed","Rebooted","Second DCM Patch","Second Reboot","Healh Check","DB services Start/Stop","Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Saturday.csv" -NoTypeInformation
Write-Host "Success" -ForegroundColor Green
#endregion Remove unnecessary columns

#region Convert Saturday.csv to Saturday.xlsx
Write-Host "Converting Saturday.csv to Saturday.xlsx : " -NoNewline
Start-Sleep 2
Start-Process -FilePath 'C:\Program Files\Microsoft Office\Office16\excelcnv.exe' -ArgumentList "-nme -oice ""$ScriptDir\Saturday.csv"" ""$ScriptDir\Saturday.xlsx""" | Out-Null
Write-Host "Success" -ForegroundColor Green
#endregion Convert to .xlsx

#region Format table Saturday.xlsx
Write-Host "Formating table Saturday.xlsx : " -NoNewline
Start-Sleep 2
$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true

$wb = $xl.Workbooks.Open("$ScriptDir\Saturday.xlsx")
$ws = $wb.Sheets.Item(1)

$rows = $ws.UsedRange.Rows.Count

$row = foreach ( $col in "A") {
  $xl.WorksheetFunction.CountIf($ws.Range($col + "1:" + $col + $rows), "<>") - 1
}

$totalrow = $row + 1

$wb.Close()
$xl.Quit()

Start-Sleep 2
$doc = Get-SLDocument "$ScriptDir\Saturday.xlsx"
Start-Sleep 2
Set-SLTableStyle -WorkBookInstance $doc -WorksheetName Saturday -TableStyle Medium13 -StartRowIndex 1 -StartColumnIndex 1 -EndRowIndex $totalrow -EndColumnIndex 27 | 
#Set-SLBorder -Range A1:AA$totalrow -BorderStyle Thin -BorderColor Black -CellBorder | 
Set-SLAutoFitColumn -StartColumnName A -ENDColumnName AA |
Set-SLColumnWidth -ColumnName I -ColumnWidth 70 |
Save-SLDocument
Write-Host "Success" -ForegroundColor Green
#endregion Format table Saturday.xlsx

#SUNDAY
#region Filter the Batch-Sunday column
Write-Host "Filtering the Batch-Sunday column : " -NoNewline
Start-Sleep 2
$File1 = "C:\Users\suhail_asrulsani-ops\Desktop\Excel\Batch3.csv"
$filtered = Import-Csv -Path $File1 | Where-Object { $_."New Batch" -eq "Batch3-Sun" }
Try { New-Item -ItemType "file" -Path "$ScriptDir\Sunday.csv" -Force -ErrorAction Stop | Out-Null } Catch{}
$file2 = "$ScriptDir\Temp_Sunday.csv"
$filtered | Export-Csv -Path $File2 -NoTypeInformation
Write-Host "Success" -ForegroundColor Green
#endregion Filter the Batch-Sunday column

#region Remove unnecessary columns
Write-Host "Remove unnecessary columns : " -NoNewline
Start-Sleep 2
Import-Csv -Path "$ScriptDir\Temp_Sunday.csv" | Select-Object "ServerName","IP Address","Application","Application Owner","Domain","Site","Region","DCM Group","Remark","Ad-hoc Remark","Site Approver","Application Approval","Site Approval","Raise CR","Notification Email to Averis IT","Raise Service Request","Blackout","DCM Patch","Scheduled reboot","CyberArk","Installed","Rebooted","Second DCM Patch","Second Reboot","Healh Check","DB services Start/Stop","Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Sunday.csv" -NoTypeInformation
Write-Host "Success" -ForegroundColor Green
#endregion Remove unnecessary columns

#region Convert Sunday.csv to Sunday.xlsx
Write-Host "Converting Sunday.csv to Sunday.xlsx : " -NoNewline
Start-Sleep 2
Start-Process -FilePath 'C:\Program Files\Microsoft Office\Office16\excelcnv.exe' -ArgumentList "-nme -oice ""$ScriptDir\Sunday.csv"" ""$ScriptDir\Sunday.xlsx""" | Out-Null
Write-Host "Success" -ForegroundColor Green
#endregion Convert Sunday.csv to Sunday.xlsx

#region Format table Sunday.xlsx
Write-Host "Formating table Sunday.xlsx : " -NoNewline
Start-Sleep 2
$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true

$wb = $xl.Workbooks.Open("$ScriptDir\Sunday.xlsx")
$ws = $wb.Sheets.Item(1)

$rows = $ws.UsedRange.Rows.Count

$row = foreach ( $col in "A") {
  $xl.WorksheetFunction.CountIf($ws.Range($col + "1:" + $col + $rows), "<>") - 1
}

$totalrow = $row + 1

$wb.Close()
$xl.Quit()

Start-Sleep 2
$doc = Get-SLDocument "$ScriptDir\Sunday.xlsx"
Start-Sleep 2
Set-SLTableStyle -WorkBookInstance $doc -WorksheetName Sunday -TableStyle Medium13 -StartRowIndex 1 -StartColumnIndex 1 -EndRowIndex $totalrow -EndColumnIndex 27 | 
#Set-SLBorder -Range A1:AA$totalrow -BorderStyle Thin -BorderColor Black -CellBorder | 
Set-SLAutoFitColumn -StartColumnName A -ENDColumnName AA |
Set-SLColumnWidth -ColumnName I -ColumnWidth 70 |
Save-SLDocument
Write-Host "Success" -ForegroundColor Green
#endregion Format table Sunday.xlsx

#TRANSFER SHEET TO Main excel
#region transfer thursday, saturday and sunday sheet to main excel
Write-Host "Transffering Thursday, Saturday and Sunday sheets to Batch3.xlsx : " -NoNewline
Try
{
    New-SLDocument -WorkbookName Batch3 -WorksheetName "Sheet1" -Path "$ScriptDir" -Force -ErrorAction Stop
    Import-Excel "$ScriptDir\Thursday.xlsx" -ErrorAction Stop | Export-Excel "$ScriptDir\Batch3.xlsx" -WorksheetName "Thursday"
    Import-Excel "$ScriptDir\Saturday.xlsx" -ErrorAction Stop | Export-Excel "$ScriptDir\Batch3.xlsx" -WorksheetName "Saturday"
    Import-Excel "$ScriptDir\Sunday.xlsx" -ErrorAction Stop | Export-Excel "$ScriptDir\Batch3.xlsx" -WorksheetName "Sunday"
    Write-Host "Success" -ForegroundColor Green
}

Catch
{
    Write-Warning ($_)
    Continue
}

Finally
{
    $Error.Clear()
}

#endregion transfer thursday, saturday and sunday sheet to main excel

#region Format sheet Thursday in Batch3.xlsx
Write-Host "Formating Thursday sheet in Batch3.xlsx : " -NoNewline
Start-Sleep 2
$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true

$wb = $xl.Workbooks.Open("$ScriptDir\Thursday.xlsx")
$ws = $wb.Sheets.Item(1)

$rows_thursday = $ws.UsedRange.Rows.Count

$row_thursday = foreach ( $col in "A") {
  $xl.WorksheetFunction.CountIf($ws.Range($col + "1:" + $col + $rows_thursday), "<>") - 1
}

$totalrow_thursday = $row_thursday + 1

$wb.Close()
$xl.Quit()

Start-Sleep 2
$doc = Get-SLDocument "$ScriptDir\Batch3.xlsx"
Start-Sleep 2
Set-SLTableStyle -WorkBookInstance $doc -WorksheetName "Thursday" -TableStyle Medium9 -StartRowIndex 1 -StartColumnIndex 1 -EndRowIndex $totalrow_thursday -EndColumnIndex 27 | 
#Set-SLBorder -Range A1:AA$totalrow -BorderStyle Thin -BorderColor Black -CellBorder | 
Set-SLAutoFitColumn -StartColumnName A -ENDColumnName AA |
Set-SLColumnWidth -ColumnName I -ColumnWidth 70 |
Save-SLDocument
Write-Host "Success" -ForegroundColor Green
#endregion Format sheet Thursday in Batch3.xlsx

#region Format sheet Saturday in Batch3.xlsx
Write-Host "Formating Saturday sheet in Batch3.xlsx : " -NoNewline
Start-Sleep 2
$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true

$wb = $xl.Workbooks.Open("$ScriptDir\Saturday.xlsx")
$ws = $wb.Sheets.Item(1)

$rows_saturday = $ws.UsedRange.Rows.Count

$row_saturday = foreach ( $col in "A") {
  $xl.WorksheetFunction.CountIf($ws.Range($col + "1:" + $col + $rows_saturday), "<>") - 1
}

$totalrow_saturday = $row_saturday + 1

$wb.Close()
$xl.Quit()

Start-Sleep 2
$doc = Get-SLDocument "$ScriptDir\Batch3.xlsx"
Start-Sleep 2
Set-SLTableStyle -WorkBookInstance $doc -WorksheetName "Saturday" -TableStyle Medium9 -StartRowIndex 1 -StartColumnIndex 1 -EndRowIndex $totalrow_saturday -EndColumnIndex 27 | 
#Set-SLBorder -Range A1:AA$totalrow -BorderStyle Thin -BorderColor Black -CellBorder | 
Set-SLAutoFitColumn -StartColumnName A -ENDColumnName AA |
Set-SLColumnWidth -ColumnName I -ColumnWidth 70 |
Save-SLDocument
Write-Host "Success" -ForegroundColor Green
#endregion Format sheet Saturday in Batch3.xlsx

#region Format sheet Sunday in Batch3.xlsx
Write-Host "Formating Sunday sheet in Batch3.xlsx : " -NoNewline
Start-Sleep 2
$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true

$wb = $xl.Workbooks.Open("$ScriptDir\Sunday.xlsx")
$ws = $wb.Sheets.Item(1)

$rows_sunday = $ws.UsedRange.Rows.Count

$row_sunday = foreach ( $col in "A") {
  $xl.WorksheetFunction.CountIf($ws.Range($col + "1:" + $col + $rows_sunday), "<>") - 1
}

$totalrow_Sunday = $row_sunday + 1

$wb.Close()
$xl.Quit()

Start-Sleep 2
$doc = Get-SLDocument "$ScriptDir\Batch3.xlsx"
Start-Sleep 2
Set-SLTableStyle -WorkBookInstance $doc -WorksheetName "Sunday" -TableStyle Medium9 -StartRowIndex 1 -StartColumnIndex 1 -EndRowIndex $totalrow_Sunday -EndColumnIndex 27 | 
#Set-SLBorder -Range A1:AA$totalrow -BorderStyle Thin -BorderColor Black -CellBorder | 
Set-SLAutoFitColumn -StartColumnName A -ENDColumnName AA |
Set-SLColumnWidth -ColumnName I -ColumnWidth 70 |
Save-SLDocument
Write-Host "Success" -ForegroundColor Green
#endregion Format sheet Sunday in Batch3.xlsx

#region Cleanup2
Cleanup2
#endregion Cleanup2



