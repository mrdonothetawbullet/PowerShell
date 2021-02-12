#region Variable
Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
Remove-Variable * -ErrorAction SilentlyContinue; $Error.Clear();
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$datetime = Get-Date -Format G
$dt = (Get-Date).ToString("ddMMyyyy_HHmmss")
Import-Module -Name slpslib -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
Import-Module -Name Importexcel -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
#endregion Variable
#----------------------------------------------
#region Application Functions
#----------------------------------------------
function Write-Status
	{
		[CmdletBinding()]
		param (
			[Parameter(Mandatory)]
			[ValidateNotNullOrEmpty()]
			[String]$Message
		)
		$statusbar1.Text = $Message
	}
#endregion Application Functions

#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Show-Excel_Approval_GUI_psf {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form1 = New-Object 'System.Windows.Forms.Form'
	$statusbar1 = New-Object 'System.Windows.Forms.StatusBar'
	$button5 = New-Object 'System.Windows.Forms.Button'
	$buttonBatch3 = New-Object 'System.Windows.Forms.Button'
	$button3 = New-Object 'System.Windows.Forms.Button'
	$button2 = New-Object 'System.Windows.Forms.Button'
	$button1 = New-Object 'System.Windows.Forms.Button'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$form1_Load={
		#TODO: Initialize Form Controls here
		
	}
	
	$button1_Click={
		#TODO: Place custom script here
		
	}
	
	$button2_Click={
		#TODO: Place custom script here
		
	}
	
	$button3_Click={
		#TODO: Place custom script here
		
	}
	
	$buttonBatch3_Click={
#region Function
Function Cleanup
{
    Write-Status -Message "Cleaning up files/directory..."
    Try
    {
        Get-ChildItem -Path "$ScriptDir" -Recurse -Exclude "format it.ps1","excel approval gui.ps1" | Remove-Item -Recurse -Force
        
    }

    Catch 
    {
        Write-Status -Message ($_); Continue
    }
}

Function Cleanup2
{
    Write-Status -Message "Cleaning up..."
    Try
    {
        Get-ChildItem -Path "$ScriptDir" -Recurse -Exclude "format it.ps1","Thursday","Saturday","Sunday","Batch3.xlsx","excel approval gui.ps1" | Remove-Item -Recurse -Force
        
    }

    Catch 
    {
        Write-Status -Message ($_); Continue
    }
}
#endregion Function

#region Cleanup
Cleanup
#endregion Cleanup

#region Copy MPL to current directory
$Mpl = "\\kulbscvmfs03\IT\30 Knowledgebase\30 Infrastructure\10 Knowledge Base\10 Server\14 DCM Patching\Patching Team\01 Master Patching List\MasterPatchingList_v0.9.xlsx"
Write-Status -Message "Coying MPL to current directory..."
Try
{
    Copy-Item -Path $Mpl -Destination "$ScriptDir\" -Force -ErrorAction Stop | Out-Null
    
}

Catch 
{
    Write-Status -Message ($_); Continue
}

Finally
{
    $Error.Clear()
}
#endregion Copy MPL to current directory

#region Create Temp_Batch3.csv, Thursday, Saturday and Sunday directory
Write-Status -Message "Creating Temp_Batch3.csv, Thursday, Saturday and Sunday directory..."
Start-Sleep 2
Try
{
    New-SLDocument -WorkbookName Temp_Batch3 -WorksheetName Overwritten -Path "$ScriptDir" -Force -ErrorAction Stop
    New-Item -ItemType Directory -Path "$ScriptDir\Thursday\" -Force -ErrorAction Stop | Out-Null
    New-Item -ItemType Directory -Path "$ScriptDir\Saturday\" -Force -ErrorAction Stop | Out-Null
    New-Item -ItemType Directory -Path "$ScriptDir\Sunday\" -Force -ErrorAction Stop | Out-Null
    

}

Catch
{
    Write-Status -Message ($_)
    Continue
}

Finally
{
    $Error.Clear()
}
#endregion Create Temp_Batch3.csv, Thursday, Saturday and Sunday directory

#region Copy PatchingList sheet from MPL to Temp_Batch3.xlsx
Write-Status -Message "Copying PatchingList sheet from MPL to Temp_Batch3.xlsx..."
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
    
}

Catch
{
    Write-Status -Message ($_)
    Continue
}

Finally
{
    $Error.Clear()
}
#spps -n excel 
#endregion Copy PatchingList sheet from MPL to Temp_Batch3.xlsx

#region Temp_Batch3.xlsx to Temp_Batch3.csv with delimiter
Write-Status -Message "Converting Temp_Batch3.xlsx to Temp_Batch3.csv with delimiter..."
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
    
}

Catch
{
    Write-Status -Message ($_)
    Continue
}

Finally
{
    $Error.Clear()
}
#endregion Convert Temp_Batch3.xlsx to Temp_Batch3.csv with delimiter

#THURSDAY
#region Filter the Batch-Thursday column
Write-Status -Message "Filtering the Batch-Thursday column..."
Start-Sleep 2
$File1 = "$ScriptDir\Batch3.csv"
$filtered = Import-Csv -Path $File1 | Where-Object { $_."New Batch" -eq "Batch3-Thurs" }
Try { New-Item -ItemType "file" -Path "$ScriptDir\Thursday.csv" -Force -ErrorAction Stop | Out-Null } Catch{}
$file2 = "$ScriptDir\Temp_Thursday.csv"
$filtered | Export-Csv -Path $File2 -NoTypeInformation

#endregion Filter the Batch-Thursday column

#region Remove unnecessary columns
Write-Status -Message "Remove unnecessary columns..."
Start-Sleep 2
Import-Csv -Path "$ScriptDir\Temp_Thursday.csv" | Select-Object "ServerName","IP Address","Application","Application Owner","Domain","Site","Region","DCM Group","Remark","Ad-hoc Remark","Site Approver","Application Approval","Site Approval","Raise CR","Notification Email to Averis IT","Raise Service Request","Blackout","DCM Patch","Scheduled reboot","CyberArk","Installed","Rebooted","Second DCM Patch","Second Reboot","Healh Check","DB services Start/Stop","Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Thursday.csv" -NoTypeInformation

#endregion Remove unnecessary columns

#region Convert Thursday.csv to Thursday.xlsx
Write-Status -Message "Converting Thursday.csv to Thursday.xlsx..."
Start-Sleep 2
Start-Process -FilePath 'C:\Program Files\Microsoft Office\Office16\excelcnv.exe' -ArgumentList "-nme -oice ""$ScriptDir\Thursday.csv"" ""$ScriptDir\Thursday.xlsx""" | Out-Null

#endregion Convert to .xlsx

#region Format table Thursday.xlsx
Write-Status -Message "Formating table Thursday.xlsx..."
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

#endregion Format table Thursday.xlsx

#SATURDAY
#region Filter the Batch-Saturday column
Write-Status -Message "Filtering the Batch-Saturday column..."
Start-Sleep 2
$File1 = "$ScriptDir\Batch3.csv"
$filtered = Import-Csv -Path $File1 | Where-Object { $_."New Batch" -eq "Batch3-Sat" }
Try { New-Item -ItemType "file" -Path "$ScriptDir\Saturday.csv" -Force -ErrorAction Stop | Out-Null } Catch{}
$file2 = "$ScriptDir\Temp_Saturday.csv"
$filtered | Export-Csv -Path $File2 -NoTypeInformation

#endregion Filter the Batch-Saturday column

#region Remove unnecessary columns
Write-Status -Message "Remove unnecessary columns..."
Start-Sleep 2
Import-Csv -Path "$ScriptDir\Temp_Saturday.csv" | Select-Object "ServerName","IP Address","Application","Application Owner","Domain","Site","Region","DCM Group","Remark","Ad-hoc Remark","Site Approver","Application Approval","Site Approval","Raise CR","Notification Email to Averis IT","Raise Service Request","Blackout","DCM Patch","Scheduled reboot","CyberArk","Installed","Rebooted","Second DCM Patch","Second Reboot","Healh Check","DB services Start/Stop","Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Saturday.csv" -NoTypeInformation

#endregion Remove unnecessary columns

#region Convert Saturday.csv to Saturday.xlsx
Write-Status -Message "Converting Saturday.csv to Saturday.xlsx..."
Start-Sleep 2
Start-Process -FilePath 'C:\Program Files\Microsoft Office\Office16\excelcnv.exe' -ArgumentList "-nme -oice ""$ScriptDir\Saturday.csv"" ""$ScriptDir\Saturday.xlsx""" | Out-Null

#endregion Convert to .xlsx

#region Format table Saturday.xlsx
Write-Status -Message "Formating table Saturday.xlsx..."
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

#endregion Format table Saturday.xlsx

#SUNDAY
#region Filter the Batch-Sunday column
Write-Status -Message "Filtering the Batch-Sunday column..."
Start-Sleep 2
$File1 = "$ScriptDir\Batch3.csv"
$filtered = Import-Csv -Path $File1 | Where-Object { $_."New Batch" -eq "Batch3-Sun" }
Try { New-Item -ItemType "file" -Path "$ScriptDir\Sunday.csv" -Force -ErrorAction Stop | Out-Null } Catch{}
$file2 = "$ScriptDir\Temp_Sunday.csv"
$filtered | Export-Csv -Path $File2 -NoTypeInformation

#endregion Filter the Batch-Sunday column

#region Remove unnecessary columns
Write-Status -Message "Remove unnecessary columns..."
Start-Sleep 2
Import-Csv -Path "$ScriptDir\Temp_Sunday.csv" | Select-Object "ServerName","IP Address","Application","Application Owner","Domain","Site","Region","DCM Group","Remark","Ad-hoc Remark","Site Approver","Application Approval","Site Approval","Raise CR","Notification Email to Averis IT","Raise Service Request","Blackout","DCM Patch","Scheduled reboot","CyberArk","Installed","Rebooted","Second DCM Patch","Second Reboot","Healh Check","DB services Start/Stop","Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Sunday.csv" -NoTypeInformation

#endregion Remove unnecessary columns

#region Convert Sunday.csv to Sunday.xlsx
Write-Status -Message "Converting Sunday.csv to Sunday.xlsx..."
Start-Sleep 2
Start-Process -FilePath 'C:\Program Files\Microsoft Office\Office16\excelcnv.exe' -ArgumentList "-nme -oice ""$ScriptDir\Sunday.csv"" ""$ScriptDir\Sunday.xlsx""" | Out-Null

#endregion Convert Sunday.csv to Sunday.xlsx

#region Format table Sunday.xlsx
Write-Status -Message "Formating table Sunday.xlsx..."
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

#endregion Format table Sunday.xlsx

#TRANSFER SHEET TO Main excel
#region transfer thursday, saturday and sunday sheet to main excel
Write-Status -Message "Transffering Thursday, Saturday and Sunday sheets to Batch3.xlsx..."
Try
{
    New-SLDocument -WorkbookName Batch3 -WorksheetName "Sheet1" -Path "$ScriptDir" -Force -ErrorAction Stop
    Import-Excel "$ScriptDir\Thursday.xlsx" -ErrorAction Stop | Export-Excel "$ScriptDir\Batch3.xlsx" -WorksheetName "Thursday"
    Import-Excel "$ScriptDir\Saturday.xlsx" -ErrorAction Stop | Export-Excel "$ScriptDir\Batch3.xlsx" -WorksheetName "Saturday"
    Import-Excel "$ScriptDir\Sunday.xlsx" -ErrorAction Stop | Export-Excel "$ScriptDir\Batch3.xlsx" -WorksheetName "Sunday"
    
}

Catch
{
    Write-Status -Message ($_)
    Continue
}

Finally
{
    $Error.Clear()
}

#endregion transfer thursday, saturday and sunday sheet to main excel

#region Format sheet Thursday in Batch3.xlsx
Write-Status -Message "Formating Thursday sheet in Batch3.xlsx..."
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

#endregion Format sheet Thursday in Batch3.xlsx

#region Format sheet Saturday in Batch3.xlsx
Write-Status -Message "Formating Saturday sheet in Batch3.xlsx..."
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

#endregion Format sheet Saturday in Batch3.xlsx

#region Format sheet Sunday in Batch3.xlsx
Write-Status -Message "Formating Sunday sheet in Batch3.xlsx..."
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

#endregion Format sheet Sunday in Batch3.xlsx

#region Cleanup2
Cleanup2
Write-Status -Message "Completed..."
#endregion Cleanup2
		
	}
	
	$button5_Click={
		#TODO: Place custom script here
		
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form1.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$button5.remove_Click($button5_Click)
			$buttonBatch3.remove_Click($buttonBatch3_Click)
			$button3.remove_Click($button3_Click)
			$button2.remove_Click($button2_Click)
			$button1.remove_Click($button1_Click)
			$form1.remove_Load($form1_Load)
			$form1.remove_Load($Form_StateCorrection_Load)
			$form1.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$form1.SuspendLayout()
	#
	# form1
	#
	$form1.Controls.Add($statusbar1)
	$form1.Controls.Add($button5)
	$form1.Controls.Add($buttonBatch3)
	$form1.Controls.Add($button3)
	$form1.Controls.Add($button2)
	$form1.Controls.Add($button1)
	$form1.AutoScaleDimensions = '6, 13'
	$form1.AutoScaleMode = 'Font'
	$form1.ClientSize = '284, 261'
	$form1.Name = 'form1'
	$form1.Text = 'Form'
	$form1.add_Load($form1_Load)
	#
	# statusbar1
	#
	$statusbar1.Location = '0, 239'
	$statusbar1.Name = 'statusbar1'
	$statusbar1.Size = '284, 22'
	$statusbar1.TabIndex = 5
	$statusbar1.Text = 'No task is running at the moment'
	#
	# button5
	#
	$button5.Location = '12, 158'
	$button5.Name = 'button5'
	$button5.Size = '75, 23'
	$button5.TabIndex = 4
	$button5.Text = 'button5'
	$button5.UseCompatibleTextRendering = $True
	$button5.UseVisualStyleBackColor = $True
	$button5.add_Click($button5_Click)
	#
	# buttonBatch3
	#
	$buttonBatch3.Location = '12, 129'
	$buttonBatch3.Name = 'buttonBatch3'
	$buttonBatch3.Size = '75, 23'
	$buttonBatch3.TabIndex = 3
	$buttonBatch3.Text = 'Batch-3'
	$buttonBatch3.UseCompatibleTextRendering = $True
	$buttonBatch3.UseVisualStyleBackColor = $True
	$buttonBatch3.add_Click($buttonBatch3_Click)
	#
	# button3
	#
	$button3.Location = '12, 100'
	$button3.Name = 'button3'
	$button3.Size = '75, 23'
	$button3.TabIndex = 2
	$button3.Text = 'button3'
	$button3.UseCompatibleTextRendering = $True
	$button3.UseVisualStyleBackColor = $True
	$button3.add_Click($button3_Click)
	#
	# button2
	#
	$button2.Location = '12, 71'
	$button2.Name = 'button2'
	$button2.Size = '75, 23'
	$button2.TabIndex = 1
	$button2.Text = 'button2'
	$button2.UseCompatibleTextRendering = $True
	$button2.UseVisualStyleBackColor = $True
	$button2.add_Click($button2_Click)
	#
	# button1
	#
	$button1.Location = '12, 42'
	$button1.Name = 'button1'
	$button1.Size = '75, 23'
	$button1.TabIndex = 0
	$button1.Text = 'button1'
	$button1.UseCompatibleTextRendering = $True
	$button1.UseVisualStyleBackColor = $True
	$button1.add_Click($button1_Click)
	$form1.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form1.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form1.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $form1.ShowDialog()

} #End Function

#Call the form
Show-Excel_Approval_GUI_psf | Out-Null
