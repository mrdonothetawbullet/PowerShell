#region Variable
Clear-Host
    Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
    Remove-Variable * -ErrorAction SilentlyContinue; $Error.Clear();
    $ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $datetime = Get-Date -Format G
    $dt = (Get-Date).ToString("ddMMyyyy_HHmmss")
    Try { Import-Module -Name slpslib -ErrorAction Stop -WarningAction SilentlyContinue } Catch{}
    Try { Import-Module -Name Importexcel -ErrorAction Stop -WarningAction SilentlyContinue } Catch{}
#endregion Variable


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
	$formExcelApprovalTool = New-Object 'System.Windows.Forms.Form'
	$tabcontrol1 = New-Object 'System.Windows.Forms.TabControl'
	$tabpage1 = New-Object 'System.Windows.Forms.TabPage'
	$buttonBatch3 = New-Object 'System.Windows.Forms.Button'
	$buttonPilot = New-Object 'System.Windows.Forms.Button'
	$buttonBatch4 = New-Object 'System.Windows.Forms.Button'
	$buttonBatch1 = New-Object 'System.Windows.Forms.Button'
	$buttonBatch2 = New-Object 'System.Windows.Forms.Button'
	$tabpage2 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage3 = New-Object 'System.Windows.Forms.TabPage'
	$statusbar1 = New-Object 'System.Windows.Forms.StatusBar'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$formExcelApprovalTool_Load={
	
		
	}
	
	$buttonPilot_Click={
		#TODO: Place custom script here
		
	}
	
	$buttonBatch1_Click={
		#TODO: Place custom script here
		
	}
	
	$buttonBatch2_Click={
	
		
	}
	
	$buttonBatch3_Click={
		#region Function
		Function Cleanup
		{
			Write-Status -Message "Cleaning up files/directory..."
			Try
			{
				Get-ChildItem -Path "$ScriptDir" -Recurse -Exclude "format it.ps1", "excel approval gui.ps1" -ErrorAction Stop | Remove-Item -Recurse -Force
				
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
				Get-ChildItem -Path "$ScriptDir" -Recurse -Exclude "format it.ps1", "Thursday", "Saturday", "Sunday", "Batch3.xlsx", "excel approval gui.ps1" | Remove-Item -Recurse -Force
				
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
			$wb1 = $xl.workbooks.open($file2) # open target
			$sh1_wb1 = $wb1.sheets.item(1) # second sheet in destination workbook
			$sheetToCopy = $wb2.sheets.item('PatchingList') # source sheet to copy
			$sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook
			$wb2.close($false) # close source workbook w/o saving
			$wb1.close($true) # close and save destination workbook
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
			$Sheet.SaveAs($csv, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSVWindows)
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
		Try { New-Item -ItemType "file" -Path "$ScriptDir\Thursday.csv" -Force -ErrorAction Stop | Out-Null }
		Catch { }
		$file2 = "$ScriptDir\Temp_Thursday.csv"
		$filtered | Export-Csv -Path $File2 -NoTypeInformation
		
		#endregion Filter the Batch-Thursday column
		
		#region Remove unnecessary columns
		Write-Status -Message "Remove unnecessary columns..."
		Start-Sleep 2
		Import-Csv -Path "$ScriptDir\Temp_Thursday.csv" | Select-Object "ServerName", "IP Address", "Application", "Application Owner", "Domain", "Site", "Region", "DCM Group", "Remark", "Ad-hoc Remark", "Site Approver", "Application Approval", "Site Approval", "Raise CR", "Notification Email to Averis IT", "Raise Service Request", "Blackout", "DCM Patch", "Scheduled reboot", "CyberArk", "Installed", "Rebooted", "Second DCM Patch", "Second Reboot", "Healh Check", "DB services Start/Stop", "Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Thursday.csv" -NoTypeInformation
		
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
		
		$row = foreach ($col in "A")
		{
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
		Try { New-Item -ItemType "file" -Path "$ScriptDir\Saturday.csv" -Force -ErrorAction Stop | Out-Null }
		Catch { }
		$file2 = "$ScriptDir\Temp_Saturday.csv"
		$filtered | Export-Csv -Path $File2 -NoTypeInformation
		
		#endregion Filter the Batch-Saturday column
		
		#region Remove unnecessary columns
		Write-Status -Message "Remove unnecessary columns..."
		Start-Sleep 2
		Import-Csv -Path "$ScriptDir\Temp_Saturday.csv" | Select-Object "ServerName", "IP Address", "Application", "Application Owner", "Domain", "Site", "Region", "DCM Group", "Remark", "Ad-hoc Remark", "Site Approver", "Application Approval", "Site Approval", "Raise CR", "Notification Email to Averis IT", "Raise Service Request", "Blackout", "DCM Patch", "Scheduled reboot", "CyberArk", "Installed", "Rebooted", "Second DCM Patch", "Second Reboot", "Healh Check", "DB services Start/Stop", "Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Saturday.csv" -NoTypeInformation
		
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
		
		$row = foreach ($col in "A")
		{
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
		Try { New-Item -ItemType "file" -Path "$ScriptDir\Sunday.csv" -Force -ErrorAction Stop | Out-Null }
		Catch { }
		$file2 = "$ScriptDir\Temp_Sunday.csv"
		$filtered | Export-Csv -Path $File2 -NoTypeInformation
		
		#endregion Filter the Batch-Sunday column
		
		#region Remove unnecessary columns
		Write-Status -Message "Remove unnecessary columns..."
		Start-Sleep 2
		Import-Csv -Path "$ScriptDir\Temp_Sunday.csv" | Select-Object "ServerName", "IP Address", "Application", "Application Owner", "Domain", "Site", "Region", "DCM Group", "Remark", "Ad-hoc Remark", "Site Approver", "Application Approval", "Site Approval", "Raise CR", "Notification Email to Averis IT", "Raise Service Request", "Blackout", "DCM Patch", "Scheduled reboot", "CyberArk", "Installed", "Rebooted", "Second DCM Patch", "Second Reboot", "Healh Check", "DB services Start/Stop", "Move EXC to preferred node" | Export-Csv -Path "$ScriptDir\Sunday.csv" -NoTypeInformation
		
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
		
		$row = foreach ($col in "A")
		{
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
		
		$row_thursday = foreach ($col in "A")
		{
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
		
		$row_saturday = foreach ($col in "A")
		{
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
		
		$row_sunday = foreach ($col in "A")
		{
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
	
	$buttonBatch4_Click={
		#TODO: Place custom script here
		
	}
	
	#region Control Helper Functions
	function Update-ComboBox
	{
	<#
		.SYNOPSIS
			This functions helps you load items into a ComboBox.
		
		.DESCRIPTION
			Use this function to dynamically load items into the ComboBox control.
		
		.PARAMETER ComboBox
			The ComboBox control you want to add items to.
		
		.PARAMETER Items
			The object or objects you wish to load into the ComboBox's Items collection.
		
		.PARAMETER DisplayMember
			Indicates the property to display for the items in this control.
			
		.PARAMETER ValueMember
			Indicates the property to use for the value of the control.
		
		.PARAMETER Append
			Adds the item(s) to the ComboBox without clearing the Items collection.
		
		.EXAMPLE
			Update-ComboBox $combobox1 "Red", "White", "Blue"
		
		.EXAMPLE
			Update-ComboBox $combobox1 "Red" -Append
			Update-ComboBox $combobox1 "White" -Append
			Update-ComboBox $combobox1 "Blue" -Append
		
		.EXAMPLE
			Update-ComboBox $combobox1 (Get-Process) "ProcessName"
		
		.NOTES
			Additional information about the function.
	#>
		
		param
		(
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[System.Windows.Forms.ComboBox]
			$ComboBox,
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			$Items,
			[Parameter(Mandatory = $false)]
			[string]$DisplayMember,
			[Parameter(Mandatory = $false)]
			[string]$ValueMember,
			[switch]
			$Append
		)
		
		if (-not $Append)
		{
			$ComboBox.Items.Clear()
		}
		
		if ($Items -is [Object[]])
		{
			$ComboBox.Items.AddRange($Items)
		}
		elseif ($Items -is [System.Collections.IEnumerable])
		{
			$ComboBox.BeginUpdate()
			foreach ($obj in $Items)
			{
				$ComboBox.Items.Add($obj)
			}
			$ComboBox.EndUpdate()
		}
		else
		{
			$ComboBox.Items.Add($Items)
		}
		
		$ComboBox.DisplayMember = $DisplayMember
		$ComboBox.ValueMember = $ValueMember
	}
	#endregion
	
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formExcelApprovalTool.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$buttonBatch3.remove_Click($buttonBatch3_Click)
			$buttonPilot.remove_Click($buttonPilot_Click)
			$buttonBatch4.remove_Click($buttonBatch4_Click)
			$buttonBatch1.remove_Click($buttonBatch1_Click)
			$buttonBatch2.remove_Click($buttonBatch2_Click)
			$formExcelApprovalTool.remove_Load($formExcelApprovalTool_Load)
			$formExcelApprovalTool.remove_Load($Form_StateCorrection_Load)
			$formExcelApprovalTool.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$formExcelApprovalTool.SuspendLayout()
	$tabcontrol1.SuspendLayout()
	$tabpage1.SuspendLayout()
	#
	# formExcelApprovalTool
	#
	$formExcelApprovalTool.Controls.Add($tabcontrol1)
	$formExcelApprovalTool.Controls.Add($statusbar1)
	$formExcelApprovalTool.AutoScaleDimensions = '7, 15'
	$formExcelApprovalTool.AutoScaleMode = 'Font'
	$formExcelApprovalTool.BackColor = 'RosyBrown'
	$formExcelApprovalTool.ClientSize = '416, 302'
	$formExcelApprovalTool.Font = 'MS Reference Sans Serif, 8.25pt'
	$formExcelApprovalTool.FormBorderStyle = 'Fixed3D'
	$formExcelApprovalTool.Margin = '4, 3, 4, 3'
	$formExcelApprovalTool.Name = 'formExcelApprovalTool'
	$formExcelApprovalTool.Text = 'Excel Approval Tool'
	$formExcelApprovalTool.add_Load($formExcelApprovalTool_Load)
	#
	# tabcontrol1
	#
	$tabcontrol1.Controls.Add($tabpage1)
	$tabcontrol1.Controls.Add($tabpage2)
	$tabcontrol1.Controls.Add($tabpage3)
	$tabcontrol1.Location = '13, 12'
	$tabcontrol1.Margin = '4, 3, 4, 3'
	$tabcontrol1.Name = 'tabcontrol1'
	$tabcontrol1.SelectedIndex = 0
	$tabcontrol1.Size = '390, 244'
	$tabcontrol1.TabIndex = 6
	#
	# tabpage1
	#
	$tabpage1.Controls.Add($buttonBatch3)
	$tabpage1.Controls.Add($buttonPilot)
	$tabpage1.Controls.Add($buttonBatch4)
	$tabpage1.Controls.Add($buttonBatch1)
	$tabpage1.Controls.Add($buttonBatch2)
	$tabpage1.BackColor = 'White'
	$tabpage1.Location = '4, 24'
	$tabpage1.Margin = '4, 3, 4, 3'
	$tabpage1.Name = 'tabpage1'
	$tabpage1.Padding = '3, 3, 3, 3'
	$tabpage1.Size = '382, 216'
	$tabpage1.TabIndex = 0
	$tabpage1.Text = 'Extract'
	#
	# buttonBatch3
	#
	$buttonBatch3.BackColor = 'LightCoral'
	$buttonBatch3.FlatAppearance.MouseDownBackColor = 'IndianRed'
	$buttonBatch3.FlatAppearance.MouseOverBackColor = 'IndianRed'
	$buttonBatch3.FlatStyle = 'Flat'
	$buttonBatch3.Font = 'MS Reference Sans Serif, 8.25pt'
	$buttonBatch3.Location = '134, 105'
	$buttonBatch3.Margin = '4, 3, 4, 3'
	$buttonBatch3.Name = 'buttonBatch3'
	$buttonBatch3.Size = '88, 27'
	$buttonBatch3.TabIndex = 3
	$buttonBatch3.Text = 'Batch3'
	$buttonBatch3.UseCompatibleTextRendering = $True
	$buttonBatch3.UseVisualStyleBackColor = $False
	$buttonBatch3.add_Click($buttonBatch3_Click)
	#
	# buttonPilot
	#
	$buttonPilot.BackColor = 'LightCoral'
	$buttonPilot.FlatAppearance.MouseDownBackColor = 'IndianRed'
	$buttonPilot.FlatAppearance.MouseOverBackColor = 'IndianRed'
	$buttonPilot.FlatStyle = 'Flat'
	$buttonPilot.Font = 'MS Reference Sans Serif, 8.25pt'
	$buttonPilot.Location = '134, 6'
	$buttonPilot.Margin = '4, 3, 4, 3'
	$buttonPilot.Name = 'buttonPilot'
	$buttonPilot.Size = '88, 27'
	$buttonPilot.TabIndex = 0
	$buttonPilot.Text = 'Pilot'
	$buttonPilot.UseCompatibleTextRendering = $True
	$buttonPilot.UseVisualStyleBackColor = $False
	$buttonPilot.add_Click($buttonPilot_Click)
	#
	# buttonBatch4
	#
	$buttonBatch4.BackColor = 'LightCoral'
	$buttonBatch4.FlatAppearance.MouseDownBackColor = 'IndianRed'
	$buttonBatch4.FlatAppearance.MouseOverBackColor = 'IndianRed'
	$buttonBatch4.FlatStyle = 'Flat'
	$buttonBatch4.Font = 'MS Reference Sans Serif, 8.25pt'
	$buttonBatch4.Location = '134, 138'
	$buttonBatch4.Margin = '4, 3, 4, 3'
	$buttonBatch4.Name = 'buttonBatch4'
	$buttonBatch4.Size = '88, 27'
	$buttonBatch4.TabIndex = 4
	$buttonBatch4.Text = 'Batch4'
	$buttonBatch4.UseCompatibleTextRendering = $True
	$buttonBatch4.UseVisualStyleBackColor = $False
	$buttonBatch4.add_Click($buttonBatch4_Click)
	#
	# buttonBatch1
	#
	$buttonBatch1.BackColor = 'LightCoral'
	$buttonBatch1.FlatAppearance.MouseDownBackColor = 'IndianRed'
	$buttonBatch1.FlatAppearance.MouseOverBackColor = 'IndianRed'
	$buttonBatch1.FlatStyle = 'Flat'
	$buttonBatch1.Font = 'MS Reference Sans Serif, 8.25pt'
	$buttonBatch1.Location = '134, 39'
	$buttonBatch1.Margin = '4, 3, 4, 3'
	$buttonBatch1.Name = 'buttonBatch1'
	$buttonBatch1.Size = '88, 27'
	$buttonBatch1.TabIndex = 1
	$buttonBatch1.Text = 'Batch1'
	$buttonBatch1.UseCompatibleTextRendering = $True
	$buttonBatch1.UseVisualStyleBackColor = $False
	$buttonBatch1.add_Click($buttonBatch1_Click)
	#
	# buttonBatch2
	#
	$buttonBatch2.BackColor = 'LightCoral'
	$buttonBatch2.FlatAppearance.MouseDownBackColor = 'IndianRed'
	$buttonBatch2.FlatAppearance.MouseOverBackColor = 'IndianRed'
	$buttonBatch2.FlatStyle = 'Flat'
	$buttonBatch2.Font = 'MS Reference Sans Serif, 8.25pt'
	$buttonBatch2.Location = '134, 72'
	$buttonBatch2.Margin = '4, 3, 4, 3'
	$buttonBatch2.Name = 'buttonBatch2'
	$buttonBatch2.Size = '88, 27'
	$buttonBatch2.TabIndex = 2
	$buttonBatch2.Text = 'Batch2'
	$buttonBatch2.UseCompatibleTextRendering = $True
	$buttonBatch2.UseVisualStyleBackColor = $False
	$buttonBatch2.add_Click($buttonBatch2_Click)
	#
	# tabpage2
	#
	$tabpage2.Location = '4, 24'
	$tabpage2.Margin = '4, 3, 4, 3'
	$tabpage2.Name = 'tabpage2'
	$tabpage2.Padding = '3, 3, 3, 3'
	$tabpage2.Size = '382, 216'
	$tabpage2.TabIndex = 1
	$tabpage2.Text = 'Email Approver'
	$tabpage2.UseVisualStyleBackColor = $True
	#
	# tabpage3
	#
	$tabpage3.Location = '4, 24'
	$tabpage3.Margin = '4, 3, 4, 3'
	$tabpage3.Name = 'tabpage3'
	$tabpage3.Padding = '3, 3, 3, 3'
	$tabpage3.Size = '382, 216'
	$tabpage3.TabIndex = 2
	$tabpage3.Text = 'Email SD'
	$tabpage3.UseVisualStyleBackColor = $True
	#
	# statusbar1
	#
	$statusbar1.Font = 'MS Reference Sans Serif, 8.25pt'
	$statusbar1.Location = '0, 277'
	$statusbar1.Margin = '4, 3, 4, 3'
	$statusbar1.Name = 'statusbar1'
	$statusbar1.Size = '416, 25'
	$statusbar1.TabIndex = 5
	$statusbar1.Text = 'No task is running at the moment'
	$tabpage1.ResumeLayout()
	$tabcontrol1.ResumeLayout()
	$formExcelApprovalTool.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formExcelApprovalTool.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formExcelApprovalTool.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formExcelApprovalTool.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formExcelApprovalTool.ShowDialog()

} #End Function

#Call the form
Show-Excel_Approval_GUI_psf | Out-Null
