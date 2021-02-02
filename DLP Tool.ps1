#------------------------------------------------------------------------
# Source File Information (DO NOT MODIFY)
# Source ID: ef196e2c-42c6-4767-adb1-c13817d8c353
# Source File: C:\Users\Muhammad Suhail\Documents\SAPIEN\PowerShell Studio\Files\DLP_Tool.psf
#------------------------------------------------------------------------

#----------------------------------------------
#region Global Variable Functions
#----------------------------------------------
#BalikPapan
$BPN = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\BPNKRNVMDLPED01\BPNKRNVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Dumai
$DMI = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\PKUVMDLPED01\PKUVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Jakarta
$JKT = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\JKTVMDLPED01\JKTVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Kerinci
$KER = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\KERVMDLPED01\KERVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Marunda
$MAR = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\JKTVMDLPED01\JKTVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Medan
$MED = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\MEDVMDLPED01\MEDVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Padang
$PDG = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\PKUVMDLPED01\PKUVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#PekanBaru
$PKU = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\PKUVMDLPED01\PKUVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Porsea
$PSA = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\PSAVMDLPED01\PSAVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#Beijing
$BJ = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\XHVMDLPED01\XHVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#JiuJiang
$JJ = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\JXVMDLPED01\JXVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Longtan
$LTA = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\XHVMDLPED01\XHVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#Nanjing
$NJ = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\XHVMDLPED01\XHVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#Putian
$PT = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\JXVMDLPED01\JXVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#Rizhao
$RZ = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\RZVMDLPED01\RZVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#ShangHai
$SZDCSH = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\XHVMDLPED01\XHVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#SuQian
$SQ = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\RZVMDLPED01\RZVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#Wuxi
$WX = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\XHVMDLPED01\XHVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#Xiamen
$XM = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\XHVMDLPED01\XHVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#XinHui
$XH = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\XHVMDLPED01\XHVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#Zhangzhou
$SZDCZZ = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\XHVMDLPED01\XHVMDLPED01_AgentInstallers_15.1 MP2.zip\AgentInstaller_Win64\*"
#KL
$KUL = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\KULBSCVMDLPED01\KULBSCVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#endregion Application Functions

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
function Show-DLP_Tool_psf {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formDLPTool = New-Object 'System.Windows.Forms.Form'
	$statusbar1 = New-Object 'System.Windows.Forms.StatusBar'
	$richtextbox1 = New-Object 'System.Windows.Forms.RichTextBox'
	$labelOutput = New-Object 'System.Windows.Forms.Label'
	$textbox1 = New-Object 'System.Windows.Forms.TextBox'
	$labelServerList = New-Object 'System.Windows.Forms.Label'
	$menustrip1 = New-Object 'System.Windows.Forms.MenuStrip'
	$generalToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$installRemoveToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$reportLogToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$restartToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$pingToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$openDLPSharedDriveToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$servicesToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$restartServerToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$dLPStatusToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$getInstallationLogInstallLogAndCleanupLogToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$removeDLPToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$installDLPToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$reInstallDLPToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$stopDLPServicesToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$viewDLPServicesToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$getVersionToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$formDLPTool_Load={
		#TODO: Initialize Form Controls here
		#$Credential = Get-Credential
	}
	
	$pingToolStripMenuItem_Click={
        Write-Status -Message "Task 'Ping' currently running..."
		$richtextbox1.Clear()
        $Servers = $textbox1.lines.Split("`n")

        Foreach ($Server in $Servers)
        {
            
            $richtextbox1.AppendText("$Server : ")
            If (Test-Connection $Server -Count 1 -Quiet) 
            {
                $richtextbox1.SelectionColor = "#007F00"
                $richtextbox1.AppendText("Online`n")
            }

            else 
            {
                $richtextbox1.SelectionColor = "#FF0000"
                $richtextbox1.AppendText("Offline`n")
            } 
        }
        Write-Status -Message "Task 'Ping' completed..."
		
	}
	
	$getVersionToolStripMenuItem_Click={
        Write-Status -Message "Task 'Get Version' currently running..."
		$richtextbox1.Clear()
        $Servers = $textbox1.lines.Split("`n")

        Foreach ($Server in $Servers) {
            Get-PSSession | Remove-PSSession
            $richtextbox1.AppendText("$Server : ")
            Try { 
            $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop
        }

            Catch {
            $richtextbox1.SelectionColor = "#FF0000"
            $richtextbox1.AppendText("Failed")
            Continue
        }

            Finally { 
            $Error.Clear() 
        }

            $MyCommands = {
                $Path1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                $Installed1 = Get-ChildItem -Path $Path1 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
                $Version1 = ($Installed1).Displayversion

                $Path2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                $Installed2 = Get-ChildItem -Path $Path2 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
                $Version2 = ($Installed2).Displayversion

                If ($Version1) {
                $Version1
            }

                ElseIf ($Version2) {
                $Version2
            }

                Else {
                $Version3 = "Unable to find version"
            }
            }

            $Ver = Invoke-Command -Session $MySession -ScriptBlock $MyCommands

            If ($Ver -eq $null) 
            {
                $richtextbox1.SelectionColor = "#FF0000"
                $richtextbox1.AppendText("Unable to find version.`n")
            }

            ElseIf ($Ver)
            {
                $richtextbox1.SelectionColor = "#007F00"
                $richtextbox1.AppendText("$Ver`n")
            }
        }
		
        Write-Status -Message "Task 'Get Version' completed..."
	}
	
	$openDLPSharedDriveToolStripMenuItem_Click={
        Write-Status -Message "Task 'Opening DLP Shared Drive' currently running..."
		explorer "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP"
        Write-Status -Message "Task 'Opening DLP Shared Drive' completed..."
	}
	
	$stopDLPServicesToolStripMenuItem_Click={
		#TODO: Place custom script here
		
	}
	
	$viewDLPServicesToolStripMenuItem_Click={
		#TODO: Place custom script here
		
	}
	
	$removeDLPToolStripMenuItem_Click={
		#TODO: Place custom script here
		
	}
	
	$installDLPToolStripMenuItem_Click={
		#TODO: Place custom script here
		
	}
	
	$reInstallDLPToolStripMenuItem_Click={
		#TODO: Place custom script here
		
	}
	
	$dLPStatusToolStripMenuItem_Click={
		#TODO: Place custom script here
		
	}
	
	$getInstallationLogInstallLogAndCleanupLogToolStripMenuItem_Click={
		#TODO: Place custom script here
		
	}
	
	$restartServerToolStripMenuItem_Click={
		#TODO: Place custom script here
		
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formDLPTool.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$formDLPTool.remove_Load($formDLPTool_Load)
			$pingToolStripMenuItem.remove_Click($pingToolStripMenuItem_Click)
			$openDLPSharedDriveToolStripMenuItem.remove_Click($openDLPSharedDriveToolStripMenuItem_Click)
			$restartServerToolStripMenuItem.remove_Click($restartServerToolStripMenuItem_Click)
			$dLPStatusToolStripMenuItem.remove_Click($dLPStatusToolStripMenuItem_Click)
			$getInstallationLogInstallLogAndCleanupLogToolStripMenuItem.remove_Click($getInstallationLogInstallLogAndCleanupLogToolStripMenuItem_Click)
			$removeDLPToolStripMenuItem.remove_Click($removeDLPToolStripMenuItem_Click)
			$installDLPToolStripMenuItem.remove_Click($installDLPToolStripMenuItem_Click)
			$reInstallDLPToolStripMenuItem.remove_Click($reInstallDLPToolStripMenuItem_Click)
			$stopDLPServicesToolStripMenuItem.remove_Click($stopDLPServicesToolStripMenuItem_Click)
			$viewDLPServicesToolStripMenuItem.remove_Click($viewDLPServicesToolStripMenuItem_Click)
			$getVersionToolStripMenuItem.remove_Click($getVersionToolStripMenuItem_Click)
			$formDLPTool.remove_Load($Form_StateCorrection_Load)
			$formDLPTool.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$formDLPTool.SuspendLayout()
	$menustrip1.SuspendLayout()
	#
	# formDLPTool
	#
	$formDLPTool.Controls.Add($statusbar1)
	$formDLPTool.Controls.Add($richtextbox1)
	$formDLPTool.Controls.Add($labelOutput)
	$formDLPTool.Controls.Add($textbox1)
	$formDLPTool.Controls.Add($labelServerList)
	$formDLPTool.Controls.Add($menustrip1)
	$formDLPTool.AutoScaleDimensions = '7, 17'
	$formDLPTool.AutoScaleMode = 'Font'
	$formDLPTool.BackColor = 'White'
	$formDLPTool.BackgroundImageLayout = 'Center'
	$formDLPTool.ClientSize = '456, 376'
	$formDLPTool.Font = 'Segoe UI, 9.75pt'
	$formDLPTool.ForeColor = 'Black'
	$formDLPTool.FormBorderStyle = 'Fixed3D'
	#region Binary Data
	$formDLPTool.Icon = [System.Convert]::FromBase64String('
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
Bv8f85kJWqbvGVYAAAAASUVORK5CYII=')
	#endregion
	$formDLPTool.MainMenuStrip = $menustrip1
	$formDLPTool.Margin = '4, 4, 4, 4'
	$formDLPTool.MaximizeBox = $False
	$formDLPTool.Name = 'formDLPTool'
	$formDLPTool.StartPosition = 'CenterParent'
	$formDLPTool.Text = 'DLP Tool'
	$formDLPTool.add_Load($formDLPTool_Load)
	#
	# statusbar1
	#
	$statusbar1.Location = '0, 347'
	$statusbar1.Margin = '4, 4, 4, 4'
	$statusbar1.Name = 'statusbar1'
	$statusbar1.Size = '456, 29'
	$statusbar1.TabIndex = 6
	$statusbar1.Text = 'No task is running at the moment'
	#
	# richtextbox1
	#
	$richtextbox1.Font = 'Segoe UI Semibold, 9.75pt, style=Bold'
	$richtextbox1.Location = '137, 55'
	$richtextbox1.Margin = '4, 4, 4, 4'
	$richtextbox1.Name = 'richtextbox1'
	$richtextbox1.Size = '309, 288'
	$richtextbox1.TabIndex = 5
	$richtextbox1.Text = ''
	#
	# labelOutput
	#
	$labelOutput.AutoSize = $True
	$labelOutput.Location = '268, 29'
	$labelOutput.Margin = '4, 0, 4, 0'
	$labelOutput.Name = 'labelOutput'
	$labelOutput.Size = '47, 22'
	$labelOutput.TabIndex = 4
	$labelOutput.Text = 'Output'
	$labelOutput.UseCompatibleTextRendering = $True
	#
	# textbox1
	#
	$textbox1.Font = 'Segoe UI Semibold, 9.75pt, style=Bold'
	$textbox1.Location = '13, 55'
	$textbox1.Margin = '4, 4, 4, 4'
	$textbox1.Multiline = $True
	$textbox1.Name = 'textbox1'
	$textbox1.Size = '116, 288'
	$textbox1.TabIndex = 2
    $textbox1.WordWrap = $False
	#
	# labelServerList
	#
	$labelServerList.AutoSize = $True
	$labelServerList.FlatStyle = 'Popup'
	$labelServerList.Location = '33, 29'
	$labelServerList.Margin = '4, 0, 4, 0'
	$labelServerList.Name = 'labelServerList'
	$labelServerList.Size = '65, 22'
	$labelServerList.TabIndex = 1
	$labelServerList.Text = 'Server List'
	$labelServerList.UseCompatibleTextRendering = $True
	#
	# menustrip1
	#
	$menustrip1.BackColor = 'Gold'
	$menustrip1.Font = 'Segoe UI, 9pt'
	[void]$menustrip1.Items.Add($generalToolStripMenuItem)
	[void]$menustrip1.Items.Add($servicesToolStripMenuItem)
	[void]$menustrip1.Items.Add($installRemoveToolStripMenuItem)
	[void]$menustrip1.Items.Add($reportLogToolStripMenuItem)
	[void]$menustrip1.Items.Add($restartToolStripMenuItem)
	$menustrip1.Location = '0, 0'
	$menustrip1.Name = 'menustrip1'
	$menustrip1.Padding = '7, 3, 0, 3'
	$menustrip1.Size = '456, 25'
	$menustrip1.TabIndex = 0
	$menustrip1.Text = 'menustrip1'
	#
	# generalToolStripMenuItem
	#
	[void]$generalToolStripMenuItem.DropDownItems.Add($pingToolStripMenuItem)
	[void]$generalToolStripMenuItem.DropDownItems.Add($getVersionToolStripMenuItem)
	[void]$generalToolStripMenuItem.DropDownItems.Add($openDLPSharedDriveToolStripMenuItem)
	$generalToolStripMenuItem.Name = 'generalToolStripMenuItem'
	$generalToolStripMenuItem.Size = '59, 19'
	$generalToolStripMenuItem.Text = 'General'
	#
	# installRemoveToolStripMenuItem
	#
	[void]$installRemoveToolStripMenuItem.DropDownItems.Add($removeDLPToolStripMenuItem)
	[void]$installRemoveToolStripMenuItem.DropDownItems.Add($installDLPToolStripMenuItem)
	[void]$installRemoveToolStripMenuItem.DropDownItems.Add($reInstallDLPToolStripMenuItem)
	$installRemoveToolStripMenuItem.Name = 'installRemoveToolStripMenuItem'
	$installRemoveToolStripMenuItem.Size = '98, 19'
	$installRemoveToolStripMenuItem.Text = 'Install/Remove'
	#
	# reportLogToolStripMenuItem
	#
	[void]$reportLogToolStripMenuItem.DropDownItems.Add($dLPStatusToolStripMenuItem)
	[void]$reportLogToolStripMenuItem.DropDownItems.Add($getInstallationLogInstallLogAndCleanupLogToolStripMenuItem)
	$reportLogToolStripMenuItem.Name = 'reportLogToolStripMenuItem'
	$reportLogToolStripMenuItem.Size = '79, 19'
	$reportLogToolStripMenuItem.Text = 'Report/Log'
	#
	# restartToolStripMenuItem
	#
	[void]$restartToolStripMenuItem.DropDownItems.Add($restartServerToolStripMenuItem)
	$restartToolStripMenuItem.Name = 'restartToolStripMenuItem'
	$restartToolStripMenuItem.Size = '55, 19'
	$restartToolStripMenuItem.Text = 'Restart'
	#
	# pingToolStripMenuItem
	#
	$pingToolStripMenuItem.Name = 'pingToolStripMenuItem'
	$pingToolStripMenuItem.Size = '196, 22'
	$pingToolStripMenuItem.Text = 'Ping'
	$pingToolStripMenuItem.add_Click($pingToolStripMenuItem_Click)
	#
	# openDLPSharedDriveToolStripMenuItem
	#
	$openDLPSharedDriveToolStripMenuItem.Name = 'openDLPSharedDriveToolStripMenuItem'
	$openDLPSharedDriveToolStripMenuItem.Size = '196, 22'
	$openDLPSharedDriveToolStripMenuItem.Text = 'Open DLP Shared Drive'
	$openDLPSharedDriveToolStripMenuItem.add_Click($openDLPSharedDriveToolStripMenuItem_Click)
	#
	# servicesToolStripMenuItem
	#
	[void]$servicesToolStripMenuItem.DropDownItems.Add($stopDLPServicesToolStripMenuItem)
	[void]$servicesToolStripMenuItem.DropDownItems.Add($viewDLPServicesToolStripMenuItem)
	$servicesToolStripMenuItem.Name = 'servicesToolStripMenuItem'
	$servicesToolStripMenuItem.Size = '61, 19'
	$servicesToolStripMenuItem.Text = 'Services'
	#
	# restartServerToolStripMenuItem
	#
	$restartServerToolStripMenuItem.Name = 'restartServerToolStripMenuItem'
	$restartServerToolStripMenuItem.Size = '145, 22'
	$restartServerToolStripMenuItem.Text = 'Restart Server'
	$restartServerToolStripMenuItem.add_Click($restartServerToolStripMenuItem_Click)
	#
	# dLPStatusToolStripMenuItem
	#
	$dLPStatusToolStripMenuItem.Name = 'dLPStatusToolStripMenuItem'
	$dLPStatusToolStripMenuItem.Size = '329, 22'
	$dLPStatusToolStripMenuItem.Text = 'DLP Status'
	$dLPStatusToolStripMenuItem.add_Click($dLPStatusToolStripMenuItem_Click)
	#
	# getInstallationLogInstallLogAndCleanupLogToolStripMenuItem
	#
	$getInstallationLogInstallLogAndCleanupLogToolStripMenuItem.Name = 'getInstallationLogInstallLogAndCleanupLogToolStripMenuItem'
	$getInstallationLogInstallLogAndCleanupLogToolStripMenuItem.Size = '329, 22'
	$getInstallationLogInstallLogAndCleanupLogToolStripMenuItem.Text = 'Get Installation Log, Install Log and Cleanup Log'
	$getInstallationLogInstallLogAndCleanupLogToolStripMenuItem.add_Click($getInstallationLogInstallLogAndCleanupLogToolStripMenuItem_Click)
	#
	# removeDLPToolStripMenuItem
	#
	$removeDLPToolStripMenuItem.Name = 'removeDLPToolStripMenuItem'
	$removeDLPToolStripMenuItem.Size = '147, 22'
	$removeDLPToolStripMenuItem.Text = 'Remove DLP'
	$removeDLPToolStripMenuItem.add_Click($removeDLPToolStripMenuItem_Click)
	#
	# installDLPToolStripMenuItem
	#
	$installDLPToolStripMenuItem.Name = 'installDLPToolStripMenuItem'
	$installDLPToolStripMenuItem.Size = '147, 22'
	$installDLPToolStripMenuItem.Text = 'Install DLP'
	$installDLPToolStripMenuItem.add_Click($installDLPToolStripMenuItem_Click)
	#
	# reInstallDLPToolStripMenuItem
	#
	$reInstallDLPToolStripMenuItem.Name = 'reInstallDLPToolStripMenuItem'
	$reInstallDLPToolStripMenuItem.Size = '147, 22'
	$reInstallDLPToolStripMenuItem.Text = 'Re-Install DLP'
	$reInstallDLPToolStripMenuItem.add_Click($reInstallDLPToolStripMenuItem_Click)
	#
	# stopDLPServicesToolStripMenuItem
	#
	$stopDLPServicesToolStripMenuItem.Name = 'stopDLPServicesToolStripMenuItem'
	$stopDLPServicesToolStripMenuItem.Size = '168, 22'
	$stopDLPServicesToolStripMenuItem.Text = 'Stop DLP Services'
	$stopDLPServicesToolStripMenuItem.add_Click($stopDLPServicesToolStripMenuItem_Click)
	#
	# viewDLPServicesToolStripMenuItem
	#
	$viewDLPServicesToolStripMenuItem.Name = 'viewDLPServicesToolStripMenuItem'
	$viewDLPServicesToolStripMenuItem.Size = '168, 22'
	$viewDLPServicesToolStripMenuItem.Text = 'View DLP Services'
	$viewDLPServicesToolStripMenuItem.add_Click($viewDLPServicesToolStripMenuItem_Click)
	#
	# getVersionToolStripMenuItem
	#
	$getVersionToolStripMenuItem.Name = 'getVersionToolStripMenuItem'
	$getVersionToolStripMenuItem.Size = '196, 22'
	$getVersionToolStripMenuItem.Text = 'Get Version'
	$getVersionToolStripMenuItem.add_Click($getVersionToolStripMenuItem_Click)
	$menustrip1.ResumeLayout()
	$formDLPTool.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formDLPTool.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formDLPTool.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formDLPTool.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formDLPTool.ShowDialog()

} #End Function

#Call the form
Show-DLP_Tool_psf | Out-Null
