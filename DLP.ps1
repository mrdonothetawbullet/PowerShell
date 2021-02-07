Clear-Host

#region Global Variable Functions
#----------------------------------------------
Get-PSSession | Remove-PSSession
Remove-Variable * -ErrorAction SilentlyContinue; $Error.Clear();
$Password = ConvertTo-SecureString "Welcome12" -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential ("globalnet\suhail_asrulsani-ops", $Password)
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
$datetime = Get-Date -Format G
$dt = (Get-Date).ToString("ddMMyyyy_HHmmss") 
$translocation4 = "$ScriptDir\Verify DLP Installation Status-$dt.txt"
$Location = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\15.1 MP2 - Service Shutdown & Uninstall\service_shutdown.exe"
Try { Remove-Item -Path "$ScriptDir\DLPStatus.xlsx" -Force -Recurse -ErrorAction Stop } Catch {}
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
#Jakarta
$RAP = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\Agent Installation Instructions\JKTVMDLPED01\JKTVMDLPED01_AgentInstallers_15.1 MP2\AgentInstaller_Win64\*"
#endregion Global Variable Functions

#region Functions
#----------------------------------------------

Function Function_Zero 
{
    notepad.exe "$ScriptDir\machinelist.txt"
    Get-PSSession | Remove-PSSession
}

Function Function_One
{
    Get-PSSession | Remove-PSSession
    $Machinelist = @(get-content -Path "$ScriptDir\machinelist.txt")
    Write-Host "`n"   
    Foreach ($Machine in $machinelist)
    { 
        Write-Host $Machine ": " -NoNewline
        if(Test-Connection $Machine -Count 1 -Quiet){ 
        Write-Host "Online" -ForegroundColor Green 
    }

    else
    {
        Write-Host "Offline" -ForegroundColor Red 
    }

    }
}

Function Function_Two
{
$Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
Foreach ($Machine in $machinelist)
    {
        $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
        Get-PSSession | Remove-PSSession
        Write-Host "$machine : " -NoNewline
        Try { $MySession = New-PSSession -ComputerName $machine -Credential $Cred -ErrorAction Stop; }
        Catch { Write-Host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
        Finally { $Error.Clear() }

        $MyCommands =
        {
            $Path1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $Installed1 = Get-ChildItem -Path $Path1 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
            $Version1 = ($Installed1).Displayversion

            $Path2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            $Installed2 = Get-ChildItem -Path $Path2 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
            $Version2 = ($Installed2).Displayversion

            If ($Version1) 
            { 
            Write-Host "$Version1" -ForegroundColor Green
            }

            ElseIf ($Version2) 
            { 
            Write-Host "$Version2" -ForegroundColor Green
            }

            Else 
            { 
            Write-Warning "Unable to find version" 
            }
        }
    Invoke-Command -Session $MySession -ScriptBlock $MyCommands -ErrorAction Stop
    }
    Write-Host "`n"
}

Function Function_Three
{
    $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
    Foreach ($machine in $Machinelist)
{
    Write-Host "`n"
    Write-Host "$Machine" -ForegroundColor White -BackgroundColor Red
    Write-Host "Checking connection to $Machine : " -NoNewline
    Try 
    { 
        $MySession = New-PSSession -ComputerName $Machine -ErrorAction Stop
        Write-Host "SUCCESS" -ForegroundColor Green 
    }

    Catch 
    {
         Write-Warning ($_.Exception.Message)
         Continue
    }

    Finally 
    {
        $Error.Clear()
    }

 
    Write-Host "Site : " -NoNewline
    If ($Machine -match '^DMI')
    {
        Write-Host "Dumai" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $DMI -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^KUL')
    {
        Write-Host "KualaLumpur" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $KUL -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^BPN')
    {
        Write-Host "BalikPapan" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $BPN -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^JKT')
    {
        Write-Host "Jakarta" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $JKT -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^RAP')
    {
        Write-Host "Jakarta" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $JKT -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^KER')
    {
        Write-Host "Kerinchi / Pectech" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $KER -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^MAR')
    {
        Write-Host "Marunda" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $JKT -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^MED')
    {
        Write-Host "Marunda" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $MAR -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^PDG')
    {
        Write-Host "Padang" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $PKU -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^PKU')
    {
        Write-Host "PekanBaru" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $PKU -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^PSA')
    {
        Write-Host "Porsea" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $PSA -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^BJ')
    {
        Write-Host "Beijing" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $BJ -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^LTA')
    {
        Write-Host "Longtan" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $LTA -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^NJ')
    {
        Write-Host "Nanjing" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $NJ -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^RZ')
    {
        Write-Host "Rizhao" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $PSA -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^SZDCSH')
    {
        Write-Host "ShangHai" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $SZDCSH -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^SQ')
    {
        Write-Host "SuQian" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $SQ -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^WX')
    {
        Write-Host "Wuxi" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $WX -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^XM')
    {
        Write-Host "Xiamen" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $XM -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match '^XH')
    {
        Write-Host "Xinhui" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $XH -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match 'SZDCZZ*')
    {
        Write-Host "Zhangzhou" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $SZDCZZ -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

    ElseIf ($Machine -match 'JJ*')
    {
        Write-Host "JiuJiang" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $JJ -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }

     ElseIf ($Machine -match 'PTA*')
    {
        Write-Host "Putian" -ForegroundColor Green
        Write-Host "Copying Installer : " -NoNewline

        $Patches = "\\$Machine\c$\Patches\"
        $InstallAgent = "\\$Machine\c$\Patches\Install_Agent.bat"

        New-Item -ItemType Directory -Path $Patches -Force -ErrorAction Stop | Out-Null
        Copy-Item $PTA -Destination $Patches -Force -ErrorAction Stop | Out-Null
        If (Test-Path $InstallAgent) { Write-host "SUCCESS" -ForegroundColor Green  }
        Elseif (!(Test-Path $InstallAgent)) { Write-Host "FAILED" -ForegroundColor Red }
    }
    
    }
}

Function Function_Four
{
    $Input_restart = Read-Host "Stop DLP Agent Service Now? (y/n)"
    switch ($Input_restart) 
    {
        'y'
        {
            $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
            Foreach ($Machine in $Machinelist)
            {
                Write-Host "`n"
                #region Check PSSession Connection
                Write-Host "Establishing remote connection to $machine : "  -NoNewline
                Try { $MySession = New-PSSession -ComputerName $machine -Credential $Cred -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                Catch { Write-Host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
                #endregion Check PSSession Connection

                #region Stop EDPA & WDPA by running service_shutdown.exe
                Write-Host "Stopping EDPA & WDP Services : " -NoNewline

                Try { Copy-Item -Path $Location -Destination "\\$machine\c$\Program Files\Manufacturer\Endpoint Agent\" -ErrorAction Stop } Catch {}
                Try { Copy-Item -Path $Location -Destination "\\$machine\c$\Program Files\Manufacturer (x86)\Endpoint Agent\" -ErrorAction Stop } Catch {}

                $MyCommands = 
                {
                    Try { cd "C:\Program Files\Manufacturer\Endpoint Agent" -ErrorAction Stop } Catch { }
                    cmd.exe /c "service_shutdown.exe -p=Welcome1" > $null 2>&1
                    Try { cd "C:\Program Files (x86)\Manufacturer\Endpoint Agent" -ErrorAction Stop } Catch { }
                    cmd.exe /c "service_shutdown.exe -p=Welcome1" > $null 2>&1

                    $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State

                    If (($Edpa -eq "Stopped") -and ($Wdp -eq "Stopped") ) { Write-Host "Both service is now stopped." -ForegroundColor Green }
                    ElseIf (!($Edpa) -and !($Wdp)) { Write-Host "Both service is not exist." -ForegroundColor Green }
                    Else {Write-Host "Failed to stop service."}
                }

                Invoke-Command -Session $MySession -ScriptBlock $MyCommands
                
            #endregion Stop EDPA & WDPA by running service_shutdown.exe
            }
        }

        'n' 
        { 
            Continue
        }

        Default { Write-Warning "Invalid Input" }
    }
}

Function Function_Five
{
    $Input_restart = Read-Host "Uninstall DLP Agent Now? (y/n)"
    switch ($Input_restart) 
    {
        'y'
        {
            $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
            Foreach ($Machine in $Machinelist)
            {
                Write-Host "`n"
                #region Check PSSession Connection
                Write-Host "Establishing remote connection to $machine : "  -NoNewline
                Try { $MySession = New-PSSession -ComputerName $machine -Credential $Cred -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                Catch { Write-Host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
                #endregion Check PSSession Connection

                #region Stop EDPA & WDPA by running service_shutdown.exe
                Write-Host "Stopping EDPA & WDP Services : " -NoNewline

                Try { Copy-Item -Path $Location -Destination "\\$machine\c$\Program Files\Manufacturer\Endpoint Agent\" -ErrorAction Stop } Catch {}
                Try { Copy-Item -Path $Location -Destination "\\$machine\c$\Program Files\Manufacturer (x86)\Endpoint Agent\" -ErrorAction Stop } Catch {}

                $MyCommands = 
                {
                    Try { cd "C:\Program Files\Manufacturer\Endpoint Agent" -ErrorAction Stop } Catch { }
                    cmd.exe /c "service_shutdown.exe -p=Welcome1" > $null 2>&1
                    Try { cd "C:\Program Files (x86)\Manufacturer\Endpoint Agent" -ErrorAction Stop } Catch { }
                    cmd.exe /c "service_shutdown.exe -p=Welcome1" > $null 2>&1

                    $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State

                    If (($Edpa -eq "Stopped") -and ($Wdp -eq "Stopped") ) { Write-Host "Both service is now stopped. Uninstallation will continue now" -ForegroundColor Green }
                    ElseIf (!($Edpa) -and !($Wdp)) { Write-Host "Both service is not exist. Uninstallation will continue now" -ForegroundColor Green }
                    Else {Write-Host "Failed to stop service. Uninstallation will continue anyway"}
                }

                Invoke-Command -Session $MySession -ScriptBlock $MyCommands
                
            #endregion Stop EDPA & WDPA by running service_shutdown.exe

                #region Executing clean_agent.exe
                Write-Host "Executing clean_agent.exe : " -NoNewline

                Try { $OS = (Get-WmiObject Win32_OperatingSystem -ErrorAction Stop).OsArchitecture } Catch { } Finally { $Error.Clear() }
                
                $Path = "\\$machine\c$\Patches\"
                $Source64 = "$ScriptDir\x64\clean_agent.exe"
                $Source32 = "$ScriptDir\x86\clean_agent.exe"

                If ($OS -eq "64-bit") { Try { Copy-Item $Source64 -Destination $Path -Force -Recurse -ErrorAction Stop } Catch { } }
                ElseIf ($OS -eq "32-bit") { Try { Copy-Item $Source32 -Destination $Path -Force -Recurse -ErrorAction Stop } Catch { } }

                $MyCommands2 =
                {
                    Try { cd "C:\Patches" -ErrorAction Stop } Catch { }
                    cmd.exe /c "echo y | clean_agent.exe -p=Welcome1" #> $null 2>&1
                }

                Invoke-Command -Session $MySession -ScriptBlock $MyCommands2
                #endregion Executing clean_agent.exe

                #region Try to Uninstall using msiexec /x method
                $MyCommands3 = 
                {
                    Write-Host "Perform uninstallation using msiexec method and remove leftover : " -NoNewline

                    $Path1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" > $null 2>&1
                    $Installed1 = Get-ChildItem -Path $Path1 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") } > $null 2>&1
                    $String1 = ($Installed1).UninstallString > $null 2>&1
                    $Path2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" > $null 2>&1
                    $Installed2 = Get-ChildItem -Path $Path2 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") } > $null 2>&1
                    $String2 = ($Installed2).UninstallString > $null 2>&1

                    If ($String1) 
                    { 
                        Write-Host "$String1" -ForegroundColor Green
                        Write-Host "Uninstalling DLP : " -NoNewline
                        cmd.exe /c "$String1 /q UNINSTALLPASSWORD=Welcome1" > $null 2>&1
                        Write-Host "Done" -ForegroundColor Green

                    }

                    ElseIf ($String2) 
                    { 
                        Write-Host "$String2" -ForegroundColor Green
                        Write-Host "Uninstalling DLP : " -NoNewline
                        cmd.exe /c "$String2 /q UNINSTALLPASSWORD=Welcome1" > $null 2>&1
                        Write-Host "Done" -ForegroundColor Green
                    }

                    Else
                    {
                        Write-Host "Done" -ForegroundColor Green
                    }


     
                }

                Invoke-Command -Session $MySession -ScriptBlock $MyCommands3
                #endregion Try to Uninstall using msiexec /x method

                #region Checking status after uninstallation
                Write-Host "Checking Uninstallation Status : " -NoNewline
                $MyCommands4 =
                {

                    $PathVersion = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                    $PathVersion2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                    $Installed1 = Get-ChildItem -Path $PathVersion | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
                    $Installed2 = Get-ChildItem -Path $PathVersion2 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }

                    If ($Installed1) { $Version = ($Installed1).Displayversion }
                    ElseIf ($Installed2) { $Version = ($Installed2).Displayversion }
                    Else { $Version = "False" }

                    # Check vfsmfd.sys, vnwcd.sys and vrtam.sys
                    $Path1 = "C:\Windows\System32\drivers\vfsmfd.sys"; $Path2 = "C:\Windows\System32\drivers\vnwcd.sys"; $Path3 = "C:\Windows\System32\drivers\vrtam.sys";
                    If (Test-Path -Path $Path1) { $vfsmfd = "True" } Else { $vfsmfd = "False" }
                    If (Test-Path -Path $Path2) { $vnwcd = "True" } Else { $vnwcd = "False" }
                    If (Test-Path -Path $Path3) { $vrtam = "True" } Else { $vrtam = "False" }

                    # Check log file
                    $logfile1 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa0.log"; $logfile2 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa_ext0.log"; $Folder1 = "C:\Program Files\Manufacturer\Endpoint Agent"

                    If (Test-Path -Path $logfile1) { $edpa1 = "True" } Else { $edpa1 = "False" }
                    If (Test-Path -Path $logfile2) { $edpa2 = "True" } Else { $edpa2 = "False" }

                    # Check Services
                    $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp
                    If ($StatusE -like "*running*") { $ServiceStateE = "Running" } ElseIf ($StatusE -like "*Stopped*") { $ServiceStateE = "Stopped" } ElseIf ($StatusE -eq $null) { $ServiceStateE = "False" }
                    If ($StatusW -like "*running*") { $ServiceStateW = "Running" } ElseIf ($StatusW -like "*Stopped*") { $ServiceStateW = "Stopped" } ElseIf ($StatusW -eq $null) { $ServiceStateW = "False" }

                    [PSCustomObject]@{
                        Machine   = $env:COMPUTERNAME
                        Version        = $Version
                        'vfsmfd.sys'   = $vfsmfd
                        'vnwcd.sys'    = $vnwcd
                        'vrtam.sys'    = $vrtam
                        'edpa0.log'    = $edpa1
                        'edpa_ext0.log'= $edpa2
                        'EDPA Service' = $ServiceStateE
                        'WDP Service'  = $ServiceStateW
                        Status         = 'Success'
                    }
                }
                $results = Invoke-Command -Session $MySession -ScriptBlock $MyCommands4
                $results | Select-Object Machine, Status, Version, vfsmfd.sys, vnwcd.sys, vrtam.sys, 'edpa0.log', 'edpa_ext0.log', 'EDPA Service', 'WDP Service' | Format-Table -AutoSize
                #endregion Checking status after uninstallation

                #region Restart
                $Input_restart = Read-Host "Server need to restart after uninstallation. Do you want to restart the server now? (y/n)"
                switch ($Input_restart)
                {
                    'y'
                    {
                        Write-Host "Restarting $Machine : " -NoNewline
                        Try { Restart-Computer -ComputerName $Machine -Credential $cred -Force -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                        Catch { Write-Warning ($_); Continue }
                        Finally { $Error.Clear() }
                    }

                    'n'
                    {
                        Continue
                    }

                    Default { Write-Warning "Invalid Input" }

                }
                #endregion Restart
            }
        }

        'n' 
        { 
            Continue
        }

        Default { Write-Warning "Invalid Input" }
    }
}

Function Function_Six
{
$Input_restart = Read-Host "Reinstall DLP Agent Now? (y/n)"
 switch ($Input_restart) 
 {
    'y'
     {
        Foreach ($Machine in $Machinelist)
        {
            Get-PSSession | Remove-PSSession
            Write-Host "`n"
            Write-Host "$machine" -ForegroundColor Yellow

            #Checking OS Architecture
            Write-Host "Checking OS Architecture : " -NoNewline
            Try { $OS = (Get-WmiObject Win32_OperatingSystem -ComputerName $machine -Credential $Cred).OsArchitecture; Write-Host "$OS" -ForegroundColor Green }
            Catch { Write-Warning ($_); Continue }
            Finally { $Error.Clear() }

            #Transferring clean_agent.exe
            Write-Host "Transferring clean_agent.exe to C:\Patches : " -NoNewline
            $Path = "\\$machine\c$\Patches\"
            $Source64 = "$ScriptDir\x64\clean_agent.exe"
            $Source32 = "$ScriptDir\x86\clean_agent.exe"
            $Location = "\\kulbscvmfs03\IT\98 Software Installer\99-Other\DLP Agents\DLP\15.1 MP2 - Service Shutdown & Uninstall\service_shutdown.exe"
            If ($OS -eq "64-bit")
            {
               If (Test-Path $Path)
               {
                    Try 
                    { 
                        Copy-Item $Source64 -Destination $Path -Force -Recurse -ErrorAction Stop | Out-Null; Write-Host "SUCCESS" -ForegroundColor Green
                    }
                    Catch { Write-Warning ($_); Continue }
                    Finally { $Error.Clear() }

                    Try { Copy-Item $Location -Destination "\\$machine\c$\Program Files\Manufacturer\Endpoint Agent\" -Force -ErrorAction Stop} Catch {}
                    Try { Copy-Item $Location -Destination "\\$machine\c$\Program Files (x86)\Manufacturer\Endpoint Agent\" -Force -ErrorAction Stop} Catch {}
                    
                    #Establishing connection to machine
                    Write-Host "Establishing remote connection to $machine : " -NoNewline
                    Try { $MySession2 = New-PSSession -ComputerName $machine -Credential $Cred -ErrorAction Stop; Write-Host "SUCCESS" -ForegroundColor Green }
                    Catch { Write-Host "FAIL" -ForegroundColor Red; Write-Host "`n"; Continue }
                    Finally { $Error.Clear() }

                    $MyCommands =
                    {
                        Function Install_Agent
                        {
                            Write-Host "Executing install_agent.bat : " -NoNewline
                            cd "C:\Patches"
                            cmd.exe /c "install_agent.bat" hg st 2>&1 | Out-Null

                            $Path1 = "C:\Windows\System32\drivers\vfsmfd.sys"; $Path2 = "C:\Windows\System32\drivers\vnwcd.sys"; $Path3 = "C:\Windows\System32\drivers\vrtam.sys"
                            $logfile1 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa0.log"; $logfile2 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa_ext0.log"; $Folder1 = "C:\Program Files\Manufacturer\Endpoint Agent"
                            $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp

                            If ($Edpa -eq "Running") { $Edparun }
                            If ($Wdp -eq "Running") { $Wdprun }

                            If ((Test-Path $Path1) -and (Test-Path $Path2) -and (Test-Path $Path3) -and (Test-Path $logfile1) -and (Test-Path $logfile1) )
                            {
                                Write-Host "Installation completed" -ForegroundColor Green
                            }

                            Else 
                            {
                                Write-Host "Installation is not completed. Please rename 'Group Policy folder and try again'" -ForegroundColor Red
                            }
                        }

                        Function Uninstall_Agent
                        {
                            Write-Host "Stopping EDPA and WDP Service : " -NoNewline
                            cd "C:\Program Files\Manufacturer\Endpoint Agent"
                            cmd.exe /c "service_shutdown.exe -p=Welcome1" hg st *> $null

                            $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State
                            $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State

                            If (($Edpa -eq "Stopped") -and ($Wdp -eq "Stopped") )
                            {
                            Write-Host "SUCCESS" -ForegroundColor Green
                            }

                            Else 
                            {
                                Write-Host "Stop Service failed but the force uninstallation will continue anyway" -ForegroundColor Green
                            }
                        
                            Write-Host "Executing clean_agent.exe : " -NoNewline
                            cd "C:\Patches"
                            cmd.exe /c "echo y | clean_agent.exe" hg st 2>&1 | Out-Null

                            $Edpa = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp
                            $Ice = "C:\Program Files\Manufacturer\Endpoint Agent\ICE.exe"

                            If ((Test-Path $Ice) -or ($StatusE) )
                            {
                                Write-Host "Uninstallation still not completed" -ForegroundColor Red
                            }

                            If ((!(Test-Path $Ice)) -and ($StatusE -eq $null) -and ($StatusW -eq $null) ) 
                            {
                                Write-Host "Uninstallation is completed" -ForegroundColor Green
                                Install_Agent
                            }
                        }

                        Write-Host "Stopping EDPA and WDP Service : " -NoNewline
                        Try { cd "C:\Program Files\Manufacturer\Endpoint Agent" -ErrorAction Stop} Catch { }  
                        cmd.exe /c "service_shutdown.exe -p=Welcome1" hg st 2>&1 | Out-Null
                        Try { cd "C:\Program Files (x86)\Manufacturer\Endpoint Agent" -ErrorAction Stop} Catch { }
                        cmd.exe /c "service_shutdown.exe -p=Welcome1" hg st 2>&1 | Out-Null
     
                        $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State
                        $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State

                        If (($Edpa -eq "Stopped") -and ($Wdp -eq "Stopped") )
                        {
                            Write-Host "SUCCESS" -ForegroundColor Green
                        }

                        Else {
                            Write-Host "Stop Service failed or Service not found. But the force uninstallation will continue anyway" -ForegroundColor Green
                        }
                        
                        Write-Host "Executing clean_agent.exe : " -NoNewline
                        Try { cd "C:\Patches" -ErrorAction Stop } Catch { }
                        cmd.exe /c "echo y | clean_agent.exe"

                        #Checking Installation Status by checking the services
                        $Edpa = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp
                        $Ice = "C:\Program Files\Manufacturer\Endpoint Agent\ICE.exe"
                        If ((Test-Path $Ice) -or ($StatusE) )
                        {
                            Write-Host "Uninstallation is not complete. Please restart the server and try again option (6)" -ForegroundColor Red
                            Continue

                        }

                        If ((!(Test-Path $Ice)) -and ($StatusE -eq $null) -and ($StatusW -eq $null) ) 
                        {
                            Write-Host "Uninstallation is completed" -ForegroundColor Green
                            Install_Agent
                        }
                    }
                    Invoke-Command -Session $MySession2 -ScriptBlock $MyCommands

                }

               Elseif (!(Test-Path $Path))
               {
                    Try 
                    { 
                        New-Item -ItemType Directory $Path -Force -ErrorAction SilentlyContinue | Out-Null
                        Copy-Item $Source64 -Destination $Path -Force -Recurse -ErrorAction Stop; Write-Host "SUCCESS" -ForegroundColor Green
                    }
                    Catch { Write-Warning ($_); Continue }
                    Finally { $Error.Clear() }

                    Copy-Item $Location -Destination "\\$machine\c$\Program Files\Manufacturer\Endpoint Agent\" -Force -ErrorAction SilentlyContinue
                    #Copy-Item $Location -Destination "\\$machine\c$\Program Files (x86)\Manufacturer\Endpoint Agent\" -Force -ErrorAction SilentlyContinue| Out-Null
                    
                    #Establishing connection to machine
                    Write-Host "Establishing remote connection to $machine : " -NoNewline
                    Try { $MySession2 = New-PSSession -ComputerName $machine -Credential $Cred -ErrorAction Stop; Write-Host "SUCCESS" -ForegroundColor Green }
                    Catch { Write-Host "FAIL" -ForegroundColor Red; Write-Host "`n"; Continue }
                    Finally { $Error.Clear() }

                    $MyCommands =
                    {
                        Function Install_Agent
                        {
                            Write-Host "Executing install_agent.bat : " -NoNewline
                            cd "C:\Patches"
                            cmd.exe /c "install_agent.bat" hg st 2>&1 | Out-Null

                            $Path1 = "C:\Windows\System32\drivers\vfsmfd.sys"; $Path2 = "C:\Windows\System32\drivers\vnwcd.sys"; $Path3 = "C:\Windows\System32\drivers\vrtam.sys"
                            $logfile1 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa0.log"; $logfile2 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa_ext0.log"; $Folder1 = "C:\Program Files\Manufacturer\Endpoint Agent"
                            $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp

                            If ($Edpa -eq "Running") { $Edparun }
                            If ($Wdp -eq "Running") { $Wdprun }

                            If ((Test-Path $Path1) -and (Test-Path $Path2) -and (Test-Path $Path3) -and (Test-Path $logfile1) -and (Test-Path $logfile1) )
                            {
                                Write-Host "Installation completed" -ForegroundColor Green
                            }

                            Else 
                            {
                                Write-Host "Installation is not completed. Please generate the report using option 8" -ForegroundColor Red
                            }
                        }

                        Function Uninstall_Agent
                        {
                            Write-Host "Stopping EDPA and WDP Service : " -NoNewline
                            cd "C:\Program Files\Manufacturer\Endpoint Agent"
                            cmd.exe /c "service_shutdown.exe -p=Welcome1" hg st 2>&1 | Out-Null

                            $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State
                            $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State

                            If (($Edpa -eq "Stopped") -and ($Wdp -eq "Stopped") )
                            {
                            Write-Host "SUCCESS" -ForegroundColor Green
                            }

                            Else 
                            {
                                Write-Host "Stop Service failed but the force uninstallation will continue anyway" -ForegroundColor Green
                            }
                        
                            Write-Host "Executing clean_agent.exe : " -NoNewline
                            cd "C:\Patches"
                            cmd.exe /c "echo y | clean_agent.exe" hg st 2>&1 | Out-Null

                            $Edpa = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp
                            $Ice = "C:\Program Files\Manufacturer\Endpoint Agent\ICE.exe"

                            If ((Test-Path $Ice) -or ($StatusE) )
                            {
                                Write-Host "Uninstallation still not completed. Please REBOOT and try again with option (6)" -ForegroundColor Red
                            }

                            If ((!(Test-Path $Ice)) -and ($StatusE -eq $null) -and ($StatusW -eq $null) ) 
                            {
                                Write-Host "Uninstallation is completed" -ForegroundColor Green
                                Install_Agent
                            }
                        }

                        Write-Host "Stopping EDPA and WDP Service : " -NoNewline
                        cd "C:\Program Files\Manufacturer\Endpoint Agent"
                        cmd.exe /c "service_shutdown.exe -p=Welcome1" hg st 2>&1 | Out-Null
     
                        $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State
                        $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State

                        If (($Edpa -eq "Stopped") -and ($Wdp -eq "Stopped") )
                        {
                            Write-Host "SUCCESS" -ForegroundColor Green
                        }

                        Else {
                            Write-Host "Stop Service failed but the force uninstallation will continue anyway" -ForegroundColor Green
                        }
                        
                        Write-Host "Executing clean_agent.exe : " -NoNewline
                        cd "C:\Patches"
                        cmd.exe /c "echo y | clean_agent.exe" hg st 2>&1 | Out-Null

                        #Checking Installation Status by checking the services
                        $Edpa = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service  | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp
                        $Ice = "C:\Program Files\Manufacturer\Endpoint Agent\ICE.exe"
                        If ((Test-Path $Ice) -or ($StatusE) )
                        {
                            Write-Host "Uninstallation is not complete. Please restart the server and try again option (6)" -ForegroundColor Red
                            Continue
                        }

                        If ((!(Test-Path $Ice)) -and ($StatusE -eq $null) -and ($StatusW -eq $null) ) 
                        {
                            Write-Host "Uninstallation is completed" -ForegroundColor Green
                            Install_Agent
                        }
                    }
                    Invoke-Command -Session $MySession2 -ScriptBlock $MyCommands

                }
            }
        }
        Write-Host "`n"
     }

     'n' 
     { 
        Continue
     }
        Default { Write-Warning "Invalid Input" }
}
<#
    Foreach ($Server in $machinelist)
    {
        Write-Host "$Server" -ForegroundColor Yellow
        Write-Host "Establishing remote connection to $Server : " -NoNewline
        Try { $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        Catch { Write-Host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
        Finally { $Error.Clear() }
    
        $MyCommands = 
        {
            Write-Host "Finding Product GUID : " -NoNewline
            function Get-InstalledSoftware 
            {
    
            
            [OutputType([System.Management.Automation.PSObject])]
            [CmdletBinding()]
            param (
                [Parameter()]
                [ValidateNotNullOrEmpty()]
                [string]$Name
            )
 
            $UninstallKeys = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            $null = New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS
            $UninstallKeys += Get-ChildItem HKU: -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object { "HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall" }
            if (-not $UninstallKeys) {
                Write-Verbose -Message 'No software registry keys found'
            } else {
                foreach ($UninstallKey in $UninstallKeys) {
                    if ($PSBoundParameters.ContainsKey('Name')) {
                        $WhereBlock = { ($_.PSChildName -match '^{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$') -and ($_.GetValue('DisplayName') -like "$Name*") }
                    } else {
                        $WhereBlock = { ($_.PSChildName -match '^{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$') -and ($_.GetValue('DisplayName')) }
                    }
                    $gciParams = @{
                        Path        = $UninstallKey
                        ErrorAction = 'SilentlyContinue'
                    }
                    $selectProperties = @(
                        @{n='GUID'; e={$_.PSChildName}}, 
                        @{n='Name'; e={$_.GetValue('DisplayName')}}
                    )
                    Get-ChildItem @gciParams | Where $WhereBlock | Select-Object -Property $selectProperties
                }
                }
            }
            $ProductGUID = (Get-InstalledSoftware -Name 'AgentInstall').GUID
            $GUID = $ProductGUID

            If ($GUID)
            {
                Write-Host "$GUID" -ForegroundColor Green
            }

            ElseIf ($GUID -eq $null)
            {
                Write-Host "GUID not found. The software is not exist in current server" -ForegroundColor Red; Continue
            }


            #Write-Host "$GUID" -ForegroundColor Green
            $Path1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $Installed1 = Get-ChildItem -Path $Path1 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
            $String1 = ($Installed1).UninstallString
            $Path2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            $Installed2 = Get-ChildItem -Path $Path2 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
            $String2 = ($Installed2).UninstallString
            $Path3 = "C:\Program Files\Manufacturer"; $Path4 = "C:\Program Files (x86)\Manufacturer"
            $file1 = "C:\Windows\System32\vfsmfd.sys"; $file2 = "C:\Windows\System32\vnwcd.sys"; $file3 = "C:\Windows\System32\vrtam.sys"
            $reg1 = "HKLM:\SYSTEM\CurrentControlSet\Services\EDPA"; $reg2 = "HKLM:\SYSTEM\CurrentControlSet\Services\WDP"
            $regleftover1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$GUID"; $regleftover2 = "HKLM:\SOFTWARE\Wow6432Node\AdventNet\DesktopCentral\DCAgent\Inventory\System\Software\Wow64Node\Uninstall\$GUID"
        
           
            Write-Host "Finding UninstallString from registry : " -NoNewline

        If ($String1) 
        { 
            Write-Host "$String1" -ForegroundColor Green
            Write-Host "Uninstalling DLP : " -NoNewline
            cmd.exe /c "$String1 /q UNINSTALLPASSWORD=Welcome1" | Out-Null
            Write-Host "Done" -ForegroundColor Green
        }

        ElseIf ($String2) 
        { 
            Write-Host "$String2" -ForegroundColor Green
            Write-Host "Uninstalling DLP : " -NoNewline
            cmd.exe /c "$String2 /q UNINSTALLPASSWORD=Welcome1" | Out-Null
            Write-Host "Done" -ForegroundColor Green
        }

        Else 
        { 
            Write-Warning "Unable to find Uninstall String" 
        }
        #Write-Host "Remove leftover from folder & registry : " -NoNewline
        #Remove-Item -LiteralPath $Path3 -Force -Recurse -ErrorAction SilentlyContinue; Remove-Item -LiteralPath "$Path4" -Force -Recurse -ErrorAction SilentlyContinue
        #Remove-Item -Path $file1 -Force -ErrorAction SilentlyContinue; Remove-Item -Path $file2 -Force -ErrorAction SilentlyContinue; Remove-Item -Path $file3 -Force -ErrorAction SilentlyContinue
        #Remove-Item -Path $reg1 -Force -ErrorAction SilentlyContinue; Remove-Item -Path $reg2 -Force -ErrorAction SilentlyContinue
        #Remove-Item -Path $regleftover1 -Force -ErrorAction SilentlyContinue; Remove-Item -Path $regleftover2 -Force -ErrorAction SilentlyContinue
        #Write-Host "Done" -ForegroundColor Green
 
    }
    Invoke-Command -Session $MySession -ScriptBlock $MyCommands -ErrorAction Stop
    Write-Host "`n"
    }
#> 
}

Function Function_Seven
{
$Input_restart = Read-Host "Install DLP Agent Now? (y/n)"
    switch ($Input_restart) 
    {
        'y'
        {
            $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
            Foreach ($Machine in $Machinelist)
            {
                Write-Host "`n"
                #region Check PSSession Connection
                Write-Host "Establishing remote connection to $machine : "  -NoNewline
                Try { $MySession = New-PSSession -ComputerName $machine -Credential $Cred -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                Catch { Write-Host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
                #endregion Check PSSession Connection

                #region Executing clean_agent.exe
                Write-Host "Executing clean_agent.exe : " -NoNewline

                Try { $OS = (Get-WmiObject Win32_OperatingSystem -ErrorAction Stop).OsArchitecture } Catch { } Finally { $Error.Clear() }
                
                $Path = "\\$machine\c$\Patches\"
                $Source64 = "$ScriptDir\x64\clean_agent.exe"
                $Source32 = "$ScriptDir\x86\clean_agent.exe"

                If ($OS -eq "64-bit") { Try { Copy-Item $Source64 -Destination $Path -Force -Recurse -ErrorAction Stop } Catch { } }
                ElseIf ($OS -eq "32-bit") { Try { Copy-Item $Source32 -Destination $Path -Force -Recurse -ErrorAction Stop } Catch { } }

                $MyCommands2 =
                {
                    Try { cd "C:\Patches" -ErrorAction Stop } Catch { }
                    cmd.exe /c "echo y | clean_agent.exe -p=Welcome1" #> $null 2>&1
                }

                Invoke-Command -Session $MySession -ScriptBlock $MyCommands2
                #endregion Executing clean_agent.exe

                #region Checking status after uninstallation
                Write-Host "Checking Uninstallation Status : " -NoNewline
                $MyCommands4 =
                {

                    $PathVersion = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                    $PathVersion2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                    $Installed1 = Get-ChildItem -Path $PathVersion | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
                    $Installed2 = Get-ChildItem -Path $PathVersion2 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }

                    If ($Installed1) { $Version = ($Installed1).Displayversion }
                    ElseIf ($Installed2) { $Version = ($Installed2).Displayversion }
                    Else { $Version = "False" }

                    # Check vfsmfd.sys, vnwcd.sys and vrtam.sys
                    $Path1 = "C:\Windows\System32\drivers\vfsmfd.sys"; $Path2 = "C:\Windows\System32\drivers\vnwcd.sys"; $Path3 = "C:\Windows\System32\drivers\vrtam.sys";
                    If (Test-Path -Path $Path1) { $vfsmfd = "True" } Else { $vfsmfd = "False" }
                    If (Test-Path -Path $Path2) { $vnwcd = "True" } Else { $vnwcd = "False" }
                    If (Test-Path -Path $Path3) { $vrtam = "True" } Else { $vrtam = "False" }

                    # Check log file
                    $logfile1 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa0.log"; $logfile2 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa_ext0.log"; $Folder1 = "C:\Program Files\Manufacturer\Endpoint Agent"

                    If (Test-Path -Path $logfile1) { $edpa1 = "True" } Else { $edpa1 = "False" }
                    If (Test-Path -Path $logfile2) { $edpa2 = "True" } Else { $edpa2 = "False" }

                    # Check Services
                    $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp
                    If ($StatusE -like "*running*") { $ServiceStateE = "Running" } ElseIf ($StatusE -like "*Stopped*") { $ServiceStateE = "Stopped" } ElseIf ($StatusE -eq $null) { $ServiceStateE = "False" }
                    If ($StatusW -like "*running*") { $ServiceStateW = "Running" } ElseIf ($StatusW -like "*Stopped*") { $ServiceStateW = "Stopped" } ElseIf ($StatusW -eq $null) { $ServiceStateW = "False" }

                    [PSCustomObject]@{
                        Machine   = $env:COMPUTERNAME
                        Version        = $Version
                        'vfsmfd.sys'   = $vfsmfd
                        'vnwcd.sys'    = $vnwcd
                        'vrtam.sys'    = $vrtam
                        'edpa0.log'    = $edpa1
                        'edpa_ext0.log'= $edpa2
                        'EDPA Service' = $ServiceStateE
                        'WDP Service'  = $ServiceStateW
                        Status         = 'Success'
                    }
                }
                $results = Invoke-Command -Session $MySession -ScriptBlock $MyCommands4
                $results | Select-Object Machine, Status, Version, vfsmfd.sys, vnwcd.sys, vrtam.sys, 'edpa0.log', 'edpa_ext0.log', 'EDPA Service', 'WDP Service' | Format-Table -AutoSize
                #endregion Checking status after uninstallation

                #region Installation
                $Input_restart = Read-Host "Do you want to Install DLP Agent Now? (y/n)"
                switch ($Input_restart)
                {
                    'y'
                    {
                        $MyCommands5 =
                        {
                            Write-Host "Executing install_agent.bat : " -NoNewline
                            Try { cd "C:\Patches" -ErrorAction Stop } Catch{} 
                            cmd.exe /c "install_agent.bat" > $null 2>&1
                            
                            $Path1 = "C:\Windows\System32\drivers\vfsmfd.sys"; $Path2 = "C:\Windows\System32\drivers\vnwcd.sys"; $Path3 = "C:\Windows\System32\drivers\vrtam.sys"
                            $logfile1 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa0.log"; $logfile2 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa_ext0.log"; $Folder1 = "C:\Program Files\Manufacturer\Endpoint Agent"
                            $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp

                            If ($Edpa -eq "Running") { $Edparun }
                            If ($Wdp -eq "Running") { $Wdprun }

                            If ((Test-Path $Path1) -and (Test-Path $Path2) -and (Test-Path $Path3) -and (Test-Path $logfile1) -and (Test-Path $logfile1) )
                            {
                                Write-Host "Installation completed" -ForegroundColor Green
                            }

                            Else 
                            {
                                Write-Host "Installation is not completed." -ForegroundColor Red
                            }
                        }

                        Invoke-Command -Session $MySession -ScriptBlock $MyCommands5
                    }

                    'n'
                    {
                        Continue
                    }

                    Default { Write-Warning "Invalid Input" }
                }


                #endregion Installation
            }
        }

        'n' 
        { 
            Continue
        }

        Default { Write-Warning "Invalid Input" }
    }

}

Function Function_Eight 
{
$Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
$Results = Foreach ($machine in $machinelist)
{
    Try
    {
        $Session = New-PSSession -ComputerName $machine -Credential $Cred -ErrorAction Stop
        $MyCommands = 
        {

            #Check Version
            $PathVersion = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $PathVersion2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            $Installed1 = Get-ChildItem -Path $PathVersion | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }
            $Installed2 = Get-ChildItem -Path $PathVersion2 | ForEach { Get-ItemProperty $_.PSPath } | Where-Object { ($_.DisplayName -like "*AgentInstall*") -and ($_.Publisher -like "*Symantec Corp*") }

            If ($Installed1) { $Version = ($Installed1).Displayversion }
            ElseIf ($Installed2) { $Version = ($Installed2).Displayversion }
            Else { $Version = "False" }

            # Check vfsmfd.sys, vnwcd.sys and vrtam.sys
            $Path1 = "C:\Windows\System32\drivers\vfsmfd.sys"; $Path2 = "C:\Windows\System32\drivers\vnwcd.sys"; $Path3 = "C:\Windows\System32\drivers\vrtam.sys";
            If (Test-Path -Path $Path1) { $vfsmfd = "True" } Else { $vfsmfd = "False" }
            If (Test-Path -Path $Path2) { $vnwcd = "True" } Else { $vnwcd = "False" }
            If (Test-Path -Path $Path3) { $vrtam = "True" } Else { $vrtam = "False" }

            # Check log file
            $logfile1 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa0.log"; $logfile2 = "C:\Program Files\Manufacturer\Endpoint Agent\edpa_ext0.log"; $Folder1 = "C:\Program Files\Manufacturer\Endpoint Agent"

            If (Test-Path -Path $logfile1) { $edpa1 = "True" } Else { $edpa1 = "False" }
            If (Test-Path -Path $logfile2) { $edpa2 = "True" } Else { $edpa2 = "False" }

            # Check Services
            $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp
            If ($StatusE -like "*running*") { $ServiceStateE = "Running" } ElseIf ($StatusE -like "*Stopped*") { $ServiceStateE = "Stopped" } ElseIf ($StatusE -eq $null) { $ServiceStateE = "False" }
            If ($StatusW -like "*running*") { $ServiceStateW = "Running" } ElseIf ($StatusW -like "*Stopped*") { $ServiceStateW = "Stopped" } ElseIf ($StatusW -eq $null) { $ServiceStateW = "False" }

            [PSCustomObject]@{
            Machine   = $env:COMPUTERNAME
            Version        = $Version
            'vfsmfd.sys'   = $vfsmfd
            'vnwcd.sys'    = $vnwcd
            'vrtam.sys'    = $vrtam
            'edpa0.log'    = $edpa1
            'edpa_ext0.log'= $edpa2
            'EDPA Service' = $ServiceStateE
            'WDP Service'  = $ServiceStateW
            Status         = 'Success'
            }
        }
        Invoke-Command -Session $Session -ScriptBlock $MyCommands
    }

    Catch 
    {
        [PSCustomObject]@{
        Machine         = $machine
        Version        = $null
        'vfsmfd.sys'   = $null
        'vnwcd.sys'    = $null
        'vrtam.sys'    = $null
        'edpa0.log'    = $null
        'edpa_ext0.log'= $null
        'EDPA Service' = $null
        'WDP Service'  = $null
        Status         = 'Fail'
        }
    }
}

#$results | Select-Object Machine, Status, Version, vfsmfd.sys, vnwcd.sys, vrtam.sys, 'edpa0.log', 'edpa_ext0.log', 'EDPA Service', 'WDP Service' | Format-Table -AutoSize

$ConditionalFormat =$(
New-ConditionalText -Text Fail -Range 'B:B' -BackgroundColor Red -ConditionalTextColor Black
New-ConditionalText -Text False -Range 'C:J' -BackgroundColor Red -ConditionalTextColor Black
New-ConditionalText -Text Stopped -Range 'I:J' -BackgroundColor Red -ConditionalTextColor Black
)

$results | Select-Object Machine, Status, Version, vfsmfd.sys, vnwcd.sys, vrtam.sys, 'edpa0.log', 'edpa_ext0.log', 'EDPA Service', 'WDP Service' | Export-Excel -Path "$ScriptDir\DLPStatus.xlsx" -AutoSize -TableName "DLPStatus" -WorksheetName "DLPStatus" -ConditionalFormat $ConditionalFormat -Show -Activate
Get-PSSession | Remove-PSSession
}

Function Function_Nine
{

    $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")

    ForEach ($Machine in $Machinelist)
    {
        $MultiSession = New-PSSession -ComputerName $machinelist -Credential $cred

        $MyCommands = 
        {
            $Edpa = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*EDPA*") }).State; $Wdp = (Get-WmiObject Win32_Service | Where-Object { ($_.Name -like "*WDP*") }).State; $StatusE = $Edpa; $StatusW = $Wdp
            If ($StatusE -like "*running*") { $ServiceStateE = "Running" } ElseIf ($StatusE -like "*Stopped*") { $ServiceStateE = "Stopped" } ElseIf ($StatusE -eq $null) { $ServiceStateE = "False" }
            If ($StatusW -like "*running*") { $ServiceStateW = "Running" } ElseIf ($StatusW -like "*Stopped*") { $ServiceStateW = "Stopped" } ElseIf ($StatusW -eq $null) { $ServiceStateW = "False" }

            [PSCustomObject]@{
            Machine         = $env:COMPUTERNAME
            'EDPA Service' = $ServiceStateE
            'WDP Service'  = $ServiceStateW
            }
        }
    }

        $results = Invoke-Command -Session $MultiSession -ScriptBlock $MyCommands
        $results | Select-Object Machine, 'EDPA Service', 'WDP Service' | Format-Table -AutoSize
    
#$results | Select-Object Machine, 'EDPA Service', 'WDP Service' | Export-Excel -Path "$ScriptDir\Verify DLP Services.xlsx" -AutoSize -TableName "DLPService" -WorksheetName "DLPService"

Get-PSSession | Remove-PSSession
}

Function Function_Ten
{
    $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
    Foreach ($Machine in $Machinelist)
    {
        Write-Host "`n"
        Write-Host "$Machine : " -NoNewline

        $Path1 = "\\$Machine\c$\Program Files\Manufacturer\Endpoint Agent\edpa0.log"
        $Path2 = "\\$Machine\c$\Program Files\Manufacturer\Endpoint Agent\edpa_ext0.log"
        $Path3 = "\\$Machine\c$\InstallAgent.log"
        $Path4 = "\\$Machine\c$\Patches\Clean_Agent.log"

        Try { New-Item -ItemType Directory -Path "$ScriptDir\Log\$Machine" -Force -ErrorAction Stop | Out-Null } Catch {  }
        Try { Copy-Item $Path1 -Destination "$ScriptDir\Log\$Machine\" -Force -Recurse -ErrorAction Stop } Catch {  }
        Try { Copy-Item $Path2 -Destination "$ScriptDir\Log\$Machine\" -Force -Recurse -ErrorAction Stop } Catch {  }
        Try { Copy-Item $Path3 -Destination "$ScriptDir\Log\$Machine\" -Force -Recurse -ErrorAction Stop } Catch {  }
        Try { Copy-Item $Path4 -Destination "$ScriptDir\Log\$Machine\" -Force -Recurse -ErrorAction Stop } Catch {  }
        Write-Host "Logs file has been saved in" $ScriptDir\Log\$Machine\ -ForegroundColor Green
    }

}

Function Function_Eleven
{
    Get-PSSession | Remove-PSSession
    $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
    Foreach ($Machine in $machinelist)
    {
        $Machine
    }
    Write-Host "`n"
}

Function Function_Twelve
{
    $Input_restart = Read-Host "Restart Machine Now (y/n)"
    switch ($Input_restart) 
    {
        'y'
         {
            $Machinelist = @(get-content -Path "$ScriptDir\Machinelist.txt")
            Foreach ($Machine in $machinelist)
                {
                    Write-Host "Restarting $Machine : " -NoNewline 
                    Try { Restart-Computer -ComputerName $Machine -Credential $cred -Force -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                    Catch { Write-Warning ($_); Continue }
                    Finally { $Error.Clear() }
                }
            }

        'n' 
        { 
            Continue
        }
        Default { Write-Warning "Invalid Input" }
    }
}

#endregion Functions

#region Menu
function Show-Menu {
param ( [string]$Title = 'Menu' )
Clear-Host
Write-Host "Press 0 to Load Machine(s) List"
Write-Host "`n"									
Write-Host " [1] Ping"
Write-Host " [2] Check DLP Version"
Write-Host " [3] Copy Installer"
Write-Host " [4] Stop DLP Services"
Write-Host " [5] Uninstall DLP Agent"
#Write-Host " [6] Reinstall DLP Agent"
Write-Host " [7] Install DLP Agent"
Write-Host " [8] Verify DLP Installation Status"
Write-Host " [9] Verify DLP Services"
Write-Host " [10] Get edpa0.log, edpa_ext0.log, InstallAgent.log and clean_agent.log"
Write-Host " [11] Check List of Machine(s)"
Write-Host " [12] Restart Machine(s)"
Write-Host "`n"
Write-Host " [Q] Exit"
Write-Host "`n"
}

Do {
Show-Menu
Write-Host "Please make a selection: " -ForegroundColor Yellow -NoNewline
$input = Read-Host
Write-Host "`n"
switch ($input)
{
    '0' {Function_Zero}
    '1' {Function_One}
    '2' {Function_Two}
    '3' {Function_Three}
    '4' {Function_Four}
    '5' {Function_Five}
    '6' {Write-Host "DISABLE!"}
    '7' {Function_Seven}
    '8' {Function_Eight}
    '9' {Function_Nine}
    '10' {Function_Ten}
    '11' {Function_Eleven}
    '12' {Function_Twelve}
    'q' {Write-Host "The script has been canceled" -BackgroundColor Red -ForegroundColor White}
    Default {Write-Host "Your selection = $input, is not valid. Please try again." -BackgroundColor Red -ForegroundColor White}
}

pause
}
until ($input -eq 'q')
Get-PSSession | Remove-PSSession
#endregion Menu
