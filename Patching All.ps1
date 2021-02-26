Clear-Host
Get-PSSession | Remove-PSSession
Remove-Variable * -ErrorAction SilentlyContinue; $Error.Clear()
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$Username = 'suhail_asrulsani-ops'
$Password = 'Welcome12'
$Usernamerge = 'suhail_asrulsani-ops'
$Passwordrge = 'Welcome1'
$Passexchange = ConvertTo-SecureString -AsPlainText -Force -String "Welcome12"
$Domain = 'globalnet'
$dt = (Get-Date).ToString("ddMMyyyy_HHmmss") 
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "globalnet\suhail_asrulsani-ops",$Passexchange
$fromaddress = "suhail_asrulsani@averis.biz"
Remove-Item -Path "$ScriptDir\Report\ExchangeReport.html" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$ScriptDir\Report\KBEventViewer.xlsx" -Force -ErrorAction SilentlyContinue
#############################################################################################
### Functions
#############################################################################################

###############################################

###############################################
Function Function_Zero 
{
    notepad.exe "$ScriptDir\serverlist.txt"
    Get-PSSession | Remove-PSSession
}
###############################################
Function Function_ZeroZero 
{
    notepad.exe "$ScriptDir\kblist.txt"
    Get-PSSession | Remove-PSSession
}
###############################################
Function Function_One 
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n"   
    Foreach ($Server in $Serverlist)
    { 
        Write-Host $Server ": " -NoNewline -ForegroundColor Gray
        if(Test-Connection $Server -Count 1 -Quiet){ 
        Write-Host "Online" -ForegroundColor Green 
    }

    else
    {
        Write-Host "Offline" -ForegroundColor Red 
    }

    }
}
###############################################
Function Function_Two 
{
Clear-Host
    Function Connect-Mstsc {

    [cmdletbinding(SupportsShouldProcess,DefaultParametersetName='UserPassword')]
    param (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [Alias('CN')]
            [string[]]     $ComputerName,
        [Parameter(ParameterSetName='UserPassword',Mandatory=$true,Position=1)]
        [Alias('U')] 
            [string]       $User,
        [Parameter(ParameterSetName='UserPassword',Mandatory=$true,Position=2)]
        [Alias('P')] 
            [string]       $Password,
        [Parameter(ParameterSetName='Credential',Mandatory=$true,Position=1)]
        [Alias('C')]
            [PSCredential] $Credential,
        [Alias('A')]
            [switch]       $Admin,
        [Alias('MM')]
            [switch]       $MultiMon,
        [Alias('F')]
            [switch]       $FullScreen,
        [Alias('Pu')]
            [switch]       $Public,
        [Alias('W')]
            [int]          $Width,
        [Alias('H')]
            [int]          $Height,
        [Alias('WT')]
            [switch]       $Wait
    )

    begin {
        [string]$MstscArguments = ''
        switch ($true) {
            {$Admin}      {$MstscArguments += '/admin '}
            {$MultiMon}   {$MstscArguments += '/multimon '}
            {$FullScreen} {$MstscArguments += '/f '}
            {$Public}     {$MstscArguments += '/public '}
            {$Width}      {$MstscArguments += "/w:$Width "}
            {$Height}     {$MstscArguments += "/h:$Height "}
        }

        if ($Credential) {
            $User     = $Credential.UserName
            $Password = $Credential.GetNetworkCredential().Password
        }
    }
    process {
        foreach ($Computer in $ComputerName) {
            $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
            $Process = New-Object System.Diagnostics.Process
            
            # Remove the port number for CmdKey otherwise credentials are not entered correctly
            if ($Computer.Contains(':')) {
                $ComputerCmdkey = ($Computer -split ':')[0]
            } else {
                $ComputerCmdkey = $Computer
            }

            $ProcessInfo.FileName    = "$($env:SystemRoot)\system32\cmdkey.exe"
            $ProcessInfo.Arguments   = "/generic:TERMSRV/$ComputerCmdkey /user:$User /pass:$($Password)"
            $ProcessInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
            $Process.StartInfo = $ProcessInfo
            if ($PSCmdlet.ShouldProcess($ComputerCmdkey,'Adding credentials to store')) {
                [void]$Process.Start()
            }

            $ProcessInfo.FileName    = "$($env:SystemRoot)\system32\mstsc.exe"
            $ProcessInfo.Arguments   = "$MstscArguments /v $Computer"
            $ProcessInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal
            $Process.StartInfo       = $ProcessInfo
            if ($PSCmdlet.ShouldProcess($Computer,'Connecting mstsc')) {
                [void]$Process.Start()
                if ($Wait) {
                    $null = $Process.WaitForExit()
                }       
            }
        }
    }
}
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    Foreach ($Server in $Serverlist)
    {
        Connect-Mstsc $Server globalnet\$Username $Password -Admin 
    }
}
###############################################
Function Function_Three 
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n"
    Foreach ($Server in $Serverlist) 
    {
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green}
        catch { Write-Warning ($_); Write-Host "`n"; Continue }
        Finally { $Error.Clear() }
        $MyCommands = 
        {
            $custom_array = @()
            try { $Services = Get-WmiObject Win32_Service | Where-Object { ($_.startmode -like "*auto*") -and ` ($_.state -notlike "*running*") } }
            catch { Write-Warning ($_); Continue }
            Finally { $Error.Clear() }
            Foreach ($Service in $Services) 
            {
                $custom_array += New-Object PSObject -Property @{
                    Server_Name = $Server
                    Service_DisplayName = $Service.DisplayName
                    Service_Name = $Service.Name
                    Service_Startmode = $Service.StartMode
                    Service_State = $Service.State     
                }
            }

            $multiple_output = $custom_array
            Foreach ($output in $multiple_output)
            {
                Write-host "Starting Service" $output.Service_DisplayName ": " -NoNewline -ForegroundColor Gray
                try { Start-Service $output.Service_Name -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                catch { Write-Warning ($_); Continue }
                Finally { $Error.Clear() }
            }
        }
        Invoke-Command -Session $MySession -ScriptBlock $MyCommands
        Write-Host "`n"
    }
}
###############################################
Function Function_Four 
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n"   
    Foreach ($Server in $Serverlist) 
    {
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        catch { Write-Warning ($_); Write-Host "`n"; Continue }
        Finally { $Error.Clear() }
        $MyCommands = 
        {
            Get-WmiObject Win32_Service | Where-Object { ($_.startmode -like "*auto*") -and ` ($_.state -notlike "*running*") } | Select-Object DisplayName,Name,StartMode,State,PSComputerName|ft -AutoSize
        }
        Invoke-Command -Session $MySession -ScriptBlock $MyCommands
    }
}
###############################################
Function Function_Five 
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n"   
    Foreach ($Server in $Serverlist)
    {
        Write-Host $Server": " -NoNewline
        try {
        $reboot_time = Get-WmiObject win32_operatingsystem -ComputerName $Server -ErrorAction Stop  | ForEach-Object { $PSItem.ConvertToDateTime($PSItem.LastBootUpTime) }; Write-Host "$reboot_time" -ForegroundColor Green }
        catch { Write-Warning ($_); Continue }
        Finally { $Error.Clear() }
    }
}
###############################################
Function Function_Six 
{
Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    foreach ($server in $Serverlist) 
    {
        $session = ((quser /server:$server | ? { $_ -match $Username }) -split ' +')[2]
        logoff $session /server:$server
    }
}
###############################################
Function Function_Seven 
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n" 
    
    Foreach ($Server in $Serverlist)
    {
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        catch { Write-Warning ($_); Continue }
        Finally { $Error.Clear() }
        $MyCommands = 
        {
            $path = "C:\patches"
            if (!(test-path $path))
            {
                New-Item -ItemType Directory -Force -Path $path | Out-Null
            }
        }
        Invoke-Command -Session $MySession -ScriptBlock $MyCommands
    }

    Foreach ($Server in  $Serverlist)
    {
        Write-Host "Opening C:\patches at $Server : " -NoNewline
        try { Invoke-Item \\$Server\c$\patches -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        catch { Write-Warning ($_); Continue }
        Finally { $Error.Clear() }
    }
    
}
###############################################
Function Function_Eight
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n" 

    Foreach ($Server in $Serverlist) 
    {
        $uri = "http://$Server.globalnet.lcl/PowerShell/"
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        catch { Write-Warning ($_); Write-Host "`n"; Continue }
        Finally { $Error.Clear() }   

        Invoke-Command -Session $MySession -ScriptBlock {
        Get-MailboxDatabaseCopyStatus
    } | Select-Object Name, Status, activationpreference | Format-Table -AutoSize
}
}
###############################################
Function Function_Nine
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n" 

    Foreach ($Server in $Serverlist) 
    {
        $uri = "http://$Server.globalnet.lcl/PowerShell/"
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        catch { Write-Warning ($_); Write-Host "`n"; Continue }
        Finally { $Error.Clear() }   

        Invoke-Command -Session $MySession -ScriptBlock {
        Test-ServiceHealth
    } | Select-Object PSComputerName, Role, RequiredServicesRunning, ServicesNotRunning, ServicesRunning | Format-List
}
}
###############################################
Function Function_Ten
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n" 

    Foreach ($Server in $Serverlist) 
    {
        $uri = "http://$Server.globalnet.lcl/PowerShell/"
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        catch { Write-host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
        Finally { $Error.Clear() }   

        Invoke-Command -Session $MySession -ScriptBlock {
        Test-ReplicationHealth
    }  | Select-Object Server, Check, Result | Format-Table -AutoSize
}
}
###############################################
Function Function_Eleven
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n" 

    Foreach ($Server in $Serverlist) 
    {
        $uri = "http://$Server.globalnet.lcl/PowerShell/"
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        catch { Write-host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
        Finally { $Error.Clear() }   

        Invoke-Command -Session $MySession -ScriptBlock { Test-Mailflow } | Select-Object PSComputerName, TestMailflowResult | Format-List
    }
}
###############################################
Function Function_Twelve 
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n" 
    Foreach ($Server in $Serverlist) 
    {
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        catch { Write-host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
        Finally { $Error.Clear() }
        $MyCommands = 
        {
            Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "*sql*" } | Select-Object DisplayName,Name,StartMode,State,PSComputerName|ft -AutoSize
        }
        Invoke-Command -Session $MySession -ScriptBlock $MyCommands
    }
}
###############################################
Function Function_Thirteen 
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n"
    Foreach ($Server in $Serverlist) 
    {
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green}
        catch { Write-host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
        Finally { $Error.Clear() }
        $MyCommands = 
        {
            $custom_array = @()
            try { $Services = Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "*sql*"  } }
            catch { Write-Warning ($_); Continue }
            Finally { $Error.Clear() }
            Foreach ($Service in $Services) 
            {
                $custom_array += New-Object PSObject -Property @{
                    Server_Name = $Server
                    Service_DisplayName = $Service.DisplayName
                    Service_Name = $Service.Name
                    Service_Startmode = $Service.StartMode
                    Service_State = $Service.State     
                }
            }

            $multiple_output = $custom_array
            Foreach ($output in $multiple_output)
            {
                Write-host "Stopping Service" $output.Service_DisplayName ": " -NoNewline -ForegroundColor Gray
                try { Stop-Service $output.Service_Name -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                catch { Write-host "Failed" -ForegroundColor Red; Continue }
                Finally { $Error.Clear() }

                Write-host "Changing Service" $output.Service_DisplayName "to Manual" ": " -NoNewline -ForegroundColor Gray
                try { Set-Service -Name $output.Service_Name -StartupType "Manual" -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                catch { Write-host "Failed" -ForegroundColor Red; Continue }
                Finally { $Error.Clear() }
            }
        }
        Invoke-Command -Session $MySession -ScriptBlock $MyCommands
        Write-Host "`n"
    }
}
###############################################
Function Function_Fourteen
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n"
    Foreach ($Server in $Serverlist) 
    {
        Get-PSSession | Remove-PSSession
        Write-Host "Establishing remote connection to" $Server ": " -NoNewline
        try { $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green}
        catch { Write-host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
        Finally { $Error.Clear() }
        $MyCommands = 
        {
            $custom_array = @()
            try { $Services = Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "*sql*" } }
            catch { Write-Warning ($_); Continue }
            Finally { $Error.Clear() }
            Foreach ($Service in $Services) 
            {
                $custom_array += New-Object PSObject -Property @{
                    Server_Name = $Server
                    Service_DisplayName = $Service.DisplayName
                    Service_Name = $Service.Name
                    Service_Startmode = $Service.StartMode
                    Service_State = $Service.State     
                }
            }

            $multiple_output = $custom_array
            Foreach ($output in $multiple_output)
            {
                Write-host "Starting Service" $output.Service_DisplayName ": " -NoNewline -ForegroundColor Gray
                try { Start-Service $output.Service_Name -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                catch { Write-host "Failed" -ForegroundColor Red; Continue }
                Finally { $Error.Clear() }

                Write-host "Changing Service" $output.Service_DisplayName "to Automatic" ": " -NoNewline -ForegroundColor Gray
                try { Set-Service -Name $output.Service_Name -StartupType "Automatic" -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
                catch { Write-host "Failed" -ForegroundColor Red; Continue }
                Finally { $Error.Clear() }

            }
        }
        Invoke-Command -Session $MySession -ScriptBlock $MyCommands
        Write-Host "`n"
    }
}
###############################################
Function Function_Fifteen
{
    Clear-Host
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n"

    Foreach ($Server in $Serverlist) 
    {
        Write-Host "Establishing remote connection to $Server : " -ForegroundColor Gray -NoNewline
        Try { $MySession = New-PSSession -ComputerName $Server -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
        Catch { Write-Host "Failed" -ForegroundColor Red; Write-Host "`n"; Continue }
        Finally { $Error.Clear() }
        
        $MyCommands   = 
        {
            Write-Host "Generating list : " -ForegroundColor Gray -NoNewline
            Try { $Result = Get-Hotfix -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green; $Result | Select-Object PSComputerName, HotfixID, Description, InstalledOn | Format-Table -AutoSize }
            Catch { Write-Host "Failed" -ForegroundColor Red; Continue }
            Finally { $Error.Clear() }

            Write-Host "Total KB Installed : " -ForegroundColor Gray -NoNewline
            $Count = @(Get-HotFix).Count
            Write-Host "$Count" -ForegroundColor Green
        }

        Invoke-Command -Session $MySession -ScriptBlock $MyCommands
        Write-Host "`n"     
    }
}
###############################################
Function Function_Sixteen
{
    Clear-Host
    $Kblist = @(get-content -Path "$ScriptDir\kblist.txt")
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $datetime = Get-Date -Format G
    Write-Host "Date and Time is: $datetime"
    Write-Host "`n"
    
    Foreach ($Server in $Serverlist)
    {
        Foreach ($kb in $Kblist)
        {
            Try { Get-HotFix -ComputerName $Server -ErrorAction Stop | Where-Object { $_.HotfixID -like "*$kb*" } | Select-Object PSComputerName, HotfixID, Description, InstalledOn | Format-Table -AutoSize }
            Catch { Write-host "Failed to get Kb from $Server" -ForegroundColor Red; Write-Host "`n"; Continue } 
            Finally { $Error.Clear() }
        }
    } 
}
###############################################
Function Function_Seventeen
{
    
    $Input_restart = Read-Host "Restart Server Now (y/n)"
    switch ($Input_restart) 
    {
        'y'
         {
            Foreach ($Server in $Serverlist)
                {
                    Write-Host "Restarting $Server : " -NoNewline 
                    Try { Restart-Computer -ComputerName $Server -Force -ErrorAction Stop; Write-Host "Done" -ForegroundColor Green }
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
#################################################
Function Function_Eighteen
{
    Function Connect-Mstsc {

    [cmdletbinding(SupportsShouldProcess,DefaultParametersetName='UserPassword')]
    param (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [Alias('CN')]
            [string[]]     $ComputerName,
        [Parameter(ParameterSetName='UserPassword',Mandatory=$true,Position=1)]
        [Alias('U')] 
            [string]       $User,
        [Parameter(ParameterSetName='UserPassword',Mandatory=$true,Position=2)]
        [Alias('P')] 
            [string]       $Password,
        [Parameter(ParameterSetName='Credential',Mandatory=$true,Position=1)]
        [Alias('C')]
            [PSCredential] $Credential,
        [Alias('A')]
            [switch]       $Admin,
        [Alias('MM')]
            [switch]       $MultiMon,
        [Alias('F')]
            [switch]       $FullScreen,
        [Alias('Pu')]
            [switch]       $Public,
        [Alias('W')]
            [int]          $Width,
        [Alias('H')]
            [int]          $Height,
        [Alias('WT')]
            [switch]       $Wait
    )

    begin {
        [string]$MstscArguments = ''
        switch ($true) {
            {$Admin}      {$MstscArguments += '/admin '}
            {$MultiMon}   {$MstscArguments += '/multimon '}
            {$FullScreen} {$MstscArguments += '/f '}
            {$Public}     {$MstscArguments += '/public '}
            {$Width}      {$MstscArguments += "/w:$Width "}
            {$Height}     {$MstscArguments += "/h:$Height "}
        }

        if ($Credential) {
            $User     = $Credential.UserName
            $Password = $Credential.GetNetworkCredential().Password
        }
    }
    process {
        foreach ($Computer in $ComputerName) {
            $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
            $Process = New-Object System.Diagnostics.Process
            
            # Remove the port number for CmdKey otherwise credentials are not entered correctly
            if ($Computer.Contains(':')) {
                $ComputerCmdkey = ($Computer -split ':')[0]
            } else {
                $ComputerCmdkey = $Computer
            }

            $ProcessInfo.FileName    = "$($env:SystemRoot)\system32\cmdkey.exe"
            $ProcessInfo.Arguments   = "/generic:TERMSRV/$ComputerCmdkey /user:$User /pass:$($Password)"
            $ProcessInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
            $Process.StartInfo = $ProcessInfo
            if ($PSCmdlet.ShouldProcess($ComputerCmdkey,'Adding credentials to store')) {
                [void]$Process.Start()
            }

            $ProcessInfo.FileName    = "$($env:SystemRoot)\system32\mstsc.exe"
            $ProcessInfo.Arguments   = "$MstscArguments /v $Computer"
            $ProcessInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal
            $Process.StartInfo       = $ProcessInfo
            if ($PSCmdlet.ShouldProcess($Computer,'Connecting mstsc')) {
                [void]$Process.Start()
                if ($Wait) {
                    $null = $Process.WaitForExit()
                }       
            }
        }
    }
}
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    Foreach ($Server in $Serverlist)
    {
        Connect-Mstsc $Server rge-extranet\$Usernamerge $Passwordrge -Admin 
    }
}
#################################################
Function Function_Twenty
{
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    Foreach ($Server in $Serverlist)
    {
        $Path = "\\$Server\c$\Patches\"
        $InstallUpdatesh = "$ScriptDir\File\InstallUpdatesh.vbs"
        Write-Host "$Server : Cleaning Up C:\Patches : " -NoNewline
        If (Test-Path $Path)
        {
            Try 
            { 
                Get-ChildItem -Path $Path -File -Recurse -ErrorAction Stop | Where-Object { $_Name -ne "InstallUpdatesh.vbs" } | Remove-Item -Force -Recurse
                Copy-Item $InstallUpdatesh -Destination $Path -Force -Recurse
                Write-Host "Done" -ForegroundColor Green
            }
                
            Catch 
            { 
                Write-Warning ($_); Continue 
            }

            Finally 
            { 
                $Error.Clear() 
            }
        }

        If (!(Test-Path $Path))
        {
            Try { New-Item -ItemType Directory -Force -Path $path -ErrorAction Stop | Out-Null; Copy-Item $InstallUpdatesh -Destination $Path -Force -Recurse; Write-Host "Done" -ForegroundColor Green }
            Catch 
            { 
                Write-Warning ($_); Continue 
            }

            Finally 
            { 
                $Error.Clear() 
            }
            
        }
    }
}
#################################################
Function Function_TwentyOne
{
    Get-PSSession | Remove-PSSession
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $Results = Foreach ($Server in $Serverlist)
    {
        $uri = "http://$Server.globalnet.lcl/PowerShell/"

        Try 
        {
            $MySession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop

            $MyCommands =
            {
                Get-MailboxDatabaseCopyStatus
            }
            Invoke-Command -Session $MySession -ScriptBlock $MyCommands
        }

        Catch
        {
            Write-Host "$Server : " -NoNewline
            Write-Warning ($_) 
        }

    }
    #$Results | Select-Object PSComputerName, Name, Status, activationpreference | Format-Table -AutoSize

    $Results2 = Foreach ($Server in $Serverlist)
    {
        $uri = "http://$Server.globalnet.lcl/PowerShell/"

        Try 
        {
            $MySession2 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop

            $MyCommands2 =
            {
                Test-ServiceHealth
            }
            Invoke-Command -Session $MySession2 -ScriptBlock $MyCommands2
        }

        Catch
        {
            Write-Host "$Server : " -NoNewline
            Write-Warning ($_) 
        }
    }
    #$Results2 | Select-Object PSComputerName, Role, RequiredServicesRunning | Format-Table -AutoSize

    $Results3 = Foreach ($Server in $Serverlist)
    {
        $uri = "http://$Server.globalnet.lcl/PowerShell/"

        Try 
        {
            $MySession3 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop

            $MyCommands3 =
            {
                Test-ReplicationHealth
            }
            Invoke-Command -Session $MySession3 -ScriptBlock $MyCommands3
        }

        Catch
        {
            Write-Host "$Server : " -NoNewline
            Write-Warning ($_) 
        }
    }
    #$Results3 | Select-Object Server, Check, Result | Format-Table -AutoSize

    $Results4 = Foreach ($Server in $Serverlist)
    {
        $uri = "http://$Server.globalnet.lcl/PowerShell/"

        Try 
        {
            $MySession4 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop

            $MyCommands4 =
            {
                Test-Mailflow
            }
            Invoke-Command -Session $MySession4 -ScriptBlock $MyCommands4
        }

        Catch
        {
            Write-Host "$Server : " -NoNewline
            Write-Warning ($_) 
        }
    }
    #$Results4 | Select-Object PSComputerName, TestMailflowResult | Format-Table -AutoSize
$Header1 = @"
<h1> Exchange Database </h1>
<style>
table 
{
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    font-size: 12px;
    #width: 100%;
}

td, th
{
	border: 1px solid #ddd;
	padding: 8px;
}

tr:nth-child(even) {background-color: #f2f2f2;}

tr:hover {background-color: #ddd;}

th
{
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #4CAF50;
    color:white;
}
</style>
"@

$Header2 = @"
<h1> Exchange Services </h1>
<style>
table 
{
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    #width: 100%;
}

td, th
{
	border: 1px solid #ddd;
	padding: 8px;
}

tr:nth-child(even) {background-color: #f2f2f2;}

tr:hover {background-color: #ddd;}

th
{
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #4CAF50;
    color:white;
}
</style>
"@

$Header3 = @"
<h1> Exchange Replication Health </h1>
<style>
table 
{
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    #width: 100%;
}

td, th
{
	border: 1px solid #ddd;
	padding: 8px;
}

tr:nth-child(even) {background-color: #f2f2f2;}

tr:hover {background-color: #ddd;}

th
{
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #4CAF50;
    color:white;
}
</style>
"@

$Header4 = @"
<h1> Exchange Mail Flow </h1>
<style>
table 
{
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    #width: 100%;
}

td, th
{
	border: 1px solid #ddd;
	padding: 8px;
}

tr:nth-child(even) {background-color: #f2f2f2;}

tr:hover {background-color: #ddd;}

th
{
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #4CAF50;
    color:white;
}
</style>
"@

    $Results | Select-Object PSComputerName, Name, Status, activationpreference | ConvertTo-Html -Title "ExchangeReport" -Head $Header1 | Out-File "$ScriptDir\Report\ExchangeReport.html" 
    $Results2 | Select-Object PSComputerName, Role, RequiredServicesRunning | ConvertTo-Html -Head $Header2 | Out-File "$ScriptDir\Report\ExchangeReport.html" -Append
    $Results3 | Select-Object Server, Check, Result | ConvertTo-Html -Head $Header3 | Out-File "$ScriptDir\Report\ExchangeReport.html"  -Append 
    $Results4 | Select-Object PSComputerName, TestMailflowResult | ConvertTo-Html -Head $Header4 | Out-File "$ScriptDir\Report\ExchangeReport.html"  -Append

$datetime = Get-Date -Format G
$toaddress = "suhail_asrulsani@averis.biz" #, jessey_ui@averis.biz, honwai_Lim@averis.biz, Magaret_Hansen@averis.biz"
#$bccaddress = ""
#$CCaddress = ""
$Subject = "POST PATCH - Exchange Health Report $datetime"
#$body = get-content .\ServiceReport.html -Raw
$body = [System.IO.File]::ReadAllText("$ScriptDir\Report\ExchangeReport.html")
#$attachment = "$ScriptDir\ServiceReport.html"
$smtpserver = "kulavrcasarr01.globalnet.lcl"

####################################

$message = new-object System.Net.Mail.MailMessage
$message.From = $fromaddress
$message.To.Add($toaddress)
#$message.CC.Add($CCaddress)
#$message.Bcc.Add($bccaddress)
$message.IsBodyHtml = $True
$message.Subject = $Subject
#$attach = new-object Net.Mail.Attachment($attachment)
#$message.Attachments.Add($attach)
$message.body = $body
$smtp = new-object Net.Mail.SmtpClient($smtpserver)
$smtp.Send($message)
}
#################################################
Function Function_TwentyTwo
{
    Clear-Host
    $Serverlist = @(get-content -Path "$ScriptDir\serverlist.txt")
    $Kblist = @(get-content -Path "$ScriptDir\kblist.txt")
    $Results = Foreach ($Server in $Serverlist)
    {
        Try
        {
            $Session = New-PSSession -ComputerName $Server -ErrorAction Stop
            Copy-Item "$ScriptDir\kblist.txt" -Destination "\\$Server\c$\temp\" -Force

            $MyCommands =
            {
                $patchlist = @(get-content -Path "C:\temp\kblist.txt")
                Foreach ($patch in $patchlist) 
                {
                    $TimeCreated = (Get-WinEvent -LogName Setup -ErrorAction Stop| Where-Object { ($_.Message -like "*successfully changed to the Installed*") -and ($_.Message -like "*$patch*") }).TimeCreated

                    If ($TimeCreated -eq $null)
                    {
                        $InstalledDate = 'N/A'
                    }

                    If ($TimeCreated)
                    {
                        $InstalledDate = $TimeCreated
                    }
                
                    [PSCustomObject]@{
                    Server = $env:COMPUTERNAME
                    Status = 'Success'
                    KB = $patch
                    InstalledDate = $InstalledDate
                    }
                }

            }

            Invoke-Command -Session $Session -ScriptBlock $MyCommands
        }

        Catch
        {
            [PSCustomObject]@{
            Server = $Server
            Status = 'Fail'
            KB = 'Fail'
            InstalledDate = 'Fail'
            }
        }

        Finally
        {
            $Error.Clear()
        }
    }  

$Results | Select-Object Server, Status, KB, InstalledDate | Format-Table -AutoSize

$ConditionalFormat =$(
New-ConditionalText -Text Fail -Range 'B:B' -BackgroundColor Red -ConditionalTextColor Black
New-ConditionalText -Text Fail -Range 'C:C' -BackgroundColor Red -ConditionalTextColor Black
New-ConditionalText -Text Fail -Range 'D:D' -BackgroundColor Red -ConditionalTextColor Black
)

$Results | Select-Object Server, Status, KB, InstalledDate | Export-Excel -Path "$ScriptDir\Report\KBEventViewer.xlsx" -AutoSize -TableName "KBInstallationStatus" -WorksheetName "KBInstallationStatus" -ConditionalFormat $ConditionalFormat -Show -Activate

}
#################################################
#############################################################################################
### Menu
#############################################################################################
function Show-Menu {
param ( [string]$Title = 'Menu' )
Clear-Host
Write-Host "Press 0 to Load Server List"
Write-Host "Press 00 to Load KB List"
Write-Host "`n"
Write-Host "GENERAL:"									
Write-Host " [1] Ping                                     [17] Restart"                         							        		        
Write-Host " [2] RDP                                      [19] Run InstallUpdatesh.vbs"
Write-Host " [3] Start Auto Services                      [20] Cleanup C:\Patches"                    
Write-Host " [4] List Auto Services"                 
Write-Host " [5] Uptime"
Write-Host " [6] Logoff"
Write-Host " [7] Explore C:\Patches"
Write-Host "`n"
Write-Host "EXCHANGE:                                     RGE:"					
Write-Host " [8] DB Preferred Node                        [18] RDP" 
Write-Host " [9] Service Health"
Write-Host "[10] Replication Health"
Write-Host "[11] Mailflow"
Write-Host "[21] Exchange Health Report"
Write-Host "`n"
Write-Host "DB:"
Write-Host "[12] List SQL Services"
Write-Host "[13] Stop/Manual SQL Services"
Write-Host "[14] Start/Auto SQL Services"
Write-Host "`n"
Write-Host "AUDIT:"
Write-Host "[15] List KB Installed"
Write-Host "[16] Check KB"
Write-Host "[22] Check KB Installed in Eventviewer"
Write-Host "`n"
}

Do {
Show-Menu
Write-Host "Please make a selection: " -ForegroundColor Yellow -NoNewline
$input = Read-Host
Write-Host "`n"
switch ($input)
{
    '0' {Function_Zero}    ### Input the name of the function you want to execute when 0 is entered
    '00' {Function_ZeroZero}    ### Input the name of the function you want to execute when 0 is entered
    '1' {Function_One}    ### Input the name of the function you want to execute when 1 is entered
    '2' {Function_Two}    ### Input the name of the function you want to execute when 2 is entered
    '3' {Function_Three}  ### Input the name of the function you want to execute when 3 is entered
    '4' {Function_Four}   ### Input the name of the function you want to execute when 4 is entered
    '5' {Function_Five}   ### Input the name of the function you want to execute when 5 is entered
    '6' {Function_Six}   ### Input the name of the function you want to execute when 5 is entered
    '7' {Function_Seven}   ### Input the name of the function you want to execute when 5 is entered
    '8' {Function_Eight}   ### Input the name of the function you want to execute when 5 is entered
    '9' {Function_Nine}   ### Input the name of the function you want to execute when 5 is entered
    '10' {Function_Ten}   ### Input the name of the function you want to execute when 5 is entered
    '11' {Function_Eleven}   ### Input the name of the function you want to execute when 5 is entered
    '12' {Function_Twelve}   ### Input the name of the function you want to execute when 5 is entered
    '13' {Function_Thirteen}   ### Input the name of the function you want to execute when 5 is entered
    '14' {Function_Fourteen}   ### Input the name of the function you want to execute when 5 is entered
    '15' {Function_Fifteen}   ### Input the name of the function you want to execute when 5 is entered
    '16' {Function_Sixteen}   ### Input the name of the function you want to execute when 5 is entered
    '17' {Function_Seventeen}   ### Input the name of the function you want to execute when 5 is entered
    '18' {Function_Eighteen}   ### Input the name of the function you want to execute when 5 is entered
    '19' {Function_Nineteen}   ### Input the name of the function you want to execute when 5 is entered
    '20' {Function_Twenty}   ### Input the name of the function you want to execute when 5 is entered
    '21' {Function_TwentyOne}   ### Input the name of the function you want to execute when 5 is entered
    '22' {Function_TwentyTwo}   ### Input the name of the function you want to execute when 5 is entered
    'Q' {Write-Host "The script has been canceled" -BackgroundColor Red -ForegroundColor White}
    Default {Write-Host "Your selection = $input, is not valid. Please try again." -BackgroundColor Red -ForegroundColor White}
}
pause
}
until ($input -eq 'q')

Get-PSSession | Remove-PSSession
