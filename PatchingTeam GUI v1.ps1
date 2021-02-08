$domain = 'globalnet'
$user = 'honwai_lim-ops'
$password = 'Welcome3'
 
$Usernamerge = 'rge-extranet\honwai_lim-ops'
$Passwordrge = 'Welcome1'
$pass = ConvertTo-SecureString -AsPlainText $Passwordrge -Force
$SecureString = $pass
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Usernamerge, $SecureString
 
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
 
Add-Type -AssemblyName System.Windows.Forms
 
$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Patching Team"
$Form.TopMost = $true
$Form.Width = 1000
$Form.Height = 1000
$Form.FormBorderStyle = "Fixed3D"
$form.StartPosition = "centerScreen"
$form.ShowInTaskbar = $true
 
# Label
$label3 = New-Object system.windows.Forms.Label
$label3.Text = "Server List"
$label3.AutoSize = $true
$label3.Width = 25
$label3.Height = 10
$label3.location = new-object system.drawing.point(10, 10)
$label3.Font = "Microsoft Sans Serif,10,style=Bold"
$Form.controls.Add($label3)
 
$label3 = New-Object system.windows.Forms.Label
$label3.Text = "Output"
$label3.AutoSize = $true
$label3.Width = 25
$label3.Height = 10
$label3.location = new-object system.drawing.point(10, 290)
$label3.Font = "Microsoft Sans Serif,10,style=Bold"
$Form.controls.Add($label3)
 
# TextBox
$InputBox = New-Object system.windows.Forms.TextBox
$InputBox.Multiline = $true
$InputBox.BackColor = "#A7D4F7"
$InputBox.Width = 280
$InputBox.Height = 250
$InputBox.ScrollBars = "Vertical"
$InputBox.location = new-object system.drawing.point(10, 30)
$InputBox.Font = "Microsoft Sans Serif,10,style=Bold"
$Form.controls.Add($inputbox)
 
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Multiline = $true
$outputBox.BackColor = "#FDFEFE"
$outputBox.Width = 960
$outputBox.Height = 630
$outputBox.ReadOnly = $true
$outputBox.ScrollBars = "Both"
$outputBox.WordWrap = $false
$outputBox.location = new-object system.drawing.point(10, 310)
$outputBox.Font = "Calibri,11,style=Bold"
$Form.controls.Add($outputBox)
 
# Button
$Pingbutton = New-Object system.windows.Forms.Button
$Pingbutton.BackColor = "#5bd22c"
$Pingbutton.Text = "Ping"
$Pingbutton.Width = 120
$Pingbutton.Height = 22
$Pingbutton.location = new-object system.drawing.point(309, 30)
$Pingbutton.Font = "Microsoft Sans Serif,8,style=bold"
$Pingbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$Pingbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Pingbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$Pingbutton.Add_Click( { pingInfo }) 
$Form.controls.Add($Pingbutton)
 
$healthcheckbutton = New-Object system.windows.Forms.Button
$healthcheckbutton.BackColor = "#5bd22c"
$healthcheckbutton.Text = "Health Check"
$healthcheckbutton.Width = 120
$healthcheckbutton.Height = 22
$healthcheckbutton.location = new-object system.drawing.point(309, 60)
$healthcheckbutton.Font = "Microsoft Sans Serif,8,style=bold"
$healthcheckbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$healthcheckbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$healthcheckbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$healthcheckbutton.Add_Click( { healthcheckinfo }) 
$Form.controls.Add($healthcheckbutton)
 
$uptimebutton = New-Object system.windows.Forms.Button
$uptimebutton.BackColor = "#5bd22c"
$uptimebutton.Text = "Uptime"
$uptimebutton.Width = 120
$uptimebutton.Height = 22
$uptimebutton.location = new-object system.drawing.point(309, 90)
$uptimebutton.Font = "Microsoft Sans Serif,8,style=bold"
$uptimebutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$uptimebutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$uptimebutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$uptimebutton.Add_Click( { uptimeinfo }) 
$Form.controls.Add($uptimebutton)
 
$checksqlbutton = New-Object system.windows.Forms.Button
$checksqlbutton.BackColor = "#5bd22c"
$checksqlbutton.Text = "Check SQL Services"
$checksqlbutton.Width = 120
$checksqlbutton.Height = 22
$checksqlbutton.location = new-object system.drawing.point(309, 120)
$checksqlbutton.Font = "Microsoft Sans Serif,8,style=bold"
$checksqlbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$checksqlbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$checksqlbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$checksqlbutton.Add_Click( { checksqlinfo }) 
$Form.controls.Add($checksqlbutton)
 
$rdpbutton = New-Object system.windows.Forms.Button
$rdpbutton.BackColor = "#5bd22c"
$rdpbutton.Text = "RDP"
$rdpbutton.Width = 120
$rdpbutton.Height = 22
$rdpbutton.location = new-object system.drawing.point(309, 150)
$rdpbutton.Font = "Microsoft Sans Serif,8,style=bold"
$rdpbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$rdpbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$rdpbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$rdpbutton.Add_Click( { rdpinfo }) 
$Form.controls.Add($rdpbutton)
 
$loggedonuserbutton = New-Object system.windows.Forms.Button
$loggedonuserbutton.BackColor = "#5bd22c"
$loggedonuserbutton.Text = "Logged On User"
$loggedonuserbutton.Width = 120
$loggedonuserbutton.Height = 22
$loggedonuserbutton.location = new-object system.drawing.point(309, 180)
$loggedonuserbutton.Font = "Microsoft Sans Serif,8,style=bold"
$loggedonuserbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$loggedonuserbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$loggedonuserbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$loggedonuserbutton.Add_Click( { loggedonuserinfo }) 
$Form.controls.Add($loggedonuserbutton)

$listofservicesbutton = New-Object system.windows.Forms.Button
$listofservicesbutton.BackColor = "#5bd22c"
$listofservicesbutton.Text = "List of Services"
$listofservicesbutton.Width = 120
$listofservicesbutton.Height = 22
$listofservicesbutton.location = new-object system.drawing.point(309, 210)
$listofservicesbutton.Font = "Microsoft Sans Serif,8,style=bold"
$listofservicesbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$listofservicesbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$listofservicesbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$listofservicesbutton.Add_Click( { listofservicesinfo }) 
$Form.controls.Add($listofservicesbutton)

$explorerbutton = New-Object system.windows.Forms.Button
$explorerbutton.BackColor = "#5bd22c"
$explorerbutton.Text = "Open Explorer"
$explorerbutton.Width = 120
$explorerbutton.Height = 22
$explorerbutton.location = new-object system.drawing.point(450, 30)
$explorerbutton.Font = "Microsoft Sans Serif,8,style=bold"
$explorerbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$explorerbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$explorerbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$explorerbutton.Add_Click( { explorerInfo }) 
$Form.controls.Add($explorerbutton)
 
$healthcheckrgebutton = New-Object system.windows.Forms.Button
$healthcheckrgebutton.BackColor = "#5bd22c"
$healthcheckrgebutton.Text = "Health Check RGE"
$healthcheckrgebutton.Width = 120
$healthcheckrgebutton.Height = 22
$healthcheckrgebutton.location = new-object system.drawing.point(450, 60)
$healthcheckrgebutton.Font = "Microsoft Sans Serif,8,style=bold"
$healthcheckrgebutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$healthcheckrgebutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$healthcheckrgebutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$healthcheckrgebutton.Add_Click( { healthcheckrgeInfo }) 
$Form.controls.Add($healthcheckrgebutton)
 
$uptimergebutton = New-Object system.windows.Forms.Button
$uptimergebutton.BackColor = "#5bd22c"
$uptimergebutton.Text = "Uptime RGE"
$uptimergebutton.Width = 120
$uptimergebutton.Height = 22
$uptimergebutton.location = new-object system.drawing.point(450, 90)
$uptimergebutton.Font = "Microsoft Sans Serif,8,style=bold"
$uptimergebutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$uptimergebutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$uptimergebutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$uptimergebutton.Add_Click( { uptimergeInfo }) 
$Form.controls.Add($uptimergebutton)
 
$rdprgebutton = New-Object system.windows.Forms.Button
$rdprgebutton.BackColor = "#5bd22c"
$rdprgebutton.Text = "RDP RGE"
$rdprgebutton.Width = 120
$rdprgebutton.Height = 22
$rdprgebutton.location = new-object system.drawing.point(450, 150)
$rdprgebutton.Font = "Microsoft Sans Serif,8,style=bold"
$rdprgebutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$rdprgebutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$rdprgebutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$rdprgebutton.Add_Click( { rdprgeInfo }) 
$Form.controls.Add($rdprgebutton)
 
$logoffbutton = New-Object system.windows.Forms.Button
$logoffbutton.BackColor = "#5bd22c"
$logoffbutton.Text = "Logoff"
$logoffbutton.Width = 120
$logoffbutton.Height = 22
$logoffbutton.location = new-object system.drawing.point(800, 30)
$logoffbutton.Font = "Microsoft Sans Serif,8,style=bold"
$logoffbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$logoffbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$logoffbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$logoffbutton.Add_Click( { logoffinfo }) 
$Form.controls.Add($logoffbutton)

function pingInfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        if (Test-Connection $computer -Count 1 -Quiet) {
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#007F00"
            $outputBox.AppendText("$computer is Online`n")
        }
        else {
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#FF0000"
            $outputBox.AppendText("$computer is Offline`n")
        } 
    }
}
 
function uptimeinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        try {
            $RebootTime = Get-WmiObject win32_operatingsystem -ComputerName $computer -ErrorAction Stop | ForEach-Object { $_.ConvertToDateTime($_.LastBootUpTime) }
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#007F00"
            $outputBox.Appendtext("$computer - $RebootTime`n")
        }
        catch {
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#FF0000"
            $outputBox.Appendtext("$computer - Error`n")
        }
        finally {
            $Error.Clear()
        }
    }
}
 
function uptimergeinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        try {
            $RebootTime = Get-WmiObject win32_operatingsystem -ComputerName $computer -Credential $Cred -ErrorAction Stop | ForEach-Object { $_.ConvertToDateTime($_.LastBootUpTime) }
            $outputBox.Appendtext("$computer - $RebootTime`n")
        }
        catch {
            $outputBox.Appendtext("$computer - Error`n")
        }
        finally {
            $Error.Clear()
        }
    }
}
 
function healthcheckinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        $outputBox.AppendText($computer + ":-" + "`n")
        try {
            $Services = Get-WmiObject Win32_Service -ComputerName $computer -ErrorAction Stop | Where-Object { ($_.startmode -like "*auto*") -and ($_.state -notlike "*running*") }
            $Services | ForEach-Object { $_.StartService() } | Out-Null
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#007F00"
            $outputBox.AppendText("Health Check Done`n")
        }
        catch {
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#FF0000"
            $outputBox.Appendtext("Unable to perform Start Service`n")
        }
        Finally {
            $Error.Clear()
        }
        $outputBox.AppendText("`n")
    }
}
 
function healthcheckrgeinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        $outputBox.AppendText($computer + ":-" + "`n")
        try {
            $Services = Get-WmiObject Win32_Service -ComputerName $computer -Credential $Cred -ErrorAction Stop | Where-Object { ($_.startmode -like "*auto*") -and ($_.state -notlike "*running*") }
            $Services | ForEach-Object { $_.StartService() } | Out-Null
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#007F00"
            $outputBox.AppendText("Health Check Done`n")
        }
        catch {
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#FF0000"
            $outputBox.Appendtext("Unable to perform Start Service`n")
        }
        Finally {
            $Error.Clear()
        }
        $outputBox.AppendText("`n")
    }
}
 
Function checksqlinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    #Output Table Headers
    $C1Header = "ComputerName"
    $C2Header = "ServiceName"
    $C3Header = "StartMode"
    #Output Table Width
    $ColumnWidth = 25
    #Add Headers to OutputBox
    $outputBox.AppendText("$C1Header$(' '*($ColumnWidth-$C1Header.length))`t$C2Header$(' '*($ColumnWidth-$C2Header.Length))`t$C3Header$(' '*($ColumnWidth-$C3Header.Length))`n$('-'*($ColumnWidth*3))`n")
    foreach ($computer in $computers) {
        try {
            $outputBox.AppendText("`n")
            $sqlservices = Get-WmiObject Win32_Service -ComputerName $computer -ErrorAction Stop | Where-Object { ($_.Displayname -like "*sql*") }
            foreach ($sqls in $sqlservices) {
                #Add Values to the output
                $outputBox.AppendText("$computer$(' '*($ColumnWidth-$computer.length))`t$($sqls.Name)$(' '*$($ColumnWidth-$sqls.Name.length))`t$($sqls.StartMode)$(' '*$($ColumnWidth-$sqls.StartMode.length))`n")
            }
        }
        catch {
            $outputBox.AppendText("`n")
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#FF0000"
            $outputBox.AppendText("$computer - Error`n")
        }
        finally {
            $Error.Clear()
        }
    }
}

function rdpinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        $mstsc = Connect-Mstsc $computer $domain\$user $password -Admin -ErrorAction Stop
        $mstsc
    }
}
 
function logoffinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        $session = ((quser /server:$computer | Where-Object { $_ -match $User }) -split ' +')[2]
        logoff $session /server:$computer
       
    }
}
 
function loggedonuserinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    $outputBox.AppendText($c)
    foreach ($computer in $computers) {
        try {
            $outputBox.AppendText($computer + ":-" + "`n")
            $loggedonuser = Get-WmiObject -Class win32_computersystem -ComputerName $computer -ErrorAction Stop
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#007F00"
            $outputBox.AppendText($loggedonuser.username + "`n")
        }
 
        catch {
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#FF0000"
            $outputBox.AppendText("Error`n")
        }
 
        Finally {
            $Error.Clear()
        }
        $outputBox.AppendText("`n")
    }
 
}
 
function rdprgeinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        $mstsc = Connect-Mstsc $computer rge-extranet\suhail_asrulsani-ops Welcome1 -Admin -ErrorAction Stop
        $mstsc
    }
}
 
function explorerinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        explorer \\$computer\c$
    }
}

function listofservicesinfo {
    $outputBox.Clear()
    $computers = $InputBox.lines.Split("`n")
    foreach ($computer in $computers) {
        $outputBox.AppendText($computer + ":-" + "`n")
        try {
            $Services1 = Get-WmiObject Win32_Service -ComputerName $computer -ErrorAction Stop |
            Where-Object {
                ($_.startmode -like "*auto*") -and ($_.state -notlike "*running*")
            }
            foreach ($Service in $Services1) {
                $outputBox.Font = "Calibri,11,style=Bold"
                $outputBox.SelectionColor = "#007F00"
                $outputBox.AppendText($Service.DisplayName + "`n")
            }
        }

        catch {
            $outputBox.Font = "Calibri,11,style=Bold"
            $outputBox.SelectionColor = "#FF0000"
            $outputBox.Appendtext("Unable to list the services`n")
        }

        finally {
            $Error.Clear()
        }
        $outputBox.AppendText("`n")
    }
}

 
# show form
[void] $Form.ShowDialog()
