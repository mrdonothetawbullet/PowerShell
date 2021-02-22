function Get-LocalGroupMembers {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$Identity,
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Add-Type -AssemblyName System.DirectoryServices.AccountManagement 
    $context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine', $ComputerName)

    try {
        if (!([string]::IsNullOrEmpty($Identity))) {
            # search a specific group
            [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($context, $Identity)
        }
        else {
            # search all local groups
            $groupPrincipal    = New-Object DirectoryServices.AccountManagement.GroupPrincipal($context)
            $principalSearcher = New-Object DirectoryServices.AccountManagement.PrincipalSearcher($groupPrincipal)
        }
    }
    catch {
        throw "Error searching group(s) on '$ComputerName'. $($_.Exception.Message)"
    }
    finally {
        if ($groupPrincipal)    {$groupPrincipal.Dispose()}
        if ($principalSearcher) {$principalSearcher.FindAll()}
    }
}

(Get-LocalGroupMembers -Identity "Remote Desktop Users").Members | Select-Object -ExpandProperty Name
