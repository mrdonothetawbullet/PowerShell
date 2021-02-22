([ADSI]"WinNT://./Administrators").psbase.Invoke('Members') | % {
 ([ADSI]$_).InvokeGet('AdsPath')
}
