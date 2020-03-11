<#

#>
function GetHostname {
    (' (Computer: ' + ([System.Net.Dns]::GetHostEntry([string]"localhost").HostName) + ') ' )
}
GetHostname