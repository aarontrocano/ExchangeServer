$Import = Get-Content -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.txt')
$report = @()
$errorlog = @()
$objMapi = @()
foreach ($alias in $Import) {
    Write-Host ('Get-CASMailbox for: ' + $alias + '.')
    $report += Get-CASMailbox $alias
    $objMailbox = Get-CASMailbox $alias
    $obJMapi += $objMailbox | Select-Object @{n="Alias";e={$alias} },MapiEnabled
    Write-Host $objMailbox.MapiEnabled
    if (! ($objMailbox) ) { $errorlog += $alias }
}
Write-Host ('Done !')
<# $report | Out-String #>
$objMapi | Out-String
$out = $objMapi
$out | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\casout_1-14.txt')
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\cas_errorlog_1-14.txt')