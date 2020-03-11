$Import = Get-Content -Path C:\Users\atrocano\Documents\working_Set\mailboxes.txt
$report = @()
$errorlog = @()
$objMapi = @()
foreach ($alias in $Import) {
    Write-Host ('Get-CASMailbox for: ' + $alias + '.')
    $report += Get-CASMailbox $alias
    $objMailbox = Get-CASMailbox $alias
    $obJMapi += $objMailbox | Select-Object @{n="Alias";e={$alias} },MapiEnabled,ActiveSyncEnabled,OWAEnabled
    Write-Host $objMailbox.MapiEnabled
    if (! ($objMailbox) ) { $errorlog += $alias }
}
Write-Host ('Done !')
<# $report | Out-String #>
$objMapi | Out-String
$out = $objMapi
$out | Export-Csv -Path 'C:\Users\atrocano\Documents\working_Set\casout.csv' -NoTypeInformation
$errorlog | Out-String | Set-Content C:\Users\atrocano\Documents\working_Set\cas_errorlog.txt