$Import = Get-Content -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.txt')
$Report = @()
$errorlog = @()
$objFwd = @()
foreach ($alias in $Import) {
    $objMailbox = Get-Mailbox -Identity $alias | Select-Object ServerName, Database
    Write-Host ('Get-Mailbox for: ' + $alias + '.')
    $Report += $objMailbox | Select-Object @{n="Alias";e={$alias} },ServerName, Database
    Write-Host ($objMailbox.ServerName + ' | ' + $objMailbox.ServerName )
    if (! ($objMailbox) ) { $errorlog += $alias }
}
Write-Host ('Done !')
if ($Import.count -lt 3) {$objFwd | Format-List} else {$objFwd }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\MailboxDatabaseServerout.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\MailboxDatabaseServer_errorlog.txt')