<#
    README: This version doesn't use Invoke-Command cmdlet. It relies on PowerShell 
            profiles to setup remote session for Get-Mailbox cmdlet. 
#>
$Import = Get-Content -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.txt')
$report = @()
$errorlog = @()
$objSmtp = @()
foreach ($alias in $Import) {
    $objMailbox = Get-Mailbox -Identity $alias | Select-Object PrimarySmtpAddress
    Write-Host ('Get-Mailbox for: ' + $alias + '.')
    $report += $objMailbox | Select-Object @{n="Alias";e={$alias} },PrimarySmtpAddress
    $objSmtp += $objMailbox | Select-Object @{n="Alias";e={$alias} },PrimarySmtpAddress
    Write-Host ($objMailbox.ServerName + ' | ' + $objMailbox.PrimarySmtpAddress )
    if (! ($objMailbox) ) { $errorlog += $alias }
}
Write-Host ('Done !')
if ($Import.count -lt 3) {$objSmtp | Format-List} else {$objSmtp }
$out = $objSmtp
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\MailboxSmtpServerout.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\MailboxSmtpServer_errorlog.txt')