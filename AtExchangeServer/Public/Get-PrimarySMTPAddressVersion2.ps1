<#
    README: This version uses Invoke-Command cmdlet. It passes the $Session variable
            to pass the Get-Mailbox cmdlet to the remote session on the Exchange
            Server. 
#>
$Import = Get-Content -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.txt')
$Report = @()
$errorlog = @()
foreach ($alias in $Import) {
    $scriptBlock = {
        param (
            [string]$strAlias
        )
        Get-Mailbox -Identity $strAlias | Select-Object PrimarySmtpAddress
    }
    $objMailbox = Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias | Select-Object PrimarySmtpAddress
    Write-Host ('Get-Mailbox for: ' + $alias + '. ' + $objMailbox.ServerName + ' | ' + $objMailbox.PrimarySmtpAddress )
    $Report += $objMailbox | Select-Object @{n="Alias";e={$alias} },PrimarySmtpAddress
    if (! ($objMailbox) ) { $errorlog += $alias }
}
Write-Host ('Done !')
if ($Import.count -lt 3) {$Report | Format-List} else {$Report }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\MailboxSmtpServerout.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\MailboxSmtpServer_errorlog.txt')