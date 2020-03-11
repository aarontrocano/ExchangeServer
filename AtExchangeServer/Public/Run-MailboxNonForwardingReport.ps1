$Import = Get-Content -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.txt')
$Report = @()
$errorlog = @()
foreach ($alias in $Import) {
    $scriptBlock = {
        param (
            [string]$strAlias
        ) 
        Get-Mailbox -Identity $strAlias | Select-Object ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward
    } 
    $objMailbox = Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias | Select-Object ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward
    if ( ! ($objMailbox.ForwardingSMTPAddress -or $objMailbox.ForwardingAddress) ) {
        Write-Host ('Get-Mailbox for: ' + $alias + '.')
        $Report += $objMailbox | Select-Object @{n="Alias";e={$alias} },ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward
        Write-Host ('Forwarding: Null')
    }
    if (! ($objMailbox) ) { $errorlog += $alias }
}
Write-Host ('Done !')
if ($Import.count -lt 3) { $Report | Format-List } else { $Report }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\MailboxNonfwdout.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\MailboxNonfwd_errorlog.txt')