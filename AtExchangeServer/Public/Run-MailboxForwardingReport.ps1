$Import = Get-Content -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.txt')
$Report = @()
$errorlog = @()
foreach ($alias in $Import) {
    $scriptBlock = {
        param (
            [string]$strAlias
        ) 
        Get-Mailbox -Identity $strAlias | Select-Object DisplayName,DeliverToMailboxandForward,ForwardingAddress,ForwardingSMTPAddress,PrimarySMTPAddress
        
    } 
    $objMailbox = Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias | Select-Object DisplayName,DeliverToMailboxandForward,ForwardingAddress,ForwardingSMTPAddress,PrimarySMTPAddress
    <#
    $Report += $objMailbox | Select-Object DisplayName,DeliverToMailboxandForward,ForwardingAddress,ForwardingSMTPAddress,PrimarySMTPAddress
    #>
    if ( ($objMailbox.ForwardingSMTPAddress -or $objMailbox.ForwardingAddress) ) {
        $Report += $objMailbox | Select-Object DisplayName,DeliverToMailboxandForward,ForwardingAddress,ForwardingSMTPAddress,PrimarySMTPAddress
    }
    if (! ($objMailbox) ) { $errorlog += $alias }
}
Write-Host ('Done !')
if ($Import.count -lt 3) { $Report | Format-List } else { $Report | Format-Table }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\MailboxFwdReport.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\MailboxFwdReport.txt')