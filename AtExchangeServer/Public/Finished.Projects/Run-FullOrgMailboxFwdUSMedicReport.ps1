$Wildcard = '*usmedic.com'
$objMailbox = @()
$Report = @()
$scriptBlock = {
    param (
        [string]$strWildcard
    ) 
    Get-Mailbox -ResultSize Unlimited -Filter "ForwardingSMTPAddress -like '$strWildcard'" | Select-Object UserPrincipalName,PrimarySmtpAddress,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward
} 
$objMailbox = Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $Wildcard | Select-Object UserPrincipalName,PrimarySmtpAddress,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward
$Report += $objMailbox | Select-Object UserPrincipalName,PrimarySmtpAddress,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward

Write-Host ('Done !')
if ($Report.count -lt 3) { $Report | Format-List } else { $Report }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\FullOrgMailboxFwdUSMedicReport.csv') -NoTypeInformation