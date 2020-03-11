$Mailboxes = Get-Mailbox -ResultSize Unlimited
$Total = $Mailboxes.count
$count = 1
$report = @()
foreach ($mailboxalias in $Mailboxes) {
    
    $mailbox = Get-Mailbox -identity $mailboxalias.alias -ResultSize Unlimited 
    $report += ( $mailbox | c:\users\atrocano\Utilities\ServerScripts\Exchange\Get-MailboxDelegations.ps1 )
    if (($count % 500) -eq (0)) {Write-Host ('Working on ' + $count + ' of ' + $Total + ' mailboxes.')} elseif (($count) -lt (4)) {Write-Host ('Working on ' + $count + ' of ' + $Total + ' mailboxes.')}
    $count++
}
$out = $report
$out | Export-Csv 'C:\Users\atrocano\Documents\working_Set\ForEachFullOrgMailboxresults.csv'