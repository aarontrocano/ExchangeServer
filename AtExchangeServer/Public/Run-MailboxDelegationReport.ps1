$Mailboxes = Import-CSV 'C:\Users\atrocano\Documents\working_Set\mailboxes.csv' 
$Total = $Mailboxes.count
$count = 1
$report = @()
foreach ($alias in $Mailboxes) {
    
    $mailbox = Get-Mailbox -identity $alias.Mailbox -ResultSize Unlimited 
    $report += ( $mailbox | c:\users\atrocano\Utilities\ServerScripts\Exchange\Get-MailboxDelegations.ps1 )
    if (($count % 3) -eq (0)) {Write-Host ('Working on ' + $count + ' of ' + $Total + ' mailboxes.')} elseif (($count) -lt (4)) {Write-Host ('Working on ' + $count + ' of ' + $Total + ' mailboxes.')}
    $count++
}
$out = $report
$out | Export-Csv 'C:\Users\atrocano\Documents\working_Set\mailboxresults.csv'