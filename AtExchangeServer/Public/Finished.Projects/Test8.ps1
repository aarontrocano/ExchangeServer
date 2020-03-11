$method1 = {$Import = Import-Csv -Path C:\users\atrocano\Documents\working_Set\SamAccountName.csv
    $smbox = Get-Recipient -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"}
    $ToMatch = @()
    $report = @()
    $test = $null
    foreach ($alias in $Import) { $ToMatch += ('AMTRUSTSERVICES\' + $alias.samAccountName) }
    foreach ($alias in $smbox) {
        $test = $alias | Get-MailboxPermission | Where-Object { $ToMatch -contains $_.user } <# $_.user is a sam id #> 
        #$test = $alias | Where-Object { $ToMatch -contains $_} 
        $Report += $test.Identity
        $test.Identity | Format-List
    } 
    Write-Host ('Done!')
    $Report
    $Out = $Report
    $Out | Out-String | Set-Content 'C:\users\atrocano\Documents\working_Set\MatchUserToSharedMailboxReport.txt'
}
Measure-Command $method1 | Select-Object totalseconds