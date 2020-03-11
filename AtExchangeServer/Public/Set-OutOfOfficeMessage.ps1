<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            
    CSV: oldemail,Start,End,OOOMsg
    E.g., user@example.com,1/31/2019 3:00PM,1/31/2080 8:00PM,Thank you for your email. User no longer works at Example, Co. Please reach out to User02 for assistance (user02@example.com; 212.555.1234).
#>
$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.csv') 
$report = @()
$errorlog = @()
foreach ($alias in $Import) {
    $scriptBlock = {
        param (
            [string]$strAlias,
            [datetime]$dtStart,
            [datetime]$dtEnd,
            [string]$strOOOMsg
        )
        Set-MailboxAutoReplyConfiguration -Identity $strAlias -AutoReplyState Scheduled -StartTime $dtStart -EndTime $dtEnd -InternalMessage $strOOOMsg -ExternalMessage $strOOOMsg
    }
    $scriptBlock2 = {
        param (
            [string]$strAlias
        ) 
        Get-MailboxAutoReplyConfiguration -ResultSize Unlimited -Identity $strAlias | Select-Object AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage
    } 
    Write-Host ('Set-MailboxAutoReplyConfiguration for: ' + $alias.oldemail + ' | ' + $alias.Start + ' | ' + $alias.End)
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias.oldemail, $alias.Start, $alias.End, $alias.OOOMsg
    $objState = Invoke-Command -Session $Session -ScriptBlock $scriptBlock2 -ArgumentList $alias.oldemail | Select-Object AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage
    $Report += $objState | Select-Object @{n="Alias";e={$alias.oldemail} },AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage
    if (! ($objState) ) { $errorlog += $alias.oldemail }
    $objState = $null
}
Write-Host ('Done !')
$out = $report
$out | Export-Csv -Path  ([Environment]::GetFolderPath("Desktop")+'\OOOSetout.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\OOOSetout_errorlog.txt')