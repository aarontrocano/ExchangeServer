<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            
    CSV: oldemail
    E.g., user@example.com
#>
$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.csv')
$report = @()
$errorlog = @()
foreach ($alias in $Import) {
    $scriptBlock = {
        param (
            [string]$strAlias
        )
        Set-MailboxAutoReplyConfiguration -Identity $strAlias -AutoReplyState Disabled -InternalMessage $null -ExternalMessage $null
    } 
    $scriptBlock2 = {
        param (
            [string]$strAlias
        ) 
        Get-MailboxAutoReplyConfiguration -ResultSize Unlimited -Identity $strAlias | Select-Object AutoReplyState
    }
    Write-Host ('UnSet-MailboxAutoReplyConfiguration for: ' + $alias.oldemail)
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias.oldemail 
    $objState = Invoke-Command -Session $Session -ScriptBlock $scriptBlock2 -ArgumentList $alias.oldemail | Select-Object AutoReplyState
    $Report += $objState | Select-Object @{n="Alias";e={$alias.oldemail} },AutoReplyState
    if (! ($objState) ) { $errorlog += $alias.oldemail }
    $objState = $null
}
Write-Host ('Done !')
$out = $report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\OOOUnsetout.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\OOOUnsetout_errorlog.txt')