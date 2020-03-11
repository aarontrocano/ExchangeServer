<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            
    CSV: oldemail
    E.g., user@example.com
#>
$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.csv')
foreach ($alias in $Import) {
    $scriptBlock = {
        param (
            [string]$strAlias
        ) 
        Set-Mailbox -Identity $strAlias -DeliverToMailboxAndForward $false -ForwardingAddress $null -ForwardingSMTPAddress $null
    } 
    Write-Host ($alias.oldemail + ' undo forwarding. ')
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias.oldemail
}
Write-Host ('Done !')