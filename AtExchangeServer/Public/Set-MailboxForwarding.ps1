<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            
    CSV: oldemail,newemail
    E.g., user01@example.com,user01@example.net
#>
$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.csv')
foreach ($alias in $Import) {
    $scriptBlock = {
        param (
            [string]$strAlias,
            [string]$strFwd
        ) 
        Set-Mailbox -Identity $strAlias -DeliverToMailboxAndForward $false -ForwardingAddress $null -ForwardingSMTPAddress $strFwd 
    } 
    Write-Host ($alias.oldemail + ' to ' + $alias.newemail)
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias.oldemail, $alias.newemail    
}
Write-Host ('Done !')