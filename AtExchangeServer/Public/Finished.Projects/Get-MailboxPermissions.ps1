<# 
    c:\Utilities\ServerScripts\Exchange\Get-MailboxPermissions.ps1    
    script from George Secara 
    input file is list of Microsoft Exchange Server mailbox aliases
    No trailing spaces, no empty lines
#>
$Import = Get-Content 'C:\Users\atrocano\Documents\working_Set\mailboxaliases.txt'
$Total = $Import.count
$counter = 1
Write-Host (' Import: ' + $Import + ' Total: ' + $Total + ' Counter: ' + $Counter )
foreach ($alias in $Import)
{
    Write-Host (' foreach: alias: ' + $alias )   
    Get-MailboxPermission $alias | Where-object { ($_.IsInherited -eq $False) -and ($_.User -notlike "NT AUTHORITY\SELF") -and ($_.user.SecurityIdentifier -ne "S-1-5-10") -and($_.user -notlike "s-1-5*" )} | Select-Object Identity, @{name=" Users that have Full Access ";expression={(Get-User $_.User).Name}},@{name=" Samaccountname ";expression={(Get-User $_.User).samaccountname}}, AccessRights  | export-csv 'C:\Users\atrocano\Documents\working_Set\permissions.csv' -Append
    Write-Progress -Activity "Working... $counter of $Total complete" 
    $counter++

}  
