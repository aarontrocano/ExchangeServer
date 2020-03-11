<# 
    https://medium.com/365uc/export-fullaccess-sendas-permissions-for-shared-mailboxes-be2a93d9d206
#>
$OutFile = ([Environment]::GetFolderPath("Desktop")+'\TempPermissionExport.txt')
"DisplayName" + "^" + "Alias" + "^" + "Full Access" + "^" + "Send As" | Out-File $OutFile -Force

$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Select-Object Identity, Alias, DisplayName, DistinguishedName
ForEach ($Mailbox in $Mailboxes)	{
	$SendAs = Get-ADPermission $Mailbox.DistinguishedName | Where-Object {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITYSELF" -and !$_.IsInherited} | ForEach-Object {$_.User}
	$FullAccess = Get-MailboxPermission $Mailbox.Identity | Where-Object {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | ForEach-Object {$_.User}
	$Mailbox.DisplayName + "^" + $Mailbox.Alias + "^" + $FullAccess + "^" + $SendAs | Out-File $OutFile -Append
}