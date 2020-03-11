<#
    Requested by Guadalupe Lopez, 3/21/2019
#>
$list = 'Gap.Benefits@amtrustgroup.com','Ownerguard.claims@amtrustgroup.com','gapclaims@warrantysolutions.com'
$Report = @()
foreach ($alias in $list) {$Report += get-mailbox -Identity $alias | Select-Object PrimarySmtpAddress,RecipientTypeDetails }
$Report2 = @()
foreach ($alias in $list) {$REport2 += get-mailbox -Identity $alias | get-mailboxpermission | where-object { ($_.AccessRights -eq 'FullAccess' ) -and ( $_.IsInherited -eq $false ) -and -not ( $_.User -like 'NT AUTORITY\SELF' ) } }
$names = @()
foreach ($alias in $Report2.User) { $names += ($alias -split '\',-1,'SimpleMatch')[1] }
$names | Get-ADUser | Select-Object Name,SamAccountName | Sort-Object -Property Name
$Report | Format-Table -AutoSize