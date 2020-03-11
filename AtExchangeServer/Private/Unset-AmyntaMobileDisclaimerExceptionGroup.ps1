<# 

Feeder: AD\Get-DNsFromMail.ps1

#>
$Import = Get-Content 'C:\users\atrocano\Documents\working_Set\DistinguishedName.txt'
#$group = 'CN=AO Mobile Disclaimer Exemption,OU=Global Security Groups,DC=amtrustservices,DC=com'
$group = 'CN=Amynta Mobile Disclaimer Exemption,OU=Amynta Mobile Disclaimer Exemption,OU=SecurityGroups,OU=Amynta Migration,OU=Security_Groups,OU=IT,DC=amtrustservices,DC=com'
foreach ($alias in $Import) {
    Remove-ADGroupMember -Identity $group -Members $alias
}
Write-Host ('Done!')