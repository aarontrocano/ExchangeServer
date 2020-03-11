<# 

Feeder: AD\Get-DNsFromMail.ps1

#>
$Import = Get-Content ([Environment]::GetFolderPath("Desktop")+'\DistinguishedName.txt')
$Report = @()
$group = 'CN=Amynta Mobile Disclaimer Exemption,OU=Amynta Mobile Disclaimer Exemption,OU=SecurityGroups,OU=Amynta Migration,OU=Security_Groups,OU=IT,DC=amtrustservices,DC=com'
$members = Get-ADGroupMember -Identity $group
foreach ($alias in $Import) {
    $objUser = New-Object PSObject -Property @{
        DistinguishedName = $alias
        CanonicalName = $(Get-AdUser -Identity $alias -Properties CanonicalName).CanonicalName
        AmyntaMobileDisclaimerException = ' '
    }
    if ($members.distinguishedName -contains $alias) {$objUser.AmyntaMobileDisclaimerException='YES'} else {$objUser.AmyntaMobileDisclaimerException='NO'}
    $Report += $objUser | Select-Object DistinguishedName,CanonicalName,AmyntaMobileDisclaimerException
}
Write-Host ('Done!')
if ($Import.count -lt 3) { $Report | Format-List } else { $Report | Format-Table }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\AmyntaMobileout.csv') -NoTypeInformation