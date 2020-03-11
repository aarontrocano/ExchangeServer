<#

one-liner to export all Distribution Lists in Exchange Server with a @ownerguard.com email
#>
get-distributiongroup -resultsize unlimited -Filter 'EmailAddresses -like "*ownerguard.com*"' | 
Sort-Object -Property PrimarySmtpAddress | 
Select-Object PrimarySmtpAddress,Alias,Name,DisplayName,GroupType,RecipientTypeDetails,WhenCreated,@{Name='DistributionGroupMembers';Expression={[string]::join(";", (Get-DistributionGroupMember-Identity $($_.alias) | Sort-Object -Property Name | Select-Object -ExpandProperty PrimarySmtpAddress))}} | 
Export-Csv -Path C:\users\aatrocano\desktop\amtogdl.csv -NoTypeInformation