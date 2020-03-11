<# 
    Feeder script for this: Run-MailboxReportFINALVersion2-DistinguishedNames.ps1
    rename: UserListMailReportout.csv to FullOrgMailboxFwdAmyntaMobileReport.csv
#>
$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\FullOrgMailboxFwdAmyntaAmyntaMobileReport.csv')
[long]$Total = $Import.count
[long]$count = 0
ForEach ($alias in $Import) {if ($alias.AmyntaMobileDisclaimerException -eq "YES") {$count++}}
$count
$Total