<#
    README: use CSV with Smtp addresses as the source.
            try Get-PrimarySMTPAddressVersion2.ps1

    CSV format: "Alias","PrimarySmtpAddress"
                "example.com/Locations/North America/User","User@example.com"
#>
$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\MailboxSmtpServerout.csv')
$Report = @()
$stats = $null
foreach ($alias in $Import) {
    $stats = Get-MailboxStatistics -Identity $alias.PrimarySmtpAddress | Select-Object DisplayName,MailboxType,TotalItemSize,TotalDeletedSize,MessageTableTotalSize,AttachmentTableTotalSize,ItemCount,TotalDeletedItemSize,LastLogonTime,Database,ServerName
    $Report += $stats | Select-Object DisplayName,MailboxType,@{n="PrimarySmtpAddress";e={$alias.PrimarySmtpAddress} },@{n="CanonicalName";e={$alias.alias} },TotalItemSize,TotalDeletedSize,MessageTableTotalSize,AttachmentTableTotalSize,ItemCount,TotalDeletedItemSize,LastLogonTime,Database,ServerName
}
Write-Host ('Done !')
if ($Import.count -lt 3) { $Report | Format-List } else { $Report }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\SharedMailboxSizeReportout.csv') -NoTypeInformation