<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            
    CSV: oldemail
    E.g., Sharan.Gopi@amtrustgroup.com
#>
$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\mailboxes.csv')
$Report = @()
$ReportShort = @()
$errorlog = @()
foreach ($alias in $Import) {
    function stripHTMLRegEx {
        param (
            [string]$strMessage
        )
        $strMessage = $strMessage -replace '<[^>]+>',''
        $strMessage = $strMessage -replace '&nbsp;',''
        $strMessage
    }
    $scriptBlock = {
        param (
            [string]$strAlias
        ) 
        Get-MailboxAutoReplyConfiguration -ResultSize Unlimited -Identity $strAlias | Select-Object AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage
    } 
    Write-Host ('Get-MailboxAutoReplyConfiguration for: ' + $alias.oldemail )
    $objState = Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias.oldemail | Select-Object AutoReplyState,EndTime,InternalMessage,ExternalMessage
    [string]$strInternalMessage = (stripHTMLRegEx ([string]$objState.InternalMessage)).Trim()
    [string]$strExternalMessage = (stripHTMLRegEx ([string]$objState.ExternalMessage)).Trim()
    $objReport = New-Object PSObject -Property @{
        Alias = $alias.oldemail
        AutoReplyState = $objState.AutoReplyState
        StartTime = $objState.StartTime
        EndTime = $objState.EndTime
        InternalMessage = $strInternalMessage
        ExternalMessage = $strExternalMessage
    }
    $Report = $objReport | Select-Object Alias,AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage
    $ReportShort = $objReport | Select-Object Alias,AutoReplyState,StartTime,EndTime
    if (! ($objState) ) { $errorlog += $alias.oldemail }
    $objState, $objReport, $strInternalMessage, $strExternalMessage = $null
}
Write-Host ('Done !')
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\OOOReportout.csv') -NoTypeInformation
$out = $ReportShort
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\OOOReportShortout.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\OOOReportout_errorlog.txt')