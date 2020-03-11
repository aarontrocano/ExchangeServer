<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            
    CSV: Alias,Start,End
    E.g., Sharan.Gopi@amtrustgroup.com,1/25/2019 8:10AM,1/28/2019 8:10AM 
#>

$Import = Import-CSV 'C:\Users\atrocano\Documents\working_Set\mailboxes.csv' 
foreach ($mailbox in $Import) {
    $scriptBlock = {
        param (
            [datetime]$dtStart,
            [datetime]$dtEnd,
            [string]$strAlias
        )
        Get-TransportService | Get-MessageTrackingLog -ResultSize Unlimited -Start $dtStart -End $dtEnd -Recipients $strAlias | Select-Object Timestamp,Serverhostname,Source,EventId,Sender,Recipients,MessageSubject
    }
    $Title = 'Exchange Transport Logs - ' + $mailbox.Alias
    $filename = ('C:\Users\atrocano\Documents\working_Set\MailboxReport_' + $mailbox.Alias + '.csv')
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $mailbox.Start, $mailbox.End, $mailbox.Alias | Select-Object Timestamp,Serverhostname,Source,EventId,Sender,Recipients,MessageSubject | Sort-Object -Property Timestamp | Export-Csv -Path $filename -NoTypeInformation
}
Write-Host ('Done !')