<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            
    CSV: Alias, Start, End
    E.g., Sharan.Gopi@amtrustgroup.com,1/25/2019 8:10AM,1/28/2019 8:10AM 
#>

$Import = Import-CSV 'C:\Users\atrocano\Desktop\mailboxes.csv' 
foreach ($mailbox in $Import) {
    $scriptBlock = {
        param (
            [datetime]$dtStart,
            [datetime]$dtEnd,
            [string]$strAlias
        )
        <# Source=SMTP && EventId=SEND #>
        Get-TransportService | Get-MessageTrackingLog -ResultSize Unlimited -EventId SEND -Start $dtStart -End $dtEnd -Sender $strAlias | Select-Object Timestamp,Serverhostname,Clienthostname,Source,EventId,Sender,Recipients,MessageSubject,RecipientCount,RecipientStatus,ReturnPath,Directionality,TotalBytes
    }
    $Title = 'Exchange Transport Logs Sender - ' + $mailbox.Alias
    $filename = 'C:\Users\atrocano\Documents\working_Set\MailboxSenderReport_' + $mailbox.Alias + '_Handoff.csv'
    <# Source=SMTP && EventId=SEND #>
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $mailbox.Start, $mailbox.End, $mailbox.Alias | Where-Object {$_.Source -eq 'SMTP'} | Select-Object Timestamp,Serverhostname,Clienthostname,Source,EventId,Sender,Recipients,MessageSubject,RecipientCount,RecipientStatus,ReturnPath,Directionality,TotalBytes | Sort-Object -Property Timestamp | Export-Csv -Path $filename -NoTypeInformation
}
Write-Host ('Done !')