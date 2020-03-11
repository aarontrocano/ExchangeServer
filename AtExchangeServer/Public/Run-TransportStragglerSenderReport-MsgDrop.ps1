<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            
    CSV: Alias,Start,End
    E.g., Sharan.Gopi@amtrustgroup.com,1/25/2019 8:10AM,1/28/2019 8:10AM 
#>

$Import = Import-CSV ([Environment]::GetFolderPath("Desktop")+'\mailboxes.csv')
foreach ($mailbox in $Import) {
    $scriptBlock = {
        param (
            [datetime]$dtStart,
            [datetime]$dtEnd,
            [string]$strAlias
        )
        <# Source=STOREDRIVER && EventId=RECEIVE #>
        Get-TransportService | Get-MessageTrackingLog -ResultSize Unlimited -EventId RECEIVE -Start $dtStart -End $dtEnd -Sender $strAlias | Select-Object Timestamp,Serverhostname,Clienthostname,Source,EventId,Sender,Recipients,MessageSubject,RecipientCount,RecipientStatus,ReturnPath,Directionality,TotalBytes
    }
    $filename = ([Environment]::GetFolderPath("Desktop")+'\MsgDrop\MailboxSenderReport_' + $mailbox.Alias + '_MsgDrop.csv')
    <# Source=STOREDRIVER && EventId=RECEIVE #>
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $mailbox.Start, $mailbox.End, $mailbox.Alias | Where-Object {$_.Source -eq 'STOREDRIVER'} | Select-Object Timestamp,Serverhostname,Clienthostname,Source,EventId,Sender,Recipients,MessageSubject,RecipientCount,RecipientStatus,ReturnPath,Directionality,TotalBytes | Sort-Object -Property Timestamp | Export-Csv -Path $filename -NoTypeInformation
}
Write-Host ('Done !')