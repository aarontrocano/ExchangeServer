<# 
    README: This script uses Export-Csv, compatible with vscode, PowerShell 
            console, and PowerShell ISE. 
            This script is scheduled weekly and captures data from mon to fri.
            
    CSV: Alias
    E.g., user@example.com
#>
$Import = Import-CSV ([Environment]::GetFolderPath("Desktop")+'\stragglersmailboxes.csv')
$arrTo = 'user01@example.com','user02@example.com'
$From = 'user01@example.com'
$Subject = 'Weekly Straggler Senders Report - MAPI at Example.com'
$Body = $('All,','','Please see attached ZIP File Weekly Report.','','STOREDRIVE (plus) RECEIVE suggests a MAPI event.  The conclusion can get fuzzy from there.  One scenario that triggers this is a Outlook client or a WebMail client.  There are doubtless other ways that could also trigger it as could any other automated process.','','--','Automated system report')
$Priority = 'High'
$logging = 'OnSuccess','OnFailure','Delay'
$SMTPServer = 'mail01.example.com'
$port = 587
$working_set = ([Environment]::GetFolderPath("Desktop")+'\MsgDrop')
[long]$Epsilon = 256
$End = (Get-Date)
$Start = (Get-Date).AddDays(-4.5)
$FilesInReport = @()
$attachmentfilenameAndPath = ($working_set+'\TransportLogs.Zip')
$cleanupLogfilesScriptBlock = {
    Get-ChildItem ($working_set+'\*') -include *.csv -Recurse | Remove-Item
    Get-ChildItem ($working_set+'\*') -include *.zip -Recurse | Remove-Item
}
$zipLogfilesScriptBlock = {
    Get-ChildItem ($working_set+'\*') -include *.csv -Recurse | Where-Object { $_.Length -lt $Epsilon } | Remove-Item <#Prune (near) 0-byte CSVs#>
    Start-Sleep -Seconds 5
    $FilesInReport = Get-ChildItem ($working_set+'\*')
    Compress-Archive -Path $FilesInReport -CompressionLevel Optimal -DestinationPath $attachmentfilenameAndPath
}
$sendEmailScriptBlock = {
    $Message = New-Object System.Net.Mail.MailMessage
    $Message.From = $From
    ForEach ($alias in $arrTo) { $Message.To.Add($alias) }
    $Message.Subject = $Subject
    $Message.Body = $Body
    $Message.IsBodyHtml = $false
    $Message.Priority = $Priority
    $Message.DeliveryNotificationOptions = $logging
    $attachment = New-Object System.Net.Mail.Attachment($attachmentfilenameAndPath)
    $Message.Attachments.Add($attachment)
    
    $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, $port) 
    $SMTPClient.EnableSsl = $true 
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($User, $Pword)
    $SMTPClient.Send($Message)
}
Invoke-Command -ScriptBlock $cleanupLogfilesScriptBlock
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
    $filename = ($working_set + '\MailboxSenderReport_' + $mailbox.Alias + '_MsgDrop.csv')
    <# Source=STOREDRIVER && EventId=RECEIVE #>
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $Start, $End, $mailbox.Alias | Where-Object {$_.Source -eq 'STOREDRIVER'} | Select-Object Timestamp,Serverhostname,Clienthostname,Source,EventId,Sender,Recipients,MessageSubject,RecipientCount,RecipientStatus,ReturnPath,Directionality,TotalBytes | Sort-Object -Property Timestamp | Export-Csv -Path $filename -NoTypeInformation
}
Invoke-Command -ScriptBlock $zipLogfilesScriptBlock
Invoke-Command -ScriptBlock $sendEmailScriptBlock
Write-Host ('Done !')