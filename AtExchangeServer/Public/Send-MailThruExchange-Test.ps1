$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.csv')
$To = '<user@example.com>'
$From = '<user@example.com>'
$Subject = 'Forwarding Results for a Test Company'
$SMTPServer = 'mail01.example.com'
$Priority = 'High'
$logging = 'OnSuccess','OnFailure','Delay'
$Head = C:\Utilities\ServerScripts\ExchangeServer\GetHead-cssReportTable.ps1
$Body = $Import | ConvertTo-Html -Property DisplayName,SamAccountName,EmployeeID,PrimarySmtpAddress,RecipientType,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward,ServerName,Database,LitigationHoldEnabled,AmyntaMobileDisclaimerExemption,Office,Title,Department,Description,CanonicalName,AutoReplyState,EndTime,InternalMessage,ExternalMessage -head $Head | Out-String
Start-Sleep -Seconds 1


$Message = New-Object System.Net.Mail.MailMessage($From,$To,$Subject,$Body)
$Message.IsBodyHtml = $true
$Message.Priority = $Priority
$Message.DeliveryNotificationOptions = $logging
#$attachment = New-Object System.Net.Mail.Attachment($filenameAndPath)
#$Message.Attachments.Add($attachment)


$SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, 587) 
$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($User, $Pword); 
$SMTPClient.Send($Message)


