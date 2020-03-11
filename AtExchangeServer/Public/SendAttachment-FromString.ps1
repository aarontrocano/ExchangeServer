<#

#>
$To = 'user01@example.com'
$From = 'user01@example.com'
$Subject = 'Forwarding Results for a Test Company'
$SMTPServer = 'mail01.example.com'
$Priority = 'High'
$logging = 'OnSuccess','OnFailure','Delay'
$Body = 'Body goes here'
$attText = "The text of the attachment"
$attName = "Test.txt"
Start-Sleep -Seconds 1


$Message = New-Object System.Net.Mail.MailMessage($From,$To,$Subject,$Body)
$Message.IsBodyHtml = $false
$Message.Priority = $Priority
$Message.DeliveryNotificationOptions = $logging
#$attachment = New-Object System.Net.Mail.Attachment($filenameAndPath)
#$Message.Attachments.Add($attachment)
<#$attachment = [System.Net.Mail.Attachment]::CreateAttachmentFromString($attText,$attName) #>
$attachment = [System.Net.Mail.Attachment]::CreateAttachmentFromString($attText, 'foobar.csv') 
$Message.Attachments.Add($attachment)


$SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, 587) 
$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($User, $Pword); 
$SMTPClient.Send($Message)
