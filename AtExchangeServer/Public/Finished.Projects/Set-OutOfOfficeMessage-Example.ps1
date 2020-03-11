$UserCredential = (Get-Credential amtrustservices\atrocano)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://cle-mail04.amtrustservices.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -AllowClobber

Set-MailboxAutoReplyConfiguration -Identity "Aaron Trocano" -AutoReplyState Enabled -StartTime "08:00 12/12/2018" -EndTime "10:00 11/21/2080" -InternalMessage "Thank you for your email. Maria Mastrogiacomo no longer works at AmTrust. Please reach out to Erin Harker for assistance (Erin.Harker@amtrustgroup.com; 646.458.7965)." -ExternalMessage "Thank you for your email. Maria Matrogiacomo no longer works at AmTrust. Please reach out to Erin Harker for assistance (Erin.Harker@amtrustgroup.com; 646.458.7965)."