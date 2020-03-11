<# $UserCredential = (Get-Credential amtrustservices\atrocano)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://cle-mail04.amtrustservices.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -AllowClobber #>

Set-MailboxAutoReplyConfiguration -Identity "Aaron Trocano" -AutoReplyState Disabled -InternalMessage $null -ExternalMessage $null