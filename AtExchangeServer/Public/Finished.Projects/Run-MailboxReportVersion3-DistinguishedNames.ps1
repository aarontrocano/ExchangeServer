<#
    README: This version uses Invoke-Command cmdlet. It passes the $Session variable
            to pass the Get-Mailbox cmdlet to the remote session on the Exchange
            Server. 

    Feeder: AD\Get-DNsFromMail.ps1
#>
$Import = Get-Content ([Environment]::GetFolderPath("Desktop")+'\DistinguishedName.txt')
$Report = @()
$errorlog = @()
Write-Host ('Collecting Amynta Mobile Exemption Group. | ' + (Get-Date).toString() )
#$group = 'CN=AO Mobile Disclaimer Exemption,OU=Global Security Groups,DC=amtrustservices,DC=com'
$group = 'CN=Amynta Mobile Disclaimer Exemption,OU=Amynta Mobile Disclaimer Exemption,OU=SecurityGroups,OU=Amynta Migration,OU=Security_Groups,OU=IT,DC=amtrustservices,DC=com'
$members = Get-ADGroupMember -Identity $group; if ($members) {}
Set-Variable -Name constStrCommand -Value 'if ($members.distinguishedName -contains $alias) {return "YES"} else {return "NO"}' -Option Constant -Scope Global -Force
Write-Host ('Collecting report objects. | ' + (Get-Date).toString() )
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
        Get-Mailbox -Identity $strAlias | Select-Object PrimarySmtpAddress,RecipientType,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward,ServerName,Database,LitigationHoldEnabled
    }
    $scriptBlock2 = {
        param (
            [string]$strAlias
        ) 
        Get-MailboxAutoReplyConfiguration -ResultSize Unlimited -Identity $strAlias | Select-Object AutoReplyState,EndTime,InternalMessage,ExternalMessage
    }
    $searcher  = [adsisearcher]"(distinguishedname=$alias)"
    $objMailbox = New-Object PSObject -Property @{
        PrimarySmtpAddress = ' '
        RecipientType = ' '
        ForwardingSMTPAddress = ' '
        ForwardingAddress = ' '
        DeliverToMailboxandForward = ' '
        ServerName = ' '
        Database = ' '
        LitigationHoldEnabled = ' '
    }
    $objMailbox = Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias
    $objState = New-Object PSObject -Property @{
        AutoReplyState = ' '
        EndTime = ' '
        InternalMessage = ' '
        ExternalMessage = ' '
    }
    $objState = Invoke-Command -Session $Session -ScriptBlock $scriptBlock2 -ArgumentList $alias | Select-Object AutoReplyState,EndTime,InternalMessage,ExternalMessage
    [string]$strInternalMessage = (stripHTMLRegEx ([string]$objState.InternalMessage)).Trim()
    [string]$strExternalMessage = (stripHTMLRegEx ([string]$objState.ExternalMessage)).Trim()
    $objUser = New-Object PSObject -Property @{
        DisplayName = $($searcher.FindOne().Properties.displayname).toString()
        SamAccountName = $($searcher.FindOne().Properties.samaccountname).toString()
        EmployeeID = $([string]$searcher.FindOne().Properties.employeeid).toString()
        PrimarySmtpAddress = $($objMailbox.PrimarySmtpAddress).toString()
        RecipientType = $($objMailbox.RecipientType).toString()
        ForwardingSMTPAddress = $([string]$objMailbox.ForwardingSMTPAddress).toString() #cast string coalesces a possible null value, for .toString() method
        ForwardingAddress = $([string]$objMailbox.ForwardingAddress).toString() #cast string coalesces a possible null value, for .toString() method
        DeliverToMailboxandForward = $($objMailbox.DeliverToMailboxandForward).toString()
        ServerName = $($objMailbox.ServerName).toString()
        Database = $($objMailbox.Database).toString()
        LitigationHoldEnabled = $($objMailbox.LitigationHoldEnabled).toString()
        AmyntaMobileDisclaimerExemption = Invoke-Expression -Command $constStrCommand
        Office = $([string]$searcher.FindOne().Properties.physicaldeliveryofficename).toString()
        Title = $([string]$searcher.FindOne().Properties.title).toString()
        Department = $([string]$searcher.FindOne().Properties.department).toString()
        Description = $([string]$searcher.FindOne().Properties.description).toString()
        CanonicalName = $(Get-AdUser -Identity $alias -Properties CanonicalName).CanonicalName.toString()
        AutoReplyState = $([string]$objState.AutoReplyState).toString()
        EndTime = $([string]$objState.EndTime).toString()
        InternalMessage = $strInternalMessage
        ExternalMessage = $strExternalMessage
    }
    $Report += $objUser | Select-Object DisplayName,SamAccountName,EmployeeID,PrimarySmtpAddress,RecipientType,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward,ServerName,Database,LitigationHoldEnabled,AmyntaMobileDisclaimerExemption,Office,Title,Department,Description,CanonicalName,AutoReplyState,EndTime,InternalMessage,ExternalMessage
    if (! ($objMailbox) ) { $errorlog += $alias }
}
if ($Import.count -lt 3) {$Report | Format-List} else {$Report | Format-Table }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport_errorlog.txt')
Write-Host ('Done ! | ' + (Get-Date).toString() )
Write-Host ('Output written to: ' + ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.csv') )


Write-Host ('Building Html Email message. | ' + (Get-Date).toString() )
$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.csv')
$To = 'Aaron Trocano <aaron.trocano@amtrustgroup.com>'
$From = 'Aaron Trocano <aaron.trocano@amtrustgroup.com>'
$Subject = 'Forwarding Results for a Test Company'
$SMTPServer = 'cle-mail04.amtrustservices.com'
$Priority = 'High'
$logging = 'OnSuccess','OnFailure','Delay'
$Head = C:\Utilities\ServerScripts\Exchange\GetHead-cssReportTable.ps1
$Foot = ('(Script: ' + $MyInvocation.MyCommand.Definition + ') ') <#GetFoot-ScriptName#>
$Foot += C:\Utilities\ServerScripts\Exchange\GetFoot-Hostname.ps1
$Foot += C:\Utilities\ServerScripts\Exchange\GetFoot-Signature.ps1
$PostContent = $Foot | Out-String
$Body = $Import | ConvertTo-Html -Property DisplayName,SamAccountName,EmployeeID,PrimarySmtpAddress,RecipientType,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward,ServerName,Database,LitigationHoldEnabled,AmyntaMobileDisclaimerExemption,Office,Title,Department,Description,CanonicalName,AutoReplyState,EndTime,InternalMessage,ExternalMessage -head $Head -PostContent $PostContent | Out-String
Start-Sleep -Seconds 1


Write-Host ('Building Excel Workbook attachment. | ' + (Get-Date).toString() )
<#Define locations and delimiter#>
$csv = ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.csv') 
$xlsx = ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.xlsx') 
$delimiter = "," 
<#(dotFile)Import PowerShell file as a Library , then call function#>
. C:\Utilities\ServerScripts\MSOffice\ConvertCsv-Excel.ps1
funcConvertCsv-Excel -csv $csv -xlsx $xlsx
Write-Host ('Sleeping | ' + (Get-Date).toString() )
Start-Sleep -Seconds 15


Write-Host ('Smtp sending. | ' + (Get-Date).toString() )
$Message = New-Object System.Net.Mail.MailMessage($From,$To,$Subject,$Body)
$Message.IsBodyHtml = $true
$Message.Priority = $Priority
$Message.DeliveryNotificationOptions = $logging
$attachment = New-Object System.Net.Mail.Attachment( ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.xlsx') )
$Message.Attachments.Add($attachment)


$SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, 587) 
$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($User, $Pword); 
$SMTPClient.Send($Message)
$SMTPClient.Dispose()
$Message.Dispose()
Write-Host ('Done! | ' + (Get-Date).toString() )