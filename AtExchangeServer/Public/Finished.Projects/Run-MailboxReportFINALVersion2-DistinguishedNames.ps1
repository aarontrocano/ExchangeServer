<#
    README: This version uses Invoke-Command cmdlet. It passes the $Session variable
            to pass the Get-Mailbox cmdlet to the remote session on the Exchange
            Server. 

    Feeder: AD\Get-DNsFromMail.ps1
#>
$Import = Get-Content ([Environment]::GetFolderPath("Desktop")+'\DistinguishedName.txt')
$Report = @()
$errorlog = @()
#$group = 'CN=AO Mobile Disclaimer Exemption,OU=Global Security Groups,DC=amtrustservices,DC=com'
$group = 'CN=Amynta Mobile Disclaimer Exemption,OU=Amynta Mobile Disclaimer Exemption,OU=SecurityGroups,OU=Amynta Migration,OU=Security_Groups,OU=IT,DC=amtrustservices,DC=com'
$members = Get-ADGroupMember -Identity $group; if ($members) {}
Set-Variable -Name constStrCommand -Value 'if ($members.distinguishedName -contains $alias) {return "YES"} else {return "NO"}' -Option Constant -Scope Global -Force
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
        Get-Mailbox -Identity $strAlias | Select-Object PrimarySmtpAddress,RecipientTypeDetails,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward,ServerName,Database,LitigationHoldEnabled
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
        RecipientTypeDetails = ' '
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
        RecipientTypeDetails = $($objMailbox.RecipientTypeDetails).toString()
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
    $Report += $objUser | Select-Object DisplayName,SamAccountName,EmployeeID,PrimarySmtpAddress,RecipientTypeDetails,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward,ServerName,Database,LitigationHoldEnabled,AmyntaMobileDisclaimerExemption,Office,Title,Department,Description,CanonicalName,AutoReplyState,EndTime,InternalMessage,ExternalMessage
    if (! ($objMailbox) ) { $errorlog += $alias }
}
if ($Import.count -lt 3) {$Report | Format-List} else {$Report | Format-Table }
$out = $Report
$out | Export-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.csv') -NoTypeInformation
$errorlog | Out-String | Set-Content ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport_errorlog.txt')
Write-Host ('Done !')
Write-Host ('Output written to: ' + ([Environment]::GetFolderPath("Desktop")+'\UserListMailReport.csv') )