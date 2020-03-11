<# 
  .Synopsis 
  Collect common mailbox delegations - FullAccess, Send-As, and Send on Behalf. 
 
  .Description 
  Collects mailbox delegations for the purpose of planning migration batches to Exchange Online. Only tested on  
  Exchange 2010 SP3 (UR 16 & 17) and 2013 (CU 15 & 16). See Notes section for specific Exchange versions' usage  
  requirements. Delegate SamAccountName is included in case a group delegate has no DisplayName or  
  PrimarySmtpAddress (i.e. non-mail-enabled group). 
 
  .Parameter Mailbox 
  Accepts only proper mailbox objects ([Microsoft.Exchange.Data.Directory.Management.Mailbox]). These can come  
  from the pipeline, or can be manually specified. 
 
  .Parameter AdditionalFilteredUsers 
  Accepts one or more string values for known accounts that should be omitted from the output. See examples. 
 
  .Example 
  Get-Mailbox -ResultSize Unlimited | .\Get-MailboxDelegations.ps1 | Export-Csv C:\MailboxDelegations.csv 
  Import-CSV "c:\users\radcox\desktop\mailboxes.csv" |foreach-object {Get-Mailbox -identity $_.Mailbox -ResultSize Unlimited | .\Get-MailboxDelegations.ps1} | Export-Csv "c:\users\radcox\desktop\mailboxresults.csv"
 
  .Notes 
  1. For Exchange 2010, a regular PowerShell console must be used and have the Exchange snapin added. 
   
  2. For Exchange 2013 and 2016, either the Exchange Management Shell, or a regular PowerShell console with the  
  Exchange snapin added, can be used. 
   
  3. In any version of Exchange, simply being remoted to an Exchange server from a regular PowerShell console  
  will not work because the object type changes from [Microsoft.Exchange.Data.Directory.Management.Mailbox] to  
  [Deserialized.Microsoft.Exchange.Data.Directory.Management.Mailbox] and the Mailbox parameter only accepts the  
  former. 
   
  4. The approach used for filtering out users is borrowed from Zachary Loeber's script: 
  https://raw.githubusercontent.com/zloeber/Powershell/master/Exchange/Get-MailboxFullAccessPermission.ps1 
#> 

[CmdletBinding()] 
Param( 
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)] 
    [Microsoft.Exchange.Data.Directory.Management.Mailbox]$Mailbox, 
 
    [Parameter(Mandatory = $false)] 
    [string[]]$AdditionalFilteredUsers = @() 
)

Begin { 
 
    $FilteredUsers = @('S-1-*', 
        "*\Administrator", 
        "*\Discovery Management", 
        "*\Organization Management", 
        "*\Domain Admins", 
        "*\Enterprise Admins", 
        "*\Exchange Services", 
        "*\Exchange Trusted Subsystem", 
        "*\Exchange Servers", 
        "*\Exchange View-Only Administrators", 
        "*\Exchange Admins", 
        "*\Managed Availability Servers", 
        "*\Public Folder Administrators", 
        "*\Exchange Domain Servers", 
        "*\Exchange Organization Administrators", 
        "NT AUTHORITY\*") 
    $FilteredUsers += $AdditionalFilteredUsers 
    $FilteredUsersString = @($FilteredUsers | ForEach-Object {[regex]::Escape($_)}) 
    $FilteredUsersString = '^(' + ($FilteredUsersString -join '|') + ')$' 
    $FilteredUsersString = $FilteredUsersString -replace '\\\*', '.*' 
} 
Process { 
 
    if (-not ($Mailbox.RecipientTypeDetails -eq 'LegacyMailbox')) { 
 
        $AllDelegations = @() 
        $FullAccess = Get-MailboxPermission -Identity $Mailbox | Where-Object {($_.AccessRights -like '*FullAccess') -and ($_.isinherited -eq $False) -and ($_.User -notlike 'Blackberry 5 Admin*') -and -not ($_.User -match $FilteredUsersString) -and ($_.Deny -eq $false)} 
        $SendAs = Get-ADPermission -Identity $($Mailbox.DistinguishedName) | Where-Object {($_.ExtendedRights -like '*Send-As*') -and ($_.isinherited -eq $False) -and -not ($_.User -match $FilteredUsersString) -and ($_.Deny -eq $false)} 
        $SendOnBehalf = ($Mailbox | Where-Object {-not ($_.GrantSendOnBehalfTo -match $FilteredUsersString) -and -not ($_.GrantSendOnBehalfTo -eq $null)}).GrantSendOnBehalfTo 
    } 
    if ($FullAccess) { 
 
        foreach ($fa in $FullAccess) { 
 
            try { 
                $faUser = Get-User -Identity $fa.User.RawIdentity -ErrorAction SilentlyContinue 
                $faGroup = Get-Group -Identity $fa.User.RawIdentity -ErrorAction SilentlyContinue 
                if ($faGroup) {$faUser = $faGroup} 
            } 
            catch {} 
            $faDelegation = New-Object -TypeName PSObject -Property @{ 
                'MailboxDisplayName'     = $Mailbox.DisplayName; 
                'MailboxPrimarySmtp'     = $Mailbox.PrimarySmtpAddress; 
                'MailboxType'            = $Mailbox.RecipientTypeDetails; 
                'DelegationType'         = 'FullAccess'; 
                'DelegateType'           = $faUser.RecipientTypeDetails; 
                'DelegateDisplayName'    = $faUser.UserPrincipalName; 
                'DelegatePrimarySmtp'    = $faUser.WindowsEmailAddress; 
                'DelegateSamAccountName' = $faUser.SamAccountName; 
                'IsInherited'            = $fa.IsInherited; 
            } 
            $AllDelegations += $faDelegation 
        } 
    } 
    if ($SendAs) { 
 
        foreach ($sa in $SendAs) { 
 
            try { 
                $saUser = Get-User -Identity $sa.User.RawIdentity -ErrorAction SilentlyContinue 
                $saGroup = Get-Group -Identity $sa.User.RawIdentity -ErrorAction SilentlyContinue 
                if ($saGroup) {$saUser = $saGroup} 
            } 
            catch {} 
            $saDelegation = New-Object -TypeName PSObject -Property @{ 
                'MailboxDisplayName'     = $Mailbox.DisplayName; 
                'MailboxPrimarySmtp'     = $Mailbox.PrimarySmtpAddress; 
                'MailboxType'            = $Mailbox.RecipientTypeDetails; 
                'DelegationType'         = 'Send-As'; 
                'DelegateType'           = $saUser.RecipientTypeDetails; 
                'DelegateDisplayName'    = $saUser.DisplayName; 
                'DelegatePrimarySmtp'    = $saUser.WindowsEmailAddress; 
                'DelegateSamAccountName' = $saUser.SamAccountName; 
                'IsInherited'            = $sa.IsInherited; 
            } 
            $AllDelegations += $saDelegation 
        } 
    } 
    if ($SendOnBehalf) { 
 
        foreach ($sob in $SendOnBehalf) { 
 
            try { 
                $sobUser = Get-User -Identity $sob -ErrorAction SilentlyContinue 
                $sobGroup = Get-Group -Identity $sob -ErrorAction SilentlyContinue 
                if ($sobGroup) {$sobUser = $sobGroup} 
            } 
            catch {} 
            $sobDelegation = New-Object -TypeName PSObject -Property @{ 
                'MailboxDisplayName'     = $Mailbox.DisplayName; 
                'MailboxPrimarySmtp'     = $Mailbox.PrimarySmtpAddress; 
                'MailboxType'            = $Mailbox.RecipientTypeDetails; 
                'DelegationType'         = 'Send On Behalf'; 
                'DelegateType'           = $sobUser.RecipientTypeDetails; 
                'DelegateDisplayName'    = $sobUser.DisplayName; 
                'DelegatePrimarySmtp'    = $sobUser.WindowsEmailAddress; 
                'DelegateSamAccountName' = $sobUser.SamAccountName; 
                'IsInherited'            = 'N/A'; 
            } 
            $AllDelegations += $sobDelegation 
        } 
    } 
    Write-Output $AllDelegations | Select-Object MailboxDisplayName, 
    MailboxPrimarySmtp, 
    MailboxType, 
    DelegationType, 
    DelegateType, 
    DelegateDisplayName, 
    DelegatePrimarySmtp, 
    DelegateSamAccountName, 
    IsInherited 
} 
End {} 