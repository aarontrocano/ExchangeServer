<#************************************************************************#>
$SamAccountName = 'usertest032'
$fname = '' 
$lname = 'UserTest032'
$DisplayName = <#$fname + ' ' +#> $lname; <#$dotName = $fname + '.' + $lname#>
$alias = New-Object -TypeName psobject 
$alias | Add-Member -MemberType NoteProperty -Name oldemail -Value "$($SamAccountName)@theamyntagroup.com"
#$alias | Add-Member -MemberType NoteProperty -Name newemail -Value "$($dotName)@amyntagroup.com"
<#************************************************************************#>
Enable-Mailbox <#-WhatIf#> -Identity $SamAccountName -Alias $SamAccountName -Confirm:$false -Database 'US-OH-DB04' -DisplayName $DisplayName -PrimarySmtpAddress $alias.oldemail
$scriptBlock = {
    param (
        [string]$strAlias,
        [string]$strFwd
    ) 
    Set-Mailbox -Identity $strAlias -DeliverToMailboxAndForward $false -ForwardingAddress $null -ForwardingSMTPAddress $strFwd 
} 
#Write-Verbose -Message "$($alias.oldemail) to $($alias.newemail)" -Verbose
#Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $alias.oldemail, $alias.newemail
Write-Verbose -Message 'Done !' -Verbose
<#************************************************************************#>