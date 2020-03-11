<#
    ALPHA testing for Ryan
#>
$import = import-csv ([Environment]::GetFolderPath("Desktop") + '\test.csv') 
$Enabled = ([Environment]::GetFolderPath("Desktop") + '\Enabled.csv')
$Disabled = ([Environment]::GetFolderPath("Desktop") + '\Disabled.csv')
$Failed = ([Environment]::GetFolderPath("Desktop") + '\NotFound.csv')
$EnabledReport = @()
$DisabledReport = @()
$ErrorReport = @()
<#
$scriptBlock = {
    param (
        [string]$strAlias
    )
    Get-Mailbox -Identity $strAlias | Select-Object PrimarySmtpAddress,RecipientType,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxandForward,ServerName,Database,LitigationHoldEnabled
}
#>
$scriptBlock2 = {
    param (
        [string]$strAlias
    ) 
    Get-MailboxAutoReplyConfiguration -ResultSize Unlimited -Identity $strAlias | Select-Object AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage
}
$count = 0
foreach ($mailbox in $import) {
    $MBX = $mailbox.Emailaddress
    <#
    $Result = Get-mailbox -identity $MBX # |Get-MailboxAutoReplyConfiguration | Where-Object { $_.AutoReplyState -ne "Enabled" } | Select-object Identity, StartTime, EndTime, AutoReplyState 
    #>
    <#
    $Result = Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $MBX 
    $State = Invoke-Command -Session $Session -ScriptBlock $scriptBlock2 -ArgumentList $MBX | Select-Object AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage
    #>
    $Result = Invoke-Command -Session $Session -ScriptBlock $scriptBlock2 -ArgumentList $MBX | Select-Object AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage
    if ($Result.AutoReplyState.Value -eq 'Scheduled') {
        $count++
        $EnabledReport += $Result | Select-Object -Property AutoReplyState,@{n="Identity";e={$MBX}},StartTime,EndTime
        Write-Host ("$($count)" + ': ' + 'Getting Enabled OOO Configuration for : ' + "$($MBX)")
    }
    elseif ($Result.AutoReplyState.Value -eq 'Disabled') {
        $count++
        $DisabledReport += $Result | Select-Object -Property AutoReplyState,@{n="Identity";e={$MBX}},StartTime,EndTime
        Write-Host ("$($count)" + ': ' + 'Getting Disabled OOO Configuration for : ' + "$($MBX)")     
    }
    else {
        $count++
        $ErrorReport += [pscustomobject]@{EmailAddress = $MBX}
        Write-Host ("$($count)" + ': ' + 'Mailbox was NOT found! : ' + "$($MBX)") 
    }         
}
$EnabledReport | Export-csv $Enabled
$DisabledReport | Export-csv $Disabled
$ErrorReport | Export-csv $Failed
Write-host ('All Mailboxes have been processed! ...')



if (($Result|Get-MailboxAutoReplyConfiguration).Enabled -eq $true) {
    $count++
    $EnabledReport += $Result
    Write-Host ("$($count)" + ': ' + 'Getting Enabled OOO Configuration for : ' + "$($MBX)")
}
elseif (($Result|Get-MailboxAutoReplyConfiguration).Enabled -eq $false) {
    $count++
    $DisabledReport += $Result
    Write-Host ("$($count)" + ': ' + 'Gettin Disabled OOO Configuration for : ' + "$($MBX)")     
}
else {
    $count++
    $ErrorReport += $MBX
    Write-Host ("$($count)" + ': ' + 'Mailbox was NOT found! : ' + "$($MBX)") 
}         


if ($Result.AutoReplyState.Value -eq 'Scheduled' -or $Result.AutoReplyState.Value -eq 'Enabled') {
    #do stuff
} elseif ($Result.AutoReplyState.Value -eq 'Disabled') {
    #do other stuff
} else {
    #catch-all
}

$Enabled = Get-MailboxAutoReplyConfiguration -identity $MBX | Where-Object { $_.AutoReplyState.Value -ne "Disabled" } | Select-object Identity,StartTime,EndTime,AutoReplyState
$Disabled = Get-MailboxAutoReplyConfiguration -identity $MBX | Where-Object { $_.AutoReplyState.Value -eq "Disabled" } | Select-object Identity,StartTime,EndTime,AutoReplyState
if ($Enabled) {

} elseif ($Disabled) {

} else {

}
