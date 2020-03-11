<# 
    README: This script used Out-GridView, and won't display properly using vscode. 
            So run this from PowerShell console. 
            
    CSV: Alias,Start,End
    E.g., Sharan.Gopi@amtrustgroup.com,1/25/2019 8:10AM,1/28/2019 8:10AM 
#>

$Import = Import-CSV ([Environment]::GetFolderPath("Desktop")+'\mailboxes.csv')
foreach ($mailbox in $Import) {
    $scriptBlock = {
        param (
            [datetime]$dtStart,
            [datetime]$dtEnd,
            [string]$strAlias
        )
        Get-TransportService | Get-MessageTrackingLog -ResultSize Unlimited -EventId RESUBMIT -Start $dtStart -End $dtEnd -Recipients $strAlias | Select-Object Timestamp,Serverhostname,Source,EventId,Sender,Recipients,MessageSubject
    }
    $Title = 'Exchange Transport Logs FWD - ' + $mailbox.Alias
    Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $mailbox.Start, $mailbox.End, $mailbox.Alias | Select-Object Timestamp,Serverhostname,Source,EventId,Sender,Recipients,MessageSubject | Sort-Object -Property Timestamp | Out-GridView -Title $Title
}
Write-Host ('Done !')