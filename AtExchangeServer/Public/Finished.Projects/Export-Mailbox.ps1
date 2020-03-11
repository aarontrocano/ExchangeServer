<#
    CSV format: Name
                sharan.gopi@amtrustservices.com
#>
$Import = Import-CSV ([Environment]::GetFolderPath("Desktop")+'\mailboxes.csv')
foreach ($Mailbox in $Import) {
    $Reports = "\\clepmail10\pst$\TOCOExports\$($Mailbox.Name)" + " " + (get-date -format "hhmm-MM-dd-yyyy") + ".pst"
    $CompletedMBX = ([Environment]::GetFolderPath("Desktop")+'\Completed.txt')
    $IncompleteMBX = ([Environment]::GetFolderPath("Desktop")+'\Incomplete.txt')
    Write-Host ('Exporting Mailbox for: ' + $Mailbox.name + ' ... ') -ForegroundColor Green
    $Export = New-MailboxExportRequest -Mailbox $Mailbox.Name -FilePath $Reports
        foreach ($Object in $Export ){
            start-sleep -seconds 20

            Write-Output "Waiting for Export to Complete"
            While ((Get-MailboxExportRequest -identity "$($Mailbox.Name)\MailboxExport" | 
                Where-Object {$_.Status -eq "Queued" -or $_.Status -eq "InProgress"})){
                start-sleep -seconds 60
                }    

            $Completed = Get-MailboxExportRequest -identity "$($Mailbox.Name)\MailboxExport" | 
            Where-object {$_.status -eq "Completed"}| Get-Mailboxexportrequeststatistics | 
            format-list

                if ($Completed){
                Write-Output "Writing Completed Mailbox Report to $CompletedMBX"
                $Completed | Out-File -FilePath $CompletedMBX -append 
                }

            $Incomplete = Get-MailboxExportRequest -identity "$($Mailbox.Name)\MailboxExport" | 
            Where-object {$_.status -ne "Completed"}| Get-Mailboxexportrequeststatistics | 
            format-list

                if ($Incomplete){
                Write-Output "Writing Incomplete Mailbox Report to $IncompleteMBX"
                $Incomplete | Out-File -FilePath $IncompleteMBX -append
                }

        }

}
