$Import = Import-Csv -Path ([Environment]::GetFolderPath("Desktop")+'\DisplayName.csv')
$report = @()
foreach ($i in $Import) {
    
    [string]$strImportName = $i.Name
    $objUser = Get-ADUser -Filter{displayName -eq $strImportName} -Properties DisplayName
    [string]$strDisplayName = $objUser.DisplayName
    [string]$strSamAccountName = $objUser.SamAccountName
    [string]$strErrorMessage = ("`"" + $strImportName + "`"" + ' with SamAccountName ' + "`"" + $strSamAccountName  + "`"" + ' not match ' + "`"" + $strDisplayName + "`"")
    #if (!( Get-Mailbox -Identity $i.Name )) { $report += $i.Name }  
    if ($false) { $report += $i.Name }  
    elseif ( ($strImportName) -eq $strDisplayName ) {} Else {Write-Host ($strErrorMessage)}
    if ( ($strImportName) -eq $strDisplayName ) {} Else {$report += $strErrorMessage}
} 
$out = $report
$out | Set-Content ([Environment]::GetFolderPath("Desktop")+'\BadDisplayNames.csv')