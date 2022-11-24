$MGGroupIDs = Get-MgGroup
Write-Host "List all groups followed by members of the group" -ForegroundColor Yellow
foreach ($MGGroupID in $MGGroupIDs) {
    Write-Host "-------------------------"
    Write-Host "'"($MGGroupID.displayname)"' Group"
    Write-Host "-------------------------"
    $MGMembers = Get-MgGroupMember $MGGroupID | Select-Object -ExpandProperty Id
    foreach ($MGMember in $MGMembers) {
        Get-MgUser -UserId $MGMember  | Select-Object -ExpandProperty UserPrincipalName
    }
}


Get-MgUser | Select-Object DisplayName, UserPrincipalName, LicenseDetails   