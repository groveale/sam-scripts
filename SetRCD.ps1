#############################################
# Description: This script will get site acess data from the underlying SharePoint
#              and group permissions.
#              The script will create a CSV line for each owner/member/visitors of 
#              a SPO or M365 Group detaling the site access and permissions.
#              The first version will not look at broken inheritance.
#              
#
# Alex Grover - alexgrover@microsoft.com
#
#
##############################################
# Dependencies
##############################################

try {
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
}
catch {
    Write-Error "Error importing modules required modules - $($Error[0].Exception.Message))"
    Exit
}

##############################################
# Variables
##############################################

$adminUrl = "https://m365cpi28027490-admin.sharepoint.com"
$siteListCSV = "SiteList.csv"

# Set to $true to apply RCD to sites, false to remove RCD
$setSetRCD = $true

# Set to $true to validate the RCD settings of sites (will not change)
$validateRCDSettings = $true

##############################################
# Functions
##############################################

function ConnectToSPO() {
    Connect-SPOService -Url $adminUrl
}

function SetRCD($siteUrl) {
    Set-SPOSite -Identity $siteUrl -RestrictContentOrgWideSearch $setSetRCD
}

function ValidateRCD ($siteUrl) {
    $rcd = Get-SPOSite -Identity $siteUrl | Select RestrictContentOrgWideSearch
    return $rcd.RestrictContentOrgWideSearch
}

##############################################
# Main
##############################################

# Connect to SPO
ConnectToSPO

# Get all sites from CSV
$sites = Import-Csv $siteListCSV

$currentItem = 0

## initilise progress bar
$percent = 0
Write-Progress -Activity "Processing Site $currentItem / $($sites.Count)" -Status "$percent% Complete:" -PercentComplete $percent

foreach ($site in $sites) {
    $currentItem++
    $percent = [math]::Round(($currentItem / $sites.Count) * 100)
    Write-Progress -Activity "Processing Site $currentItem / $($sites.Count)" -Status "$percent% Complete:" -PercentComplete $percent

    if ($validateRCDSettings) {
        $rcd = ValidateRCD $site.Url
        Write-Host "Site: $($site.Url) - RCD: $rcd"
    } else {
        # Set RCD
        SetRCD $site.Url
    }

    
}

Write-Progress -Activity "Processing Site $currentItem / $($sites.Count)" -Status "100% Complete:" -PercentComplete 100

Write-Host "Script complete"

