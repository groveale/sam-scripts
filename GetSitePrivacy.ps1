#######################################
# Description: This script will build a list of site urls with privacy settings.
#
#              Auth required: Application permission - Reports.Read.All, Sites.Read.All
#
# Usage:       .\GetSitePrivacy.ps1
#
# Notes:       This script requires the MS Graph PowerShell module.
#              https://docs.microsoft.com/en-us/graph/powershell/installation
#
#  


##############################################
# Dependencies
##############################################

## Load the required modules

try {
    Import-Module Microsoft.Graph.Reports
    Import-Module Microsoft.Graph.Sites
}
catch {
    Write-Error "Error importing modules required modules - $($Error[0].Exception.Message))"
    Exit
}

##############################################
# Variables
##############################################

$clientId = "6c0f4f31-bf65-4b74-8b1c-9c038ea5c102"
$tenantId = "M365CPI77517573.onmicrosoft.com"


$thumbprint = "6ADC063641A24BB0BD68786AB71F07315CED9076"

$today = Get-Date

# Include date
$outputFilePath = "SitePrivacy-$($today.ToString('yyyy-MM-dd')).csv"

$tempDataDir = ".\Temp"

##############################################
# Functions
##############################################

function ConnectToMSGraph 
{  
    try{
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint -NoWelcome
    }   
    catch{
        Write-Host "Error connecting to MS Graph" -ForegroundColor Red
    }
}

function GetSPOReportData
{
    try 
    {      
        ## Check if file exists at path and if so exit
        if (Test-Path "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-SiteUsageDetail.csv")
        {
            Write-Host "SPO report data already exists for today. Skipping data pull." -ForegroundColor Yellow
            return
        }  
        Get-MgReportSharePointSiteUsageDetail -Period D180 -OutFile "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-SiteUsageDetail.csv"
    }
    catch
    {
        Write-Host "Error getting SPO report data - $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}

function GetGroupsReportData
{
    try 
    {
        ## Check if file exists at path and if so exit 
        if (Test-Path "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-GroupDetail.csv")
        {
            Write-Host "Group report data already exists for today. Skipping data pull." -ForegroundColor Yellow
            return
        }

        Get-MgReportOffice365GroupActivityDetail -Period D180 -OutFile "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-GroupDetail.csv"
    }
    catch
    {
        Write-Host "Error getting Groups report data - $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}

function GetGroupIdFromSiteId($siteId)
{
    try {
        $defaultDrive = Get-MgSiteDefaultDrive -SiteId $siteId -Select owner -ErrorAction Stop
        return $defaultDrive.Owner.AdditionalProperties.group.id  
    }
    catch {
        Write-Host "  Error getting group ID for site ID: $siteId - site is not group connected " -ForegroundColor Yellow
        return $null
    }
}

function GetSiteUrlFromSiteId($siteId)
{
    try {
        $site = Get-MgSite -SiteId $siteId -Select WebUrl -ErrorAction Stop
        return $site.WebUrl
    }
    catch {
        Write-Host "  Error getting site URL for site ID: $siteId - $($Error[0].Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function PopulateDataFrame($siteData, $groupsData)
{
    $dataFrame = @()

    foreach ($site in $siteData.Values)
    {

        ## Process each site
        Write-Host " Processing site: $($site.'Site Id')" -ForegroundColor White

        $siteUrl = GetSiteUrlFromSiteId $site.'Site Id'
        $groupId = GetGroupIdFromSiteId $site.'Site Id'

        if ($groupId -ne $null -and $groupsData.ContainsKey($groupId))
        {
            $group = $groupsData[$groupId]

            #Write-Host "  Group ID: $($group.'Group Id')" -ForegroundColor White

            $dataFrame += New-Object PSObject -Property @{
                'Site URL' = $siteUrl
                'Group ID' = $groupId
                'Group Member Count' = $group.'Member Count'
                'Site Privacy' = $group.'Group Type'
                'SPO Total File Count' = $site.'File Count'
                'SPO Active File Count' = $site.'Active File Count'
            }
        }
    }

    return $dataFrame
}

##############################################
# Main
##############################################

## Create temp data directory if it doesn't exist
if (-not (Test-Path $tempDataDir))
{
    New-Item -Path $tempDataDir -ItemType Directory
}

## Connect to MS Graph
ConnectToMSGraph

## Get Report data
GetSPOReportData
GetGroupsReportData

## Read data in as a hash table
$siteData = @{}
$groupsData = @{}

Import-Csv "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-SiteUsageDetail.csv" | Where-Object { $_.'Is Deleted' -eq $false } | ForEach-Object {
    $siteData[$_.'Site Id'] = $_
}

Import-Csv "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-GroupDetail.csv" | Where-Object { $_.'Is Deleted' -eq $false } | ForEach-Object {
    $groupsData[$_.'Group Id'] = $_
}

## remove any deleted rows
# $siteData = $siteData | Where { $_.'Is Deleted' -eq $false }
# $groupsData = $groupsData | Where { $_.'Is Deleted' -eq $false }


## We have x sites and z groups
Write-Host "Groups data count: $($groupsData.Count)" 
Write-Host "SPO Sites count: $($siteData.Count)"

Write-Host "Processing site data..."

$dataFrame = PopulateDataFrame $siteData $groupsData

## Output to CSV
$dataFrame | Export-Csv $outputFilePath -NoTypeInformation

##############################################
# End