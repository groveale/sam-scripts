



##############################################
# Variables
##############################################

$clientId = "6c0f4f31-bf65-4b74-8b1c-9c038ea5c102"
$tenantId = "M365CPI77517573.onmicrosoft.com"


$thumbprint = "6ADC063641A24BB0BD68786AB71F07315CED9076"

$today = Get-Date
$todayString = $today.ToString('yyyy-MM-dd')

$outputFile = ".\SiteUsageReportWithUrls-$($todayString).csv"


##############################################
# Functions
##############################################

function GetSitesMSGraph ($sites)
{

    $allSites = Get-MgSite -All -Property WebUrl, SharePointIds, CreatedDateTime
    

    return $allSites
}

function GetSiteUsageReport($siteDetails) {

    $siteIds = $siteDetails.SharepointIds.SiteId
    $siteTemplates = $parametersList.siteTemplates.Value

    Get-MgReportSharePointSiteUsageDetail -Period D7 -OutFile .
    $siteUsage = Import-Csv .\SharePointSiteUsageDetail*.csv
    ## delete report file
    Remove-Item .\SharePointSiteUsageDetail*.csv

    ## Add site Url to the $siteUsage object
    foreach ($site in $siteUsage) {

        if ($site.'Site Url' -ne "") { continue }

        $siteId = $site.'Site Id'
        $siteUrl = $siteDetails | where { $_.SharePointIds.SiteId -eq $siteId } | Select -ExpandProperty WebUrl
        $site.'Site Url' = $siteUrl
    }

    ## Filter the deleted sites
    $siteUsage = $siteUsage | where { $_.'Is Deleted' -eq "FALSE" }

    if ($parametersList.allSites.Value -eq $true)
    {
        return $siteUsage | where { $siteTemplates -contains $_.'Root Web Template' }
    }

    return $siteUsage | where { $siteIds -contains $_.'Site Id' }
}

function ConnectToMSGraph 
{  
    try{
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint -NoWelcome
    }   
    catch{
        Write-Host "Error connecting to MS Graph" -ForegroundColor Red
    }
}


##############################################
# Main
##############################################

## Connect to MS Graph (to pull he report)
ConnectToMSGraph 

$siteDetails = GetSitesMSGraph $sites

## Get the site usage report for the sites we care about
$siteUsage = GetSiteUsageReport $siteDetails

## Save the report to a file
$siteUsage | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8