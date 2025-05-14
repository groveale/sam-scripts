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

$adminUrl = "https://m365cpi77517573-admin.sharepoint.com"
$siteListCSV = "SiteList.csv"

# Set to $true to apply RCD to sites, false to remove RCD
$setSetRCD = $true

# Set to $true to validate the RCD settings of sites (will not change)
$validateRCDSettings = $true

##############################################
# Functions
##############################################

function ConnectToSPO() {
    try {
        Connect-SPOService -Url $adminUrl
        Write-Log "Connected to SharePoint Online" -Level INFO
    }
    catch {
        Write-Log "Error connecting to SharePoint Online: $($_.Exception.Message)" -Level ERROR
        exit
    }
    
}

function SetRCD($siteUrl) {
    try {
        Set-SPOSite -Identity $siteUrl -RestrictContentOrgWideSearch $setSetRCD
        Write-Log "Site: $($site.Url) - RCD set to: $rcd" -Level INFO
    }
    catch {
        Write-Log "Error setting RCD for site: $($site.Url) - $($_.Exception.Message)" -Level ERROR
        return
    }
   
}

function ValidateRCD ($siteUrl) {
    try {
        $rcd = Get-SPOSite -Identity $siteUrl | Select RestrictContentOrgWideSearch
        Write-Log "Site: $($site.Url) - RCD: $rcd" -Level INFO
        return $rcd.RestrictContentOrgWideSearch
    }
    catch {
        Write-Log "Error validating RCD for site: $($site.Url) - $($_.Exception.Message)" -Level ERROR
        return
    }
    
}

function Write-Log {
    <#
    .SYNOPSIS
        Simple function to write log entries to a file.
    .DESCRIPTION
        Writes timestamped log entries to a specified file.
    .PARAMETER Message
        The message to log.
    .PARAMETER LogFile
        Path to the log file. If not specified, logs to "[script_name]_log.txt" in the script directory.
    .PARAMETER Level
        The severity level (INFO, WARNING, ERROR). Defaults to INFO.
    .EXAMPLE
        Write-Log "Process started" -Level INFO
    #>
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Message,
        
        [Parameter(Position=1)]
        [string]$LogFile,
        
        [Parameter(Position=2)]
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )
    
    # Set default log file path if not specified
    if (-not $LogFile) {
        $scriptName = if ($MyInvocation.ScriptName) {
            [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName)
        } else {
            "PowerShell"
        }
        $LogFile = Join-Path $PWD.Path "$scriptName`_log.txt"
    }
    
    # Create timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Format log entry
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file (create if it doesn't exist)
    try {
        $logEntry | Out-File -FilePath $LogFile -Append -Encoding utf8
    }
    catch {
        Write-Error "Failed to write to log file: $_"
    }
    
    # Also write to console with color based on level
    switch ($Level) {
        'INFO'    { Write-Host $logEntry -ForegroundColor White }
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
    }
}

##############################################
# Main
##############################################

Write-Log "Script started" -Level INFO
Write-Log "Connecting to SharePoint Online" -Level INFO

if ($validateRCDSettings)
{
    Write-Log "Validating RCD settings for sites - No changes will be made" -Level WARNING
} else {
    Write-Log "Setting RCD for sites" -Level WARNING
}

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
        
    } else {
        # Set RCD
        SetRCD $site.Url
        
    }
}

Write-Progress -Activity "Processing Site $currentItem / $($sites.Count)" -Status "100% Complete:" -PercentComplete 100

Write-Log "Script complete" -Level INFO

