#Requires -Version 7.0
<#
.SYNOPSIS
    Queries SharePoint Online via Graph REST API for deep details of Site Collections.
.DESCRIPTION
    Retrieves full details of all SharePoint Online site collections including creation date,
    webUrl, last modified date, and detailed permission / sharing state where available.
    Exports results to JSON/CSV.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$CertificateBase64,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$AiAgentMode
)

$ErrorActionPreference = 'Stop'

# Import our GRC Common module
$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath "../../../common/GRC-M365-Common.psm1"
if (Test-Path $commonModulePath) {
    Import-Module -Name $commonModulePath -Force
} else {
    Write-Error "Required GRC Common module not found at: $commonModulePath"
    return
}

# 1. Establish connection to Graph
try {
    if ($Interactive) {
        Connect-GRCEnvironment -Interactive
    } else {
        Connect-GRCEnvironment -TenantId $TenantId -ClientId $ClientId -CertificateBase64 $CertificateBase64
    }
} catch {
    Write-Error "Graph authentication failed: $_"
    return
}

# 2. Retrieve Site Collections from Graph
Write-Verbose "Querying all SharePoint sites via Graph REST..."
$sitesList = [System.Collections.Generic.List[PSCustomObject]]::new()
$sitesUri = "https://graph.microsoft.com/v1.0/sites?`$select=id,name,displayName,webUrl,createdDateTime,lastModifiedDateTime"

try {
    while ($sitesUri) {
        $sitesResponse = Invoke-MgGraphRequest -Method GET -Uri $sitesUri -ErrorAction Stop
        if ($sitesResponse -and $sitesResponse.value) {
            foreach ($site in $sitesResponse.value) {
                $siteId = $site.id
                
                # Try to get storage usage and quota from root drive
                $storageUsed = 0
                $storageQuota = 0
                try {
                    $driveResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drive" -ErrorAction SilentlyContinue
                    if ($driveResponse -and $driveResponse.quota) {
                        $storageUsed = $driveResponse.quota.used
                        $storageQuota = $driveResponse.quota.total
                    }
                } catch {}

                # Query Site Collection Owners/Admins via site permissions
                $owners = @()
                try {
                    $permRes = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions" -ErrorAction SilentlyContinue
                    if ($permRes -and $permRes.value) {
                        foreach ($p in $permRes.value) {
                            if ($p.roles -contains 'owner' -or $p.roles -contains 'admin') {
                                if ($p.grantedToV2 -and $p.grantedToV2.user) {
                                    $owners += $p.grantedToV2.user.displayName
                                } elseif ($p.grantedTo -and $p.grantedTo.user) {
                                    $owners += $p.grantedTo.user.displayName
                                }
                            }
                        }
                    }
                } catch {}

                $siteObj = [PSCustomObject]@{
                    Id                           = $siteId
                    Name                         = $site.name
                    DisplayName                  = $site.displayName
                    WebUrl                       = $site.webUrl
                    CreatedDateTime              = $site.createdDateTime
                    LastModifiedDateTime         = $site.lastModifiedDateTime
                    StorageUsedBytes             = $storageUsed
                    StorageQuotaBytes            = $storageQuota
                    SiteOwners                   = $owners
                }
                $sitesList.Add($siteObj)
            }
            $sitesUri = $sitesResponse.'@odata.nextLink'
        } else {
            $sitesUri = $null
        }
    }
} catch {
    Write-Error "Failed to retrieve SharePoint sites: $_"
    return
}

# 3. Convert Lists to flat strings for CSV export
$csvFormattedList = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($s in $sitesList) {
    $csvObj = $s.PSObject.Copy()
    $csvObj.SiteOwners = $s.SiteOwners -join '; '
    $csvFormattedList.Add($csvObj)
}

# 4. Handle Outputs
if ($AiAgentMode) {
    $sitesList.ToArray() | ConvertTo-Json -Depth 5
} else {
    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath "../../../exports/SharePoint/SharePointFullDetails"
    if (!(Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path -Path $exportDir -ChildPath "SharePointFullDetails_${timestamp}.json"
    $csvPath  = Join-Path -Path $exportDir -ChildPath "SharePointFullDetails_${timestamp}.csv"

    $sitesList.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8
    $csvFormattedList.ToArray() | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    Write-Host "SharePoint Full Details JSON written to: $jsonPath" -ForegroundColor Green
    Write-Host "SharePoint Full Details CSV written to: $csvPath" -ForegroundColor Green
}
