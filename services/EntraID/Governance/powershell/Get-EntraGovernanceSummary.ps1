#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Entra ID for Conditional Access policies, Access Packages, and Access Reviews.
.DESCRIPTION
    Retrieves key Entra ID Identity Governance metrics for GRC auditing.
    Exports structured JSON/CSV data.
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

# Import GRC Common library
$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath "../../../../common/GRC-M365-Common.psm1"
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

# 2. Query Governance & CA Settings using REST calls
$reportData = [Ordered]@{
    TotalConditionalAccessPolicies   = 0
    EnabledCAPoliciesCount           = 0
    ReportOnlyCAPoliciesCount        = 0
    ConditionalAccessPoliciesDetails = @()
    TotalAccessPackages              = 0
    AccessPackagesDetails            = @()
    TotalAccessReviews               = 0
    AccessReviewsDetails             = @()
}

try {
    # Query Conditional Access Policies
    $caResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -ErrorAction SilentlyContinue
    if ($caResponse -and $caResponse.value) {
        $reportData.TotalConditionalAccessPolicies = @($caResponse.value).Count
        $reportData.EnabledCAPoliciesCount = ($caResponse.value | Where-Object { $_.state -eq 'enabled' }).Count
        $reportData.ReportOnlyCAPoliciesCount = ($caResponse.value | Where-Object { $_.state -eq 'enabledForReportingButNotEnforced' }).Count
        
        # Collect details of CA Policies
        $reportData.ConditionalAccessPoliciesDetails = $caResponse.value | ForEach-Object {
            [Ordered]@{
                Id          = $_.id
                DisplayName = $_.displayName
                State       = $_.state
                Conditions  = $_.conditions
                GrantControls = $_.grantControls
            }
        }
    }

    # Query Access Packages (Identity Governance)
    $apResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages" -ErrorAction SilentlyContinue
    if ($apResponse -and $apResponse.value) {
        $reportData.TotalAccessPackages = @($apResponse.value).Count
        # Collect Access Packages details
        $reportData.AccessPackagesDetails = $apResponse.value | ForEach-Object {
            [Ordered]@{
                Id          = $_.id
                DisplayName = $_.displayName
                Description = $_.description
                IsHidden    = $_.isHidden
            }
        }
    }

    # Query Access Reviews
    $arResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions" -ErrorAction SilentlyContinue
    if ($arResponse -and $arResponse.value) {
        $reportData.TotalAccessReviews = @($arResponse.value).Count
        # Collect Access Reviews details
        $reportData.AccessReviewsDetails = $arResponse.value | ForEach-Object {
            [Ordered]@{
                Id          = $_.id
                DisplayName = $_.displayName
                Status      = $_.status
                Scope       = $_.scope
            }
        }
    }

} catch {
    Write-Warning "Could not query all Entra ID Governance endpoints: $_"
}

# 3. Handle Outputs based on execution scope
$exportObj = [PSCustomObject]$reportData
if ($AiAgentMode) {
    $exportObj | ConvertTo-Json -Depth 5
} else {
    Export-GRCAssetData -ServiceName "EntraID" -AssetName "EntraGovernanceSummary" -Data @($exportObj)
}
