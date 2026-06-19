#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Teams via Graph REST API for deep details of all Teams.
.DESCRIPTION
    Retrieves full details of all Teams including visibility, owners, member and guest counts,
    channels (and channel types), and guest permissions/settings. Exports results to JSON/CSV.
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

# 2. Query all groups that are Teams
Write-Verbose "Querying Microsoft Teams list from Graph..."
$teamsList = [System.Collections.Generic.List[PSCustomObject]]::new()
$teamsUri = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/any(x:x eq 'Team')&`$select=id,displayName,visibility,createdDateTime"

try {
    while ($teamsUri) {
        $teamsResponse = Invoke-MgGraphRequest -Method GET -Uri $teamsUri -ErrorAction Stop
        if ($teamsResponse -and $teamsResponse.value) {
            foreach ($team in $teamsResponse.value) {
                $tId = $team.id
                
                # Query owners
                $owners = @()
                try {
                    $ownerRes = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$tId/owners?`$select=displayName,userPrincipalName" -ErrorAction SilentlyContinue
                    if ($ownerRes -and $ownerRes.value) {
                        foreach ($o in $ownerRes.value) {
                            $oName = if ($o.userPrincipalName) { $o.userPrincipalName } else { $o.displayName }
                            if ($oName) { $owners += $oName }
                        }
                    }
                } catch {}

                # Query members to classify counts
                $memberCount = 0
                $guestCount = 0
                try {
                    $memberUri = "https://graph.microsoft.com/v1.0/groups/$tId/members?`$select=displayName,userPrincipalName,userType"
                    while ($memberUri) {
                        $memberRes = Invoke-MgGraphRequest -Method GET -Uri $memberUri -ErrorAction Stop
                        if ($memberRes -and $memberRes.value) {
                            foreach ($m in $memberRes.value) {
                                $memberCount++
                                if ($m.userType -eq 'Guest' -or $m.userPrincipalName -match '#EXT#') {
                                    $guestCount++
                                }
                            }
                            $memberUri = $memberRes.'@odata.nextLink'
                        } else {
                            $memberUri = $null
                        }
                    }
                } catch {}

                # Query channels
                $channels = @()
                try {
                    $channelRes = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$tId/channels?`$select=displayName,membershipType" -ErrorAction SilentlyContinue
                    if ($channelRes -and $channelRes.value) {
                        foreach ($c in $channelRes.value) {
                            $type = if ($c.membershipType) { $c.membershipType.ToString() } else { "standard" }
                            $channels += "$($c.displayName):($type)"
                        }
                    }
                } catch {}

                # Query Guest settings
                $allowCreateUpdateChannels = $false
                $allowDeleteChannels = $false
                try {
                    $settingsRes = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$tId" -ErrorAction SilentlyContinue
                    if ($settingsRes -and $settingsRes.guestSettings) {
                        $allowCreateUpdateChannels = if ($null -ne $settingsRes.guestSettings.allowCreateUpdateChannels) { $settingsRes.guestSettings.allowCreateUpdateChannels } else { $false }
                        $allowDeleteChannels = if ($null -ne $settingsRes.guestSettings.allowDeleteChannels) { $settingsRes.guestSettings.allowDeleteChannels } else { $false }
                    }
                } catch {}

                $teamObj = [PSCustomObject]@{
                    Id                           = $tId
                    DisplayName                  = $team.displayName
                    Visibility                   = $team.visibility
                    CreatedDateTime              = $team.createdDateTime
                    Owners                       = $owners
                    MemberCount                  = $memberCount
                    GuestCount                   = $guestCount
                    Channels                     = $channels
                    AllowGuestCreateUpdateChannels = $allowCreateUpdateChannels
                    AllowGuestDeleteChannels     = $allowDeleteChannels
                }
                $teamsList.Add($teamObj)
            }
            $teamsUri = $teamsResponse.'@odata.nextLink'
        } else {
            $teamsUri = $null
        }
    }
} catch {
    Write-Error "Failed to query teams: $_"
    return
}

# 3. Convert Lists to flat strings for CSV export
$csvFormattedList = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($t in $teamsList) {
    $csvObj = $t.PSObject.Copy()
    $csvObj.Owners = $t.Owners -join '; '
    $csvObj.Channels = $t.Channels -join '; '
    $csvFormattedList.Add($csvObj)
}

# 4. Handle Outputs
if ($AiAgentMode) {
    $teamsList.ToArray() | ConvertTo-Json -Depth 5
} else {
    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath "../../../exports/Teams/TeamsFullDetails"
    if (!(Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path -Path $exportDir -ChildPath "TeamsFullDetails_${timestamp}.json"
    $csvPath  = Join-Path -Path $exportDir -ChildPath "TeamsFullDetails_${timestamp}.csv"

    $teamsList.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8
    $csvFormattedList.ToArray() | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    Write-Host "Teams Full Details JSON written to: $jsonPath" -ForegroundColor Green
    Write-Host "Teams Full Details CSV written to: $csvPath" -ForegroundColor Green
}
