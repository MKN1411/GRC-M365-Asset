#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Exchange Online for deep details of all Mailboxes.
.DESCRIPTION
    Retrieves full mailbox settings including quotas, email forwarding configurations,
    and detailed delegation permissions (Full Access, Send As, Send on Behalf).
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

# 1. Establish connection to Exchange Online
$ippsConnected = $false
try {
    if ($Interactive) {
        Connect-GRCExchange -Interactive
    } else {
        Connect-GRCExchange -TenantId $TenantId -ClientId $ClientId -CertificateBase64 $CertificateBase64
    }
    $ippsConnected = $true
} catch {
    Write-Error "Exchange Online authentication failed: $_"
    return
}

try {
    # 2. Retrieve Mailboxes
    Write-Verbose "Retrieving Mailbox list from Exchange Online..."
    $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    $exchangeDetailsList = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($mb in $mailboxes) {
        $upn = $mb.UserPrincipalName
        $smtp = $mb.PrimarySmtpAddress
        
        # Query Mailbox Permissions (Full Access)
        $fullAccessList = @()
        try {
            $perms = Get-MailboxPermission -Identity $smtp -ErrorAction SilentlyContinue
            if ($perms) {
                foreach ($p in $perms) {
                    # Filter out system and self accounts to show real delegation
                    if ($p.User -notmatch 'NT AUTHORITY\\SELF|MicrosoftExchange|System|Administrator|Domain Admins' -and $p.IsInherited -eq $false) {
                        $fullAccessList += "$($p.User):($($p.AccessRights -join ','))"
                    }
                }
            }
        } catch {}

        # Query Recipient Permissions (Send As)
        $sendAsList = @()
        try {
            $sendAs = Get-RecipientPermission -Identity $smtp -ErrorAction SilentlyContinue
            if ($sendAs) {
                foreach ($sa in $sendAs) {
                    if ($sa.Trustee -notmatch 'NT AUTHORITY\\SELF|MicrosoftExchange|System|Administrator|Domain Admins') {
                        $sendAsList += $sa.Trustee
                    }
                }
            }
        } catch {}

        # Query Send on Behalf
        $sendOnBehalfList = @()
        if ($mb.GrantSendOnBehalfTo) {
            foreach ($sob in $mb.GrantSendOnBehalfTo) {
                $sendOnBehalfList += $sob
            }
        }

        # Forwarding details
        $forwardingAddress = if ($mb.ForwardingAddress) { $mb.ForwardingAddress.ToString() } else { "" }
        $forwardingSmtpAddress = if ($mb.ForwardingSmtpAddress) { $mb.ForwardingSmtpAddress.ToString() } else { "" }
        $deliverToMailboxAndForward = $mb.DeliverToMailboxAndForward

        # Inbox rules forwarding analysis (Check if they have any rules forwarding external mail)
        $hasForwardingRules = $false
        $forwardingRuleDetails = @()
        try {
            $rules = Get-InboxRule -Mailbox $smtp -ErrorAction SilentlyContinue
            if ($rules) {
                foreach ($r in $rules) {
                    if ($r.RedirectTo -or $r.ForwardTo -or $r.ForwardAsAttachmentTo) {
                        $hasForwardingRules = $true
                        $targets = @($r.RedirectTo + $r.ForwardTo + $r.ForwardAsAttachmentTo) -join ','
                        $forwardingRuleDetails += "$($r.Name):(Redirect/Forward to $targets)"
                    }
                }
            }
        } catch {}

        $mbObj = [PSCustomObject]@{
            MailboxAddress              = $smtp.ToString()
            DisplayName                 = $mb.DisplayName
            MailboxType                 = $mb.RecipientTypeDetails.ToString()
            ProhibitSendReceiveQuota    = if ($mb.ProhibitSendReceiveQuota) { $mb.ProhibitSendReceiveQuota.ToString() } else { "Unlimited" }
            ProhibitSendQuota           = if ($mb.ProhibitSendQuota) { $mb.ProhibitSendQuota.ToString() } else { "Unlimited" }
            UseDatabaseQuotaDefaults    = $mb.UseDatabaseQuotaDefaults
            FullAccessPermissions       = $fullAccessList
            SendAsPermissions           = $sendAsList
            SendOnBehalfPermissions     = $sendOnBehalfList
            ForwardingAddress           = if ($forwardingSmtpAddress) { $forwardingSmtpAddress } else { $forwardingAddress }
            DeliverToMailboxAndForward  = $deliverToMailboxAndForward
            HasForwardingRules          = $hasForwardingRules
            ForwardingRuleDetails       = $forwardingRuleDetails
        }
        $exchangeDetailsList.Add($mbObj)
    }

    # Convert Lists to flat strings for CSV export
    $csvFormattedList = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($e in $exchangeDetailsList) {
        $csvObj = $e.PSObject.Copy()
        $csvObj.FullAccessPermissions = $e.FullAccessPermissions -join '; '
        $csvObj.SendAsPermissions = $e.SendAsPermissions -join '; '
        $csvObj.SendOnBehalfPermissions = $e.SendOnBehalfPermissions -join '; '
        $csvObj.ForwardingRuleDetails = $e.ForwardingRuleDetails -join '; '
        $csvFormattedList.Add($csvObj)
    }

    # 3. Handle Outputs
    if ($AiAgentMode) {
        $exchangeDetailsList.ToArray() | ConvertTo-Json -Depth 5
    } else {
        $exportDir = Join-Path -Path $PSScriptRoot -ChildPath "../../../exports/ExchangeOnline/ExchangeFullDetails"
        if (!(Test-Path $exportDir)) {
            New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
        }
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $jsonPath = Join-Path -Path $exportDir -ChildPath "ExchangeFullDetails_${timestamp}.json"
        $csvPath  = Join-Path -Path $exportDir -ChildPath "ExchangeFullDetails_${timestamp}.csv"

        $exchangeDetailsList.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8
        $csvFormattedList.ToArray() | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

        Write-Host "Exchange Online Full Details JSON written to: $jsonPath" -ForegroundColor Green
        Write-Host "Exchange Online Full Details CSV written to: $csvPath" -ForegroundColor Green
    }

} finally {
    if ($ippsConnected) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
}
