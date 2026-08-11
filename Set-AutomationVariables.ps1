# Set-AutomationVariables.ps1
# ---------------------------
# Run this script once after Step 10 (modules imported) to create all Automation Variables
# in both the provisioning and reporting Automation Accounts (Step 11 of the migration TODO).
#
# Prerequisites:
#   Connect-AzAccount                                         # already done in Step 0
#   Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All" # already done in Step 0
#
# Usage:
#   .\Set-AutomationVariables.ps1
#
# To update a single variable later (e.g. add a recipient) use the snippet at the bottom.
# This script is safe to re-run — Set-AzAutomationVariable overwrites existing values.

# ==============================================================================
# SECTION 1 — Infrastructure names
# Must match what you used in Steps 1–2. Edit here if you used different names.
# ==============================================================================

$rg               = "rg-exo-automation"
$provAccountName  = "aa-exo-provisioning"
$repAccountName   = "aa-exo-reporting"
$storageAcctName  = "voristexoreports"        # must match the storage account created in Step 8
$storageContainer = "reports"

# ==============================================================================
# SECTION 2 — Tenant-specific values   <-- FILL THESE IN BEFORE RUNNING
# ==============================================================================

$tenantId           = "f641d663-a36b-443a-b2f0-7dcc69a4af06"  # Azure AD tenant ID (Entra Overview)
$organization       = "MngEnvMCAP545510.onmicrosoft.com"               # .onmicrosoft.com domain
$senderEmail        = "admin@MngEnvMCAP545510.onmicrosoft.com"       # licensed mailbox — FROM address
$recipientEmail     = "admin@MngEnvMCAP545510.onmicrosoft.com"                   # provisioning completion notification
$retentionPolicy    = "Two Year Retention"                    # exact name of the EXO retention policy

# Report recipients — JSON array; edit here or use the update snippet at the bottom
$reportRecipients   = '["admin@MngEnvMCAP545510.onmicrosoft.com"]'

# Litigation Hold duration — integer days (e.g. 2555 = 7 years, 1825 = 5 years) or "Unlimited".
# Leave empty to default to Unlimited; an empty Automation Variable will be written.
$litigationHoldDuration = "2555"

# Mailbox creation date filter — ISO 8601 date (e.g. "2025-01-01").
# Only mailboxes created on or after this date will be provisioned.
# Leave empty to process all mailboxes.
$provisioningCreatedAfter = "2026-03-17"

# Optional mailbox exclusions and live-run safety limit.
# All settings are written so the Automation Account has a consistent variable set.
$provisioningExcludeUpnRegex          = "" # Example: '^(?=[^@]*[A-Za-z])(?=[^@]*\d)[^@]+@'
$provisioningExcludeUpns              = '[]' # Example: '["svc-backup@contoso.com","admin@contoso.com"]'
$provisioningExcludeRetentionPolicies = '[]' # Example: '["VIP Retention Policy","Legal Hold Policy"]'
$provisioningMaximumChanges           = "" # Example: "500"

# ==============================================================================
# SECTION 3 — Confirm context before writing
# ==============================================================================

$ctx = Get-AzContext
$requiresLogin = -not $ctx -or $ctx.Tenant.Id -ne $tenantId
if (-not $requiresLogin) {
    try {
        Get-AzAccessToken -ResourceUrl "https://management.azure.com/" -ErrorAction Stop | Out-Null
    } catch {
        $requiresLogin = $true
        Write-Warning "The cached Azure context token is expired or unavailable."
    }
}

if ($requiresLogin) {
    Write-Host "Signing in to Azure tenant $tenantId..." -ForegroundColor Yellow
    try {
        Connect-AzAccount -Tenant $tenantId -ErrorAction Stop | Out-Null
        Get-AzAccessToken -ResourceUrl "https://management.azure.com/" -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext
    } catch {
        throw "Azure authentication failed or was canceled. Run 'Connect-AzAccount -Tenant $tenantId' and retry. Detail: $($_.Exception.Message)"
    }
}
Write-Host ""
Write-Host "Azure context : $($ctx.Subscription.Name) ($($ctx.Subscription.Id))"
Write-Host "Resource group: $rg"
Write-Host "Prov account  : $provAccountName"
Write-Host "Rep account   : $repAccountName"
Write-Host ""

$confirm = Read-Host "Continue? [Y/N]"
if ($confirm -notmatch "^[Yy]") {
    Write-Host "Aborted."
    exit
}

# ==============================================================================
# SECTION 4 — Provisioning account variables
# ==============================================================================

Write-Host "`nWriting provisioning account variables..." -ForegroundColor Cyan

$provVars = [ordered]@{
    TenantId             = $tenantId
    Organization         = $organization
    SenderEmail          = $senderEmail
    RecipientEmail       = $recipientEmail
    RetentionPolicyName  = $retentionPolicy
    ReportingAccountRG   = $rg
    ReportingAccountName = $repAccountName
    StorageAccountName   = $storageAcctName   # provisioning runbook uploads its log CSV to the same container
    StorageContainer     = $storageContainer
    LitigationHoldDuration                = $litigationHoldDuration
    ProvisioningCreatedAfter              = $provisioningCreatedAfter
    ProvisioningExcludeUpnRegex           = $provisioningExcludeUpnRegex
    ProvisioningExcludeUpns               = $provisioningExcludeUpns
    ProvisioningExcludeRetentionPolicies  = $provisioningExcludeRetentionPolicies
    ProvisioningMaximumChanges            = $provisioningMaximumChanges
}

function Set-OrNewAutomationVariable {
    param(
        [string]$ResourceGroupName,
        [string]$AutomationAccountName,
        [string]$Name,
        [string]$Value
    )
    $existing = Get-AzAutomationVariable `
        -ResourceGroupName     $ResourceGroupName `
        -AutomationAccountName $AutomationAccountName `
        -Name                  $Name `
        -ErrorAction           SilentlyContinue
    if ($existing) {
        if ($existing.Encrypted) {
            # Azure won't let you change encryption in-place — remove and recreate
            Remove-AzAutomationVariable `
                -ResourceGroupName     $ResourceGroupName `
                -AutomationAccountName $AutomationAccountName `
                -Name                  $Name
            New-AzAutomationVariable `
                -ResourceGroupName     $ResourceGroupName `
                -AutomationAccountName $AutomationAccountName `
                -Name                  $Name `
                -Value                 $Value `
                -Encrypted             $false
        } else {
            Set-AzAutomationVariable `
                -ResourceGroupName     $ResourceGroupName `
                -AutomationAccountName $AutomationAccountName `
                -Name                  $Name `
                -Value                 $Value `
                -Encrypted             $false
        }
    } else {
        New-AzAutomationVariable `
            -ResourceGroupName     $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -Name                  $Name `
            -Value                 $Value `
            -Encrypted             $false
    }
}

foreach ($v in $provVars.GetEnumerator()) {
    Set-OrNewAutomationVariable -ResourceGroupName $rg -AutomationAccountName $provAccountName -Name $v.Key -Value $v.Value
    Write-Host "  [prov] $($v.Key) = $($v.Value)"
}

# ==============================================================================
# SECTION 5 — Reporting account variables
# ==============================================================================

Write-Host "`nWriting reporting account variables..." -ForegroundColor Cyan

$repVars = [ordered]@{
    TenantId              = $tenantId
    Organization          = $organization
    SenderEmail           = $senderEmail
    StorageAccountName    = $storageAcctName
    StorageContainer      = $storageContainer
    ReportRecipients      = $reportRecipients   # JSON array — runbook parses with ConvertFrom-Json
    AutomationAccountRG   = $rg                 # Generate-MailboxReport uses this to call Start-AzAutomationRunbook
    AutomationAccountName = $repAccountName     # Generate-MailboxReport uses this to call Start-AzAutomationRunbook
}

foreach ($v in $repVars.GetEnumerator()) {
    Set-OrNewAutomationVariable -ResourceGroupName $rg -AutomationAccountName $repAccountName -Name $v.Key -Value $v.Value
    Write-Host "  [rep]  $($v.Key) = $($v.Value)"
}

# ==============================================================================
# SECTION 6 — Verify
# ==============================================================================

Write-Host "`nVerifying..." -ForegroundColor Cyan

foreach ($account in @($provAccountName, $repAccountName)) {
    Write-Host "`n  $account"
    Get-AzAutomationVariable -ResourceGroupName $rg -AutomationAccountName $account |
        Select-Object Name, Value |
        Format-Table -AutoSize
}

Write-Host "Done. All variables set.`n" -ForegroundColor Green

# ==============================================================================
# SNIPPET — Update report recipients later (run independently, not part of main flow)
# ==============================================================================
<#
Set-AzAutomationVariable `
    -ResourceGroupName     $rg `
    -AutomationAccountName $repAccountName `
    -Name                  "ReportRecipients" `
    -Encrypted             $false `
    -Value                 '["it-team@contoso.com","manager@contoso.com","ciso@contoso.com"]'
#>

# ==============================================================================
# SNIPPET — Set or update LitigationHoldDuration later (run independently)
# ==============================================================================
<#
Set-AzAutomationVariable `
    -ResourceGroupName     $rg `
    -AutomationAccountName $provAccountName `
    -Name                  "LitigationHoldDuration" `
    -Encrypted             $false `
    -Value                 "2555"   # days — or "Unlimited"
#>

# ==============================================================================
# SNIPPET — Set ProvisioningCreatedAfter to process only mailboxes created on or
#           after a given date (ISO 8601, e.g. "2025-01-01"). Remove the variable
#           entirely to revert to processing all mailboxes.
# ==============================================================================
<#
Set-AzAutomationVariable `
    -ResourceGroupName     $rg `
    -AutomationAccountName $provAccountName `
    -Name                  "ProvisioningCreatedAfter" `
    -Encrypted             $false `
    -Value                 "2025-01-01"

# To remove the filter and process all mailboxes again:
# Remove-AzAutomationVariable -ResourceGroupName $rg -AutomationAccountName $provAccountName -Name "ProvisioningCreatedAfter"
#>

# ==============================================================================
# SNIPPET — Run a filter evaluation with ad hoc overrides.
# Omitted settings continue to use their Automation Variables.
# ==============================================================================
<#
Start-AzAutomationRunbook `
    -ResourceGroupName     $rg `
    -AutomationAccountName $provAccountName `
    -Name                  "Invoke-ScheduledMailboxProvisioning" `
    -Parameters            @{
        WhatIf                                      = $true
        ProvisioningCreatedAfter                    = "2026-01-01"
        ProvisioningExcludeUpns                     = '["temporary-skip@contoso.com"]'
        ProvisioningExcludeRetentionPolicies        = '[]'
        ProvisioningMaximumChanges                  = "1000"
        RetentionPolicyName                         = "Two Year Retention"
        LitigationHoldDuration                      = "2555"
    }
#>
