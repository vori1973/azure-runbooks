#Requires -Version 7.4
<#
.SYNOPSIS
    Azure Automation runbook: enable archive, enable auto-expanding archive, and assign a
    retention policy to every licensed user mailbox in the tenant. Fully unattended.

.DESCRIPTION
    Adapted from the Task Scheduler version for Azure Automation with Managed Identity.
    Authentication is handled automatically via IMDS — no app credentials required.

    Microsoft Graph (User.Read.All)      — enumerate licensed member users
    EXO Admin REST (Exchange.ManageAsApp) — enable archive, auto-expand, set retention

    On completion, triggers Send-ReportNotification in the reporting account
    via cross-account Start-AzAutomationRunbook.

    Automation Variables required (in aa-exo-provisioning):
      TenantId             — Azure AD directory ID
      Organization         — .onmicrosoft.com domain
      SenderEmail          — FROM address for notification email
      RecipientEmail       — TO address for provisioning completion notification
      RetentionPolicyName  — exact EXO retention policy name
      ReportingAccountRG   — resource group of the reporting Automation Account
      ReportingAccountName — name of the reporting Automation Account
      StorageAccountName   — blob storage account name
      StorageContainer     — blob container name

.PARAMETER SkipArchive
    Skip enabling archive mailboxes.

.PARAMETER SkipAutoExpand
    Skip enabling auto-expanding archive.

.PARAMETER SkipRetentionPolicy
    Skip assigning the retention policy.
#>

param(
    [switch]$SkipArchive,
    [switch]$SkipAutoExpand,
    [switch]$SkipRetentionPolicy,
    [bool]$SendAsAttachment = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Helpers ───────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$ts] [$Level] $Message"
}

function Get-ManagedIdentityToken {
    param([string]$Resource)
    (Get-AzAccessToken -ResourceUrl $Resource).Token
}

function Invoke-HttpRequest {
    param([string]$Method, [string]$Uri, [string]$Token, [string]$Body = $null)

    $handler = [System.Net.Http.HttpClientHandler]::new()
    $handler.AutomaticDecompression = [System.Net.DecompressionMethods]::None
    $client  = [System.Net.Http.HttpClient]::new($handler)

    try {
        $client.DefaultRequestHeaders.Add("Authorization", "Bearer $Token")
        $client.DefaultRequestHeaders.Add("Accept", "application/json")

        $request = [System.Net.Http.HttpRequestMessage]::new(
            [System.Net.Http.HttpMethod]::new($Method), $Uri)
        if ($Body) {
            $request.Content = [System.Net.Http.StringContent]::new(
                $Body, [System.Text.Encoding]::UTF8, "application/json")
        }

        $resp = $client.SendAsync($request).Result
        $text = $resp.Content.ReadAsStringAsync().Result

        if (-not $resp.IsSuccessStatusCode) {
            if ($text -match 'already has an archive') { return $text }
            if ($text -match 'Readonly field AutoExpandingArchiveEnabled') { return $text }
            throw "HTTP $([int]$resp.StatusCode) $($resp.ReasonPhrase): $text"
        }
        return $text
    } finally {
        $client.Dispose()
        $handler.Dispose()
    }
}

function Invoke-WithRetry {
    param([scriptblock]$ScriptBlock, [int]$MaxRetries = 3, [int]$DelaySeconds = 5)
    $attempt = 0
    while ($true) {
        try   { return & $ScriptBlock }
        catch {
            $attempt++
            if ($attempt -ge $MaxRetries) { throw }
            Write-Log "Attempt $attempt/$MaxRetries failed: $_. Retrying in ${DelaySeconds}s..." "WARN"
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Get-AllLicensedUsers {
    param([string]$GraphToken)
    $users   = [System.Collections.Generic.List[object]]::new()
    $uri     = "https://graph.microsoft.com/v1.0/users" +
               "?`$filter=assignedLicenses/`$count ne 0 and userType eq 'Member'" +
               "&`$count=true" +
               "&`$select=id,displayName,userPrincipalName" +
               "&`$top=999"
    $headers = @{ Authorization = "Bearer $GraphToken"; ConsistencyLevel = "eventual" }

    $page = 0
    do {
        $page++
        Write-Log "Fetching Graph user page $page ($($users.Count) so far)..."
        $response = Invoke-WithRetry { Invoke-RestMethod -Method GET -Uri $uri -Headers $headers }
        foreach ($u in $response.value) {
            if ($u.userPrincipalName) { $users.Add($u) }
        }
        $uri = if ($response.PSObject.Properties['@odata.nextLink']) { $response.'@odata.nextLink' } else { $null }
    } while ($uri)

    return @($users)
}

# ── Entry point ───────────────────────────────────────────────────────────────

try {
    Write-Log "=== Invoke-ScheduledMailboxProvisioning Started ==="

    if ($SkipArchive -and $SkipAutoExpand -and $SkipRetentionPolicy) {
        throw "All three actions are skipped. Remove at least one -Skip* switch."
    }

    # Authenticate Az cmdlets using the System-Assigned Managed Identity
    Connect-AzAccount -Identity | Out-Null

    # Read configuration from Automation Variables
    $tenantId            = Get-AutomationVariable -Name "TenantId"
    $organization        = Get-AutomationVariable -Name "Organization"
    $retentionPolicyName = Get-AutomationVariable -Name "RetentionPolicyName"
    $senderEmail         = Get-AutomationVariable -Name "SenderEmail"
    $recipientEmail      = Get-AutomationVariable -Name "RecipientEmail"
    $reportingRG         = Get-AutomationVariable -Name "ReportingAccountRG"
    $reportingAccount    = Get-AutomationVariable -Name "ReportingAccountName"
    $storageAcctName     = Get-AutomationVariable -Name "StorageAccountName"
    $storageContainer    = Get-AutomationVariable -Name "StorageContainer"

    Write-Log "Organization : $organization"
    Write-Log "Retention    : $retentionPolicyName"
    if (-not $SkipArchive)         { Write-Log "Action: Enable archive mailbox" }
    if (-not $SkipAutoExpand)      { Write-Log "Action: Enable auto-expanding archive" }
    if (-not $SkipRetentionPolicy) { Write-Log "Action: Assign retention policy '$retentionPolicyName'" }

    # Acquire tokens via Managed Identity IMDS
    Write-Log "Acquiring Graph token (User.Read.All) via Managed Identity..."
    $graphToken = Get-ManagedIdentityToken -Resource "https://graph.microsoft.com/"
    Write-Log "Graph token acquired." "SUCCESS"

    Write-Log "Acquiring EXO admin token (Exchange.ManageAsApp) via Managed Identity..."
    $exoToken        = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
    $tokenAcquiredAt = Get-Date
    Write-Log "EXO admin token acquired." "SUCCESS"

    # Enumerate licensed mailbox users
    Write-Log "Enumerating licensed user mailboxes via Microsoft Graph..."
    $users = Get-AllLicensedUsers -GraphToken $graphToken
    Write-Log "Found $($users.Count) licensed user(s)." "SUCCESS"

    # Process each mailbox
    $results   = [System.Collections.Generic.List[object]]::new()
    $total     = $users.Count
    $idx       = 0

    foreach ($user in $users) {
        $idx++
        $upn = $user.userPrincipalName
        Write-Log "[$idx/$total] Processing $upn..."

        # Refresh tokens every ~40 min (access tokens expire after ~60 min)
        if (((Get-Date) - $tokenAcquiredAt).TotalMinutes -ge 40) {
            Write-Log "Refreshing access tokens at mailbox $idx..."
            $graphToken      = Get-ManagedIdentityToken -Resource "https://graph.microsoft.com/"
            $exoToken        = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
            $tokenAcquiredAt = Get-Date
        }

        $encodedUPN       = [Uri]::EscapeDataString($upn)
        $archiveResult    = "N/A"
        $autoExpandResult = "N/A"
        $retentionResult  = "N/A"

        # 1 ── Enable Archive ──────────────────────────────────────────────────
        if (-not $SkipArchive) {
            try {
                $uri  = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$encodedUPN')/Exchange.UpdateMailboxArchive"
                $resp = Invoke-WithRetry { Invoke-HttpRequest -Method POST -Uri $uri -Token $exoToken -Body '{"archive":true}' }
                $archiveResult = if ($resp -match 'already has an archive') {
                    "Already enabled (no-op)"
                } else {
                    Write-Log "$upn`: archive enabled." "SUCCESS"
                    "Enabled"
                }
            } catch {
                $archiveResult = "Error: $($_.Exception.Message)"
                Write-Log "Failed to enable archive for $upn`: $_" "ERROR"
            }
        }

        # 2 ── Auto-Expanding Archive ──────────────────────────────────────────
        if (-not $SkipAutoExpand) {
            try {
                $uri    = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$encodedUPN')"
                $result = Invoke-WithRetry { Invoke-HttpRequest -Method PATCH -Uri $uri -Token $exoToken -Body '{"AutoExpandingArchiveEnabled":true}' }
                $autoExpandResult = if ($result -match 'Readonly field AutoExpandingArchiveEnabled') {
                    "Already enabled (no-op)"
                } else {
                    Write-Log "$upn`: auto-expanding archive enabled." "SUCCESS"
                    "Enabled"
                }
            } catch {
                $autoExpandResult = "Error: $($_.Exception.Message)"
                Write-Log "Failed to set auto-expand for $upn`: $_" "ERROR"
            }
        }

        # 3 ── Retention Policy ────────────────────────────────────────────────
        if (-not $SkipRetentionPolicy) {
            try {
                $uri           = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$encodedUPN')"
                $retentionJson = "{`"RetentionPolicy`":`"$retentionPolicyName`"}"
                Invoke-WithRetry { Invoke-HttpRequest -Method PATCH -Uri $uri -Token $exoToken -Body $retentionJson | Out-Null }
                $retentionResult = "Set: '$retentionPolicyName'"
                Write-Log "$upn`: retention policy set." "SUCCESS"
            } catch {
                $retentionResult = "Error: $($_.Exception.Message)"
                Write-Log "Failed to set retention policy for $upn`: $_" "ERROR"
            }
        }

        $results.Add([PSCustomObject]@{
            EOID             = $user.id
            DisplayName      = $user.displayName
            UPN              = $upn
            ArchiveResult    = $archiveResult
            AutoExpandResult = $autoExpandResult
            RetentionResult  = $retentionResult
            Timestamp        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        })
    }

    # ── Summary ───────────────────────────────────────────────────────────────
    $applied = @($results | Where-Object {
        $_.ArchiveResult    -match '^Enabled' -or
        $_.AutoExpandResult -match '^Enabled' -or
        $_.RetentionResult  -match '^Set'
    }).Count
    $noops  = @($results | Where-Object { $_.ArchiveResult -match '^Already' }).Count
    $errors = @($results | Where-Object {
        $_.ArchiveResult    -match '^Error' -or
        $_.AutoExpandResult -match '^Error' -or
        $_.RetentionResult  -match '^Error'
    }).Count

    Write-Log "Total: $total  |  Applied: $applied  |  No-op: $noops  |  Errors: $errors"

    # ── Export log CSV and upload to Blob Storage ─────────────────────────────
    $blobName   = "MailboxProvisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $logPath    = Join-Path $env:TEMP $blobName
    $results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSV written: $logPath ($($results.Count) rows)." "SUCCESS"

    Write-Log "Uploading to Blob Storage ($storageAcctName / $storageContainer)..."
    $storageCtx = New-AzStorageContext -StorageAccountName $storageAcctName -UseConnectedAccount
    Set-AzStorageBlobContent -Container $storageContainer -File $logPath -Blob $blobName -Context $storageCtx -Force | Out-Null
    Write-Log "Blob uploaded: $blobName" "SUCCESS"

    # Generate 24-hour User Delegation SAS URL (backed by Entra, not the storage key)
    $sasUrl = New-AzStorageBlobSASToken `
        -Container  $storageContainer `
        -Blob       $blobName `
        -Permission r `
        -ExpiryTime (Get-Date).AddHours(24) `
        -Context    $storageCtx `
        -FullUri
    Write-Log "SAS URL (24h): generated." "SUCCESS"

    # ── Cross-account notification via Send-ReportNotification ────────────────
    Write-Log "Triggering Send-ReportNotification in $reportingAccount..."
    Start-AzAutomationRunbook `
        -ResourceGroupName     $reportingRG `
        -AutomationAccountName $reportingAccount `
        -Name                  "Send-ReportNotification" `
        -Parameters @{
            Subject          = "Mailbox Provisioning Complete — $(Get-Date -Format 'yyyy-MM-dd') — Total: $total  Errors: $errors"
            PortalUrl        = if ($SendAsAttachment) { '' } else { $sasUrl }
            BlobName         = $blobName
            RowCount         = $total
            SendAsAttachment = $SendAsAttachment
        }
    Write-Log "Send-ReportNotification triggered." "SUCCESS"

} catch {
    Write-Log "Script failed: $_" "ERROR"
    throw
} finally {
    Write-Log "=== Invoke-ScheduledMailboxProvisioning Finished ==="
}
