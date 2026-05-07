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

    # Acquire EXO admin token via Managed Identity IMDS (Graph no longer needed)
    Write-Log "Acquiring EXO admin token (Exchange.ManageAsApp) via Managed Identity..."
    $exoToken        = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
    $tokenAcquiredAt = Get-Date
    Write-Log "EXO admin token acquired." "SUCCESS"

    # Fetch current mailbox state via EXO Admin REST (replaces Graph enumeration)
    # Gets ArchiveStatus, AutoExpandingArchiveEnabled, RetentionPolicy in the same bulk call —
    # used for both pre-flight filtering and skipping no-op API calls per mailbox.
    Write-Log "Fetching current mailbox state from EXO Admin REST API..."
    $allMailboxes = [System.Collections.Generic.List[object]]::new()
    $uri = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox" +
           "?`$filter=RecipientTypeDetails eq 'UserMailbox'" +
           "&`$select=ExternalDirectoryObjectId,UserPrincipalName,DisplayName,ArchiveStatus,AutoExpandingArchiveEnabled,RetentionPolicy" +
           "&`$top=1000"
    do {
        $resp = Invoke-WithRetry { Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $exoToken" } -Method GET }
        foreach ($m in $resp.value) { if ($m.UserPrincipalName) { $allMailboxes.Add($m) } }
        $uri = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    } while ($uri)
    Write-Log "Found $($allMailboxes.Count) user mailbox(es)." "SUCCESS"

    # Pre-flight filter: skip mailboxes that already have every required setting
    $users = @($allMailboxes | Where-Object {
        $m = $_
        (-not $SkipArchive         -and $m.ArchiveStatus -ne 'Active') -or
        (-not $SkipAutoExpand      -and -not [bool]$m.AutoExpandingArchiveEnabled) -or
        (-not $SkipRetentionPolicy -and $m.RetentionPolicy -ne $retentionPolicyName)
    })
    Write-Log "Pre-flight filter: $($users.Count) of $($allMailboxes.Count) mailbox(es) need action." "SUCCESS"

    # Process mailboxes in parallel (ThrottleLimit=8)
    # At 15K mailboxes, worst-case ~19 min — well within the 60-min token lifetime,
    # so no token refresh is needed inside the parallel block.
    Write-Log "Processing $($users.Count) mailbox(es) in parallel (ThrottleLimit=8)..."
    $total = $users.Count

    $p_token           = $exoToken
    $p_tenantId        = $tenantId
    $p_retentionPolicy = $retentionPolicyName
    $p_skipArchive     = [bool]$SkipArchive
    $p_skipAutoExpand  = [bool]$SkipAutoExpand
    $p_skipRetention   = [bool]$SkipRetentionPolicy

    $results = [System.Collections.Generic.List[object]]::new()
    $users | ForEach-Object -Parallel {
        $user  = $_
        $upn   = $user.UserPrincipalName
        $token = $using:p_token
        $tid   = $using:p_tenantId
        $retPol        = $using:p_retentionPolicy
        $skipArchive   = $using:p_skipArchive
        $skipAutoExpand = $using:p_skipAutoExpand
        $skipRetention = $using:p_skipRetention

        # Self-contained HTTP helper — no dependency on outer-scope functions.
        # Creates a fresh HttpClient per call (thread-safe); handles EXO no-op responses.
        function Invoke-EXORequest {
            param([string]$Method, [string]$Uri, [string]$Token, [string]$Body = $null)
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                $handler = [System.Net.Http.HttpClientHandler]::new()
                $handler.AutomaticDecompression = [System.Net.DecompressionMethods]::None
                $client  = [System.Net.Http.HttpClient]::new($handler)
                try {
                    $client.DefaultRequestHeaders.Add("Authorization", "Bearer $Token")
                    $client.DefaultRequestHeaders.Add("Accept", "application/json")
                    $req = [System.Net.Http.HttpRequestMessage]::new(
                        [System.Net.Http.HttpMethod]::new($Method), $Uri)
                    if ($Body) {
                        $req.Content = [System.Net.Http.StringContent]::new(
                            $Body, [System.Text.Encoding]::UTF8, "application/json")
                    }
                    $resp = $client.SendAsync($req).Result
                    $text = $resp.Content.ReadAsStringAsync().Result
                    if (-not $resp.IsSuccessStatusCode) {
                        if ($text -match 'already has an archive')                     { return $text }
                        if ($text -match 'Readonly field AutoExpandingArchiveEnabled') { return $text }
                        throw "HTTP $([int]$resp.StatusCode): $text"
                    }
                    return $text
                } catch {
                    if ($attempt -ge 3 -or $_.Exception.Message -notmatch '429|throttl') { throw }
                    Start-Sleep -Seconds ([math]::Pow(2, $attempt))
                } finally {
                    $client.Dispose()
                    $handler.Dispose()
                }
            }
        }

        $encodedUPN       = [Uri]::EscapeDataString($upn)
        $archiveResult    = "N/A"
        $autoExpandResult = "N/A"
        $retentionResult  = "N/A"

        # 1 ── Enable Archive ──────────────────────────────────────────────────
        if (-not $skipArchive) {
            if ($user.ArchiveStatus -eq 'Active') {
                $archiveResult = "Already enabled (no-op)"
            } else {
                try {
                    $uri  = "https://outlook.office365.com/adminapi/beta/$tid/Mailbox('$encodedUPN')/Exchange.UpdateMailboxArchive"
                    $resp = Invoke-EXORequest -Method POST -Uri $uri -Token $token -Body '{"archive":true}'
                    $archiveResult = if ($resp -match 'already has an archive') { "Already enabled (no-op)" } else { "Enabled" }
                } catch {
                    $archiveResult = "Error: $($_.Exception.Message)"
                    Write-Host "[ERROR] Failed to enable archive for $upn`: $_"
                }
            }
        }

        # 2 ── Auto-Expanding Archive ──────────────────────────────────────────
        if (-not $skipAutoExpand) {
            if ([bool]$user.AutoExpandingArchiveEnabled) {
                $autoExpandResult = "Already enabled (no-op)"
            } else {
                try {
                    $uri    = "https://outlook.office365.com/adminapi/beta/$tid/Mailbox('$encodedUPN')"
                    $result = Invoke-EXORequest -Method PATCH -Uri $uri -Token $token -Body '{"AutoExpandingArchiveEnabled":true}'
                    $autoExpandResult = if ($result -match 'Readonly field AutoExpandingArchiveEnabled') { "Already enabled (no-op)" } else { "Enabled" }
                } catch {
                    $autoExpandResult = "Error: $($_.Exception.Message)"
                    Write-Host "[ERROR] Failed to set auto-expand for $upn`: $_"
                }
            }
        }

        # 3 ── Retention Policy ────────────────────────────────────────────────
        if (-not $skipRetention) {
            if ($user.RetentionPolicy -eq $retPol) {
                $retentionResult = "Already set: '$retPol'"
            } else {
                try {
                    $uri = "https://outlook.office365.com/adminapi/beta/$tid/Mailbox('$encodedUPN')"
                    Invoke-EXORequest -Method PATCH -Uri $uri -Token $token -Body "{`"RetentionPolicy`":`"$retPol`"}" | Out-Null
                    $retentionResult = "Set: '$retPol'"
                } catch {
                    $retentionResult = "Error: $($_.Exception.Message)"
                    Write-Host "[ERROR] Failed to set retention policy for $upn`: $_"
                }
            }
        }

        [PSCustomObject]@{
            EOID             = $user.ExternalDirectoryObjectId
            DisplayName      = $user.DisplayName
            UPN              = $upn
            ArchiveResult    = $archiveResult
            AutoExpandResult = $autoExpandResult
            RetentionResult  = $retentionResult
            Timestamp        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
    } -ThrottleLimit 8 | ForEach-Object { $results.Add($_) }

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
