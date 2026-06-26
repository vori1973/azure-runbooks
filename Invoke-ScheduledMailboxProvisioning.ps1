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
      TenantId                  — Azure AD directory ID
      Organization              — .onmicrosoft.com domain
      SenderEmail               — FROM address for notification email
      RecipientEmail            — TO address for provisioning completion notification
      RetentionPolicyName       — exact EXO retention policy name
      ReportingAccountRG        — resource group of the reporting Automation Account
      ReportingAccountName      — name of the reporting Automation Account
      StorageAccountName        — blob storage account name
      StorageContainer          — blob container name
      LitigationHoldDuration    — (optional) hold duration in days (e.g. 2555) or "Unlimited".
                                   If the variable does not exist, duration defaults to Unlimited.
      ProvisioningCreatedAfter  — (optional) ISO 8601 date string (e.g. "2025-01-01").
                                   When set, only mailboxes created on or after this date are
                                   processed. Filtered client-side after fetch (EXO Admin REST
                                   exposes WhenMailboxCreated as Edm.String so OData datetime
                                   operators are not supported). If absent, all mailboxes are
                                   processed.

.PARAMETER SkipArchive
    Skip enabling archive mailboxes.

.PARAMETER SkipAutoExpand
    Skip enabling auto-expanding archive.

.PARAMETER SkipRetentionPolicy
    Skip assigning the retention policy.

.PARAMETER SkipLitigationHold
    Skip enabling Litigation Hold.

.PARAMETER DebugLogs
    When $true, emit verbose DEBUG-level lines (token acquisition, action announcements, SAS generation, notification trigger).
    Default $false — only INFO / WARN / ERROR / SUCCESS lines are written, reducing job output volume.
#>

param(
    [switch]$SkipArchive,
    [switch]$SkipAutoExpand,
    [switch]$SkipRetentionPolicy,
    [switch]$SkipLitigationHold,
    [bool]$SendAsAttachment = $false,
    [bool]$DebugLogs = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Helpers ───────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if ($Level -eq "DEBUG" -and -not $script:DebugLogs) { return }
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "[$ts] [$Level] $Message"
}

function Get-ManagedIdentityToken {
    param([string]$Resource)
    (Get-AzAccessToken -ResourceUrl $Resource -AsSecureString).Token | ConvertFrom-SecureString -AsPlainText
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

    if ($SkipArchive -and $SkipAutoExpand -and $SkipRetentionPolicy -and $SkipLitigationHold) {
        throw "All actions are skipped. Remove at least one -Skip* switch."
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

    # Read optional LitigationHoldDuration variable (missing variable = Unlimited)
    $litigationHoldDuration = $null
    try   { $litigationHoldDuration = Get-AutomationVariable -Name "LitigationHoldDuration" }
    catch { Write-Log "LitigationHoldDuration variable not found — defaulting to Unlimited." }

    # Read optional ProvisioningCreatedAfter variable (missing = process all mailboxes)
    $createdAfter = $null
    try {
        $val = Get-AutomationVariable -Name "ProvisioningCreatedAfter"
        if ($val) {
            $createdAfter = [datetime]::Parse($val, [System.Globalization.CultureInfo]::InvariantCulture,
                                              [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()
        }
    } catch { Write-Log "ProvisioningCreatedAfter variable not found — processing all mailboxes." }

    Write-Log "Organization : $organization"
    Write-Log "Retention    : $retentionPolicyName"
    if ($createdAfter) { Write-Log "Date filter  : mailboxes created on or after $($createdAfter.ToString('yyyy-MM-dd')) UTC" }
    if (-not $SkipArchive)         { Write-Log "Action: Enable archive mailbox" "DEBUG" }
    if (-not $SkipAutoExpand)      { Write-Log "Action: Enable auto-expanding archive" "DEBUG" }
    if (-not $SkipRetentionPolicy) { Write-Log "Action: Assign retention policy '$retentionPolicyName'" "DEBUG" }
    if (-not $SkipLitigationHold) {
        $durationLabel = if ($litigationHoldDuration) { $litigationHoldDuration } else { "Unlimited" }
        Write-Log "Action: Enable Litigation Hold (duration: $durationLabel)" "DEBUG"
    }

    # ── Pre-flight: validate Blob Storage permissions before the long provisioning loop ──
    Write-Log "Pre-flight: validating Blob Storage permissions (Storage Blob Data Contributor + Storage Blob Delegator)..."
    $storageCtx = New-AzStorageContext -StorageAccountName $storageAcctName -UseConnectedAccount
    $checkBlob  = ".perm-check-$(New-Guid)"
    $checkFile  = Join-Path $env:TEMP $checkBlob
    try {
        [System.IO.File]::WriteAllText($checkFile, "perm-check")
        Set-AzStorageBlobContent -Container $storageContainer -File $checkFile -Blob $checkBlob -Context $storageCtx -Force | Out-Null
        New-AzStorageBlobSASToken -Container $storageContainer -Blob $checkBlob -Permission r `
            -ExpiryTime (Get-Date).AddMinutes(5) -Context $storageCtx -FullUri | Out-Null
        Remove-AzStorageBlob -Container $storageContainer -Blob $checkBlob -Context $storageCtx -Force -ErrorAction SilentlyContinue
        Write-Log "Pre-flight passed — Storage Blob Data Contributor (upload) and Storage Blob Delegator (SAS) both confirmed on $storageAcctName/$storageContainer." "DEBUG"
    } catch {
        $permMsg = "Pre-flight check failed — Blob Storage 403. " +
                   "Grant the provisioning Managed Identity 'Storage Blob Data Contributor' on the container " +
                   "and 'Storage Blob Delegator' on the storage account. " +
                   "See README Troubleshooting section. Detail: $_"
        Write-Log $permMsg "ERROR"
        throw $permMsg
    } finally {
        Remove-Item $checkFile -Force -ErrorAction SilentlyContinue
    }

    # Acquire EXO admin token via Managed Identity IMDS (Graph no longer needed)
    Write-Log "Acquiring EXO admin token (Exchange.ManageAsApp) via Managed Identity..." "DEBUG"
    $exoToken        = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
    $tokenAcquiredAt = Get-Date
    Write-Log "EXO admin token acquired." "DEBUG"

    # Fetch current mailbox state via EXO Admin REST (replaces Graph enumeration)
    # Gets ArchiveStatus, AutoExpandingArchiveEnabled, RetentionPolicy in the same bulk call —
    # used for both pre-flight filtering and skipping no-op API calls per mailbox.
    # WhenMailboxCreated is Edm.String in the EXO Admin REST schema — OData datetime
    # comparison operators are unsupported on it, so date filtering is done client-side below.
    Write-Log "Fetching all user mailboxes from EXO Admin REST API..."
    $allMailboxes = [System.Collections.Generic.List[object]]::new()
    $uri = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox" +
           "?`$filter=RecipientTypeDetails eq 'UserMailbox'" +
           "&`$select=ExternalDirectoryObjectId,UserPrincipalName,DisplayName,ArchiveStatus,AutoExpandingArchiveEnabled,RetentionPolicy,LitigationHoldEnabled,WhenMailboxCreated" +
           "&`$top=1000"
    do {
        $resp = Invoke-WithRetry { Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $exoToken" } -Method GET }
        foreach ($m in $resp.value) { if ($m.UserPrincipalName) { $allMailboxes.Add($m) } }
        $uri = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    } while ($uri)
    Write-Log "API returned $($allMailboxes.Count) user mailbox(es)." "SUCCESS"

    # Client-side safety net: re-apply the date filter in case the server ignored it
    if ($createdAfter) {
        $beforeCount  = $allMailboxes.Count
        $allMailboxes = [System.Collections.Generic.List[object]]($allMailboxes | Where-Object {
            $_.WhenMailboxCreated -and [datetime]$_.WhenMailboxCreated -ge $createdAfter
        })
        Write-Log "Date filter (client-side): $($allMailboxes.Count) of $beforeCount mailbox(es) match." "SUCCESS"
    }

    # Pre-flight filter: skip mailboxes that already have every required setting
    $users = @($allMailboxes | Where-Object {
        $m = $_
        (-not $SkipArchive         -and $m.ArchiveStatus -ne 'Active') -or
        (-not $SkipAutoExpand      -and -not [bool]$m.AutoExpandingArchiveEnabled) -or
        (-not $SkipRetentionPolicy -and $m.RetentionPolicy -ne $retentionPolicyName) -or
        (-not $SkipLitigationHold  -and -not [bool]$m.LitigationHoldEnabled)
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
    $p_skipLitHold     = [bool]$SkipLitigationHold
    $p_litHoldDuration = $litigationHoldDuration

    $results = [System.Collections.Generic.List[object]]::new()
    $users | ForEach-Object -Parallel {
        $user  = $_
        $upn   = $user.UserPrincipalName
        $token = $using:p_token
        $tid   = $using:p_tenantId
        $retPol          = $using:p_retentionPolicy
        $skipArchive     = $using:p_skipArchive
        $skipAutoExpand  = $using:p_skipAutoExpand
        $skipRetention   = $using:p_skipRetention
        $skipLitHold     = $using:p_skipLitHold
        $litHoldDuration = $using:p_litHoldDuration

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
        $litigationResult = "N/A"

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

        # 4 ── Litigation Hold ─────────────────────────────────────────────────
        if (-not $skipLitHold) {
            if ([bool]$user.LitigationHoldEnabled) {
                $litigationResult = "Already enabled (no-op)"
            } else {
                try {
                    $bodyObj = [ordered]@{ LitigationHoldEnabled = $true }
                    if ($litHoldDuration) { $bodyObj['LitigationHoldDuration'] = $litHoldDuration }
                    $bodyJson = $bodyObj | ConvertTo-Json -Compress
                    $uri = "https://outlook.office365.com/adminapi/beta/$tid/Mailbox('$encodedUPN')"
                    Invoke-EXORequest -Method PATCH -Uri $uri -Token $token -Body $bodyJson | Out-Null
                    $holdSuffix   = if ($litHoldDuration) { " (duration: $litHoldDuration)" } else { "" }
                    $litigationResult = "Enabled$holdSuffix"
                } catch {
                    $litigationResult = "Error: $($_.Exception.Message)"
                    Write-Host "[ERROR] Failed to enable Litigation Hold for $upn`: $_"
                }
            }
        }

        [PSCustomObject]@{
            EOID                 = $user.ExternalDirectoryObjectId
            DisplayName          = $user.DisplayName
            UPN                  = $upn
            ArchiveResult        = $archiveResult
            AutoExpandResult     = $autoExpandResult
            RetentionResult      = $retentionResult
            LitigationHoldResult = $litigationResult
            Timestamp            = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
    } -ThrottleLimit 8 | ForEach-Object { $results.Add($_) }

    # ── Summary ───────────────────────────────────────────────────────────────
    $applied = @($results | Where-Object {
        $_.ArchiveResult        -match '^Enabled' -or
        $_.AutoExpandResult     -match '^Enabled' -or
        $_.RetentionResult      -match '^Set' -or
        $_.LitigationHoldResult -match '^Enabled'
    }).Count
    $noops  = @($results | Where-Object {
        $_.ArchiveResult        -match '^Already' -or
        $_.AutoExpandResult     -match '^Already' -or
        $_.LitigationHoldResult -match '^Already'
    }).Count
    $errors = @($results | Where-Object {
        $_.ArchiveResult        -match '^Error' -or
        $_.AutoExpandResult     -match '^Error' -or
        $_.RetentionResult      -match '^Error' -or
        $_.LitigationHoldResult -match '^Error'
    }).Count

    Write-Log "Total: $total  |  Applied: $applied  |  No-op: $noops  |  Errors: $errors"

    # ── Export log CSV and upload to Blob Storage ─────────────────────────────
    $blobName   = "MailboxProvisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $logPath    = Join-Path $env:TEMP $blobName
    $results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSV written: $logPath ($($results.Count) rows)." "DEBUG"

    Write-Log "Uploading to Blob Storage ($storageAcctName / $storageContainer)..."
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
    Write-Log "SAS URL (24h): generated." "DEBUG"

    # ── Cross-account notification via Send-ReportNotification ────────────────
    Write-Log "Triggering Send-ReportNotification in $reportingAccount..." "DEBUG"
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
        } | Out-Null
    Write-Log "Send-ReportNotification triggered." "DEBUG"

} catch {
    Write-Log "Script failed: $_" "ERROR"
    throw
} finally {
    Write-Log "=== Invoke-ScheduledMailboxProvisioning Finished ==="
}
