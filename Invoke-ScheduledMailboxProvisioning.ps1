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

    Large-tenant provisioning uses one shared HttpClient with bounded asynchronous
    requests. It does not use ForEach-Object -Parallel because Azure Automation
    runspaces can exhaust the cloud sandbox's memory at scale.

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
    [bool]$SendAsAttachment = $true,
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

function Invoke-ThrottledRequest {
    <#
        Issues POST/PATCH requests concurrently through one shared HttpClient using
        Task.WaitAny. No PowerShell runspaces are created, so memory remains stable
        when processing large tenants in the Azure Automation cloud sandbox.

        Items must have Key, Method, Uri, and Body properties. Results are returned
        in a dictionary keyed by Key with Success, StatusCode, Body, and Error fields.
    #>
    param(
        [Parameter(Mandatory)][object[]]$Items,
        [Parameter(Mandatory)][string]$BearerToken,
        [int]$MaxConcurrency = 8,
        [int]$MaxRetries = 3,
        [string]$ProgressActivity = $null,
        [int]$ProgressEvery = 1000
    )

    $results = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    if ($Items.Count -eq 0) { return $results }

    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromSeconds(100)
    $client.DefaultRequestHeaders.Authorization =
        [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $BearerToken)
    $client.DefaultRequestHeaders.Accept.Add(
        [System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('application/json')
    )
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        $queue = [System.Collections.Generic.Queue[object]]::new()
        foreach ($item in $Items) {
            $queue.Enqueue([PSCustomObject]@{
                Key     = $item.Key
                Method  = $item.Method
                Uri     = $item.Uri
                Body    = $item.Body
                Attempt = 0
            })
        }
        $total = $queue.Count
        $done  = 0

        $inFlight = [System.Collections.Generic.Dictionary[object, object]]::new()
        $startRequest = {
            param([object]$Work)

            $request = [System.Net.Http.HttpRequestMessage]::new(
                [System.Net.Http.HttpMethod]::new($Work.Method),
                $Work.Uri
            )
            if ($Work.Body) {
                $request.Content = [System.Net.Http.StringContent]::new(
                    $Work.Body,
                    [System.Text.Encoding]::UTF8,
                    'application/json'
                )
            }

            $task = $client.SendAsync($request)
            $inFlight[$task] = [PSCustomObject]@{
                Work    = $Work
                Request = $request
                IsDelay = $false
            }
        }
        $startNext = {
            if ($queue.Count -gt 0) {
                & $startRequest $queue.Dequeue()
            }
        }

        for ($i = 0; $i -lt $MaxConcurrency; $i++) { & $startNext }

        while ($inFlight.Count -gt 0) {
            [System.Threading.Tasks.Task[]]$tasks = @($inFlight.Keys)
            $completedIndex = [System.Threading.Tasks.Task]::WaitAny($tasks)
            $completedTask  = $tasks[$completedIndex]
            $state          = $inFlight[$completedTask]
            [void]$inFlight.Remove($completedTask)

            if ($state.IsDelay) {
                & $startRequest $state.Work
                continue
            }

            $statusCode     = -1
            $body           = $null
            $errorMessage   = $null
            $transportError = $false
            $response       = $null
            try {
                $response   = $completedTask.GetAwaiter().GetResult()
                $statusCode = [int]$response.StatusCode
                $body       = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            } catch {
                $errorMessage = if ($_.Exception.InnerException) {
                    $_.Exception.InnerException.Message
                } else {
                    $_.Exception.Message
                }
                $transportError = $true
            } finally {
                if ($response) { $response.Dispose() }
                $state.Request.Dispose()
            }

            $finish = $false
            if (-not $transportError -and $statusCode -ge 200 -and $statusCode -lt 300) {
                $results[$state.Work.Key] = [PSCustomObject]@{
                    Success = $true; StatusCode = $statusCode; Body = $body; Error = $null
                }
                $finish = $true
            } else {
                if (-not $errorMessage) {
                    $errorMessage = "HTTP $statusCode"
                    if ($body) { $errorMessage += ": $body" }
                }

                $retryable = $transportError -or $statusCode -eq 429 -or $statusCode -ge 500
                if ($retryable -and $state.Work.Attempt -lt $MaxRetries) {
                    $state.Work.Attempt++
                    $retryDelay = [math]::Min(
                        ([math]::Pow(2, $state.Work.Attempt) + (Get-Random -Minimum 1 -Maximum 5)),
                        20
                    )
                    $delayTask = [System.Threading.Tasks.Task]::Delay(
                        [TimeSpan]::FromSeconds($retryDelay)
                    )
                    $inFlight[$delayTask] = [PSCustomObject]@{
                        Work = $state.Work; Request = $null; IsDelay = $true
                    }
                } else {
                    $results[$state.Work.Key] = [PSCustomObject]@{
                        Success    = $false
                        StatusCode = $statusCode
                        Body       = $body
                        Error      = "(attempt $($state.Work.Attempt + 1)) $errorMessage"
                    }
                    $finish = $true
                }
            }

            if ($finish) {
                & $startNext
                $done++
                if ($ProgressActivity -and ($done % $ProgressEvery -eq 0 -or $done -eq $total)) {
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Write-Host "[$timestamp] [INFO] $ProgressActivity - $done / $total ($([math]::Round($stopwatch.Elapsed.TotalSeconds))s elapsed)"
                }
            }
        }
    } finally {
        $client.Dispose()
    }

    return $results
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

    # Org-level AutoExpandingArchiveEnabled: when true, mailbox-level is read-only — PATCH
    # returns "Readonly field AutoExpandingArchiveEnabled" for every mailbox. Detect this once
    # to skip the per-mailbox calls entirely and avoid misleading "Already enabled (no-op)" output.
    $orgAutoExpand = $false
    try {
        $orgCfg = Invoke-RestMethod -Uri "https://outlook.office365.com/adminapi/beta/$tenantId/OrganizationConfig" -Headers @{ Authorization = "Bearer $exoToken" } -Method GET
        if ($orgCfg.PSObject.Properties['AutoExpandingArchiveEnabled']) {
            $orgAutoExpand = [bool]$orgCfg.AutoExpandingArchiveEnabled
        }
        if ($orgAutoExpand -and -not $SkipAutoExpand) {
            Write-Log "Org-level AutoExpandingArchiveEnabled is ON — skipping per-mailbox auto-expand (effective for all mailboxes)."
        }
    } catch {
        Write-Log "OrganizationConfig unavailable — will attempt per-mailbox auto-expand check. Detail: $_" "DEBUG"
    }

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
        (-not $SkipAutoExpand      -and -not $orgAutoExpand -and -not [bool]$m.AutoExpandingArchiveEnabled) -or
        (-not $SkipRetentionPolicy -and $m.RetentionPolicy -ne $retentionPolicyName) -or
        (-not $SkipLitigationHold  -and -not [bool]$m.LitigationHoldEnabled)
    })
    Write-Log "Pre-flight filter: $($users.Count) of $($allMailboxes.Count) mailbox(es) need action." "SUCCESS"

    # Process each action as a bounded async phase through one shared HttpClient.
    # This follows Generate-MailboxReport.ps1 and avoids PowerShell runspaces entirely.
    Write-Log "Processing $($users.Count) mailbox(es) asynchronously (concurrency=8)..."
    $total = $users.Count

    $results = [System.Collections.Generic.List[object]]::new()
    $resultByUPN = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $archiveItems    = [System.Collections.Generic.List[object]]::new()
    $autoExpandItems = [System.Collections.Generic.List[object]]::new()
    $retentionItems  = [System.Collections.Generic.List[object]]::new()
    $litigationItems = [System.Collections.Generic.List[object]]::new()

    $retentionBody = @{ RetentionPolicy = $retentionPolicyName } | ConvertTo-Json -Compress
    $litigationBodyObject = [ordered]@{ LitigationHoldEnabled = $true }
    if ($litigationHoldDuration) {
        $litigationBodyObject['LitigationHoldDuration'] = $litigationHoldDuration
    }
    $litigationBody = $litigationBodyObject | ConvertTo-Json -Compress

    foreach ($user in $users) {
        $upn = $user.UserPrincipalName
        $encodedUPN = [Uri]::EscapeDataString($upn)
        $mailboxUri = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$encodedUPN')"

        $row = [PSCustomObject]@{
            EOID                 = $user.ExternalDirectoryObjectId
            DisplayName          = $user.DisplayName
            UPN                  = $upn
            ArchiveResult        = if ($SkipArchive) { "N/A" } elseif ($user.ArchiveStatus -eq 'Active') { "Already enabled (no-op)" } else { "Pending" }
            AutoExpandResult     = if ($SkipAutoExpand) { "N/A" } elseif ($orgAutoExpand -or [bool]$user.AutoExpandingArchiveEnabled) { "Already enabled (no-op)" } else { "Pending" }
            RetentionResult      = if ($SkipRetentionPolicy) { "N/A" } elseif ($user.RetentionPolicy -eq $retentionPolicyName) { "Already set: '$retentionPolicyName'" } else { "Pending" }
            LitigationHoldResult = if ($SkipLitigationHold) { "N/A" } elseif ([bool]$user.LitigationHoldEnabled) { "Already enabled (no-op)" } else { "Pending" }
            Timestamp            = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        [void]$results.Add($row)
        $resultByUPN[$upn] = $row

        if ($row.ArchiveResult -eq "Pending") {
            [void]$archiveItems.Add([PSCustomObject]@{
                Key = $upn; Method = 'POST'
                Uri = "$mailboxUri/Exchange.UpdateMailboxArchive"
                Body = '{"archive":true}'
            })
        }
        if ($row.AutoExpandResult -eq "Pending") {
            [void]$autoExpandItems.Add([PSCustomObject]@{
                Key = $upn; Method = 'PATCH'; Uri = $mailboxUri
                Body = '{"AutoExpandingArchiveEnabled":true}'
            })
        }
        if ($row.RetentionResult -eq "Pending") {
            [void]$retentionItems.Add([PSCustomObject]@{
                Key = $upn; Method = 'PATCH'; Uri = $mailboxUri; Body = $retentionBody
            })
        }
        if ($row.LitigationHoldResult -eq "Pending") {
            [void]$litigationItems.Add([PSCustomObject]@{
                Key = $upn; Method = 'PATCH'; Uri = $mailboxUri; Body = $litigationBody
            })
        }
    }

    if ($archiveItems.Count -gt 0) {
        if (((Get-Date) - $tokenAcquiredAt).TotalMinutes -ge 40) {
            Write-Log "Refreshing EXO token before archive provisioning..."
            $exoToken = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
            $tokenAcquiredAt = Get-Date
        }
        Write-Log "Enabling archive for $($archiveItems.Count) mailbox(es)..."
        $phaseResults = Invoke-ThrottledRequest -Items $archiveItems -BearerToken $exoToken `
            -MaxConcurrency 8 -ProgressActivity "Archive provisioning"
        foreach ($item in $archiveItems) {
            $requestResult = $phaseResults[$item.Key]
            $resultByUPN[$item.Key].ArchiveResult =
                if ($requestResult.Body -match 'already has an archive') {
                    "Already enabled (no-op)"
                } elseif ($requestResult.Success) {
                    "Enabled"
                } else {
                    "Error: $($requestResult.Error)"
                }
        }
    }

    if ($autoExpandItems.Count -gt 0) {
        if (((Get-Date) - $tokenAcquiredAt).TotalMinutes -ge 40) {
            Write-Log "Refreshing EXO token before auto-expanding archive provisioning..."
            $exoToken = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
            $tokenAcquiredAt = Get-Date
        }
        Write-Log "Enabling auto-expanding archive for $($autoExpandItems.Count) mailbox(es)..."
        $phaseResults = Invoke-ThrottledRequest -Items $autoExpandItems -BearerToken $exoToken `
            -MaxConcurrency 8 -ProgressActivity "Auto-expanding archive provisioning"
        foreach ($item in $autoExpandItems) {
            $requestResult = $phaseResults[$item.Key]
            $resultByUPN[$item.Key].AutoExpandResult =
                if ($requestResult.Body -match 'Readonly field AutoExpandingArchiveEnabled') {
                    "Already enabled (no-op)"
                } elseif ($requestResult.Success) {
                    "Enabled"
                } else {
                    "Error: $($requestResult.Error)"
                }
        }
    }

    if ($retentionItems.Count -gt 0) {
        if (((Get-Date) - $tokenAcquiredAt).TotalMinutes -ge 40) {
            Write-Log "Refreshing EXO token before retention policy provisioning..."
            $exoToken = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
            $tokenAcquiredAt = Get-Date
        }
        Write-Log "Assigning retention policy to $($retentionItems.Count) mailbox(es)..."
        $phaseResults = Invoke-ThrottledRequest -Items $retentionItems -BearerToken $exoToken `
            -MaxConcurrency 8 -ProgressActivity "Retention policy provisioning"
        foreach ($item in $retentionItems) {
            $requestResult = $phaseResults[$item.Key]
            $resultByUPN[$item.Key].RetentionResult =
                if ($requestResult.Success) {
                    "Set: '$retentionPolicyName'"
                } else {
                    "Error: $($requestResult.Error)"
                }
        }
    }

    if ($litigationItems.Count -gt 0) {
        if (((Get-Date) - $tokenAcquiredAt).TotalMinutes -ge 40) {
            Write-Log "Refreshing EXO token before Litigation Hold provisioning..."
            $exoToken = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
            $tokenAcquiredAt = Get-Date
        }
        Write-Log "Enabling Litigation Hold for $($litigationItems.Count) mailbox(es)..."
        $phaseResults = Invoke-ThrottledRequest -Items $litigationItems -BearerToken $exoToken `
            -MaxConcurrency 8 -ProgressActivity "Litigation Hold provisioning"
        $holdSuffix = if ($litigationHoldDuration) { " (duration: $litigationHoldDuration)" } else { "" }
        foreach ($item in $litigationItems) {
            $requestResult = $phaseResults[$item.Key]
            $resultByUPN[$item.Key].LitigationHoldResult =
                if ($requestResult.Success) {
                    "Enabled$holdSuffix"
                } else {
                    "Error: $($requestResult.Error)"
                }
        }
    }

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
    $failureDetails = @(
        "Message=$($_.Exception.Message)"
        "ExceptionType=$($_.Exception.GetType().FullName)"
        "FullyQualifiedErrorId=$($_.FullyQualifiedErrorId)"
        "Category=$($_.CategoryInfo.Category)"
        "Target=$($_.CategoryInfo.TargetName)"
        "Position=$($_.InvocationInfo.PositionMessage -replace '\r?\n', ' ')"
        "Stack=$($_.ScriptStackTrace -replace '\r?\n', ' <- ')"
    ) -join " | "
    Write-Log "Script failed: $failureDetails" "ERROR"
    throw
} finally {
    Write-Log "=== Invoke-ScheduledMailboxProvisioning Finished ==="
}
