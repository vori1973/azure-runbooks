#Requires -Version 7.4
<#
.SYNOPSIS
    Azure Automation runbook: collect mailbox usage data, upload CSV to Blob Storage,
    and trigger Send-ReportNotification with a 24-hour SAS download link.

.DESCRIPTION
    Uses the Reports API hybrid mode (fast, ~2 min for large tenants):
      Phase 1 — EXO Admin REST API for mailbox metadata (no module required)
      Phase 2 — Graph getMailboxUsageDetail report for bulk usage stats (24-48h stale)
      Phase 3 — Join metadata + usage by UPN
      Phase 4 — Gap fill (async, concurrency=8): live EXO Admin REST for mailboxes missing from report
      Phase 5 — Archive stats for report-matched mailboxes with active archives (async, concurrency=8)

    Phases 4 and 5 issue many independent EXO REST calls. They use a single shared
    HttpClient with bounded concurrency (Invoke-ThrottledGet, below) rather than
    ForEach-Object -Parallel — the latter allocates a full PowerShell runspace per
    concurrent item, which is fine at a few hundred items but causes
    System.OutOfMemoryException in the Azure Automation cloud sandbox once mailbox
    counts reach into the thousands. Invoke-ThrottledGet runs entirely on the single
    calling runspace and only overlaps network I/O, so memory stays flat regardless
    of how many mailboxes are processed.

    Exports final CSV, uploads to Blob Storage, generates 24-hour SAS URL,
    and calls Send-ReportNotification via Start-AzAutomationRunbook.

    No ExchangeOnlineManagement module required. All EXO operations use the
    EXO Admin REST API directly with a Managed Identity token.

    Permissions required on the reporting Managed Identity:
      Graph  Reports.Read.All      — Reports API call
      Graph  Mail.Send             — delegated to Send-ReportNotification
      EXO    Exchange.ManageAsApp  — EXO Admin REST API
      Entra  Exchange Recipient Administrator directory role
      Azure  Storage Blob Data Contributor on the reports container

    Automation Variables required (in aa-exo-reporting):
      TenantId           — Azure AD directory ID
      Organization       — .onmicrosoft.com domain
      StorageAccountName — blob storage account name
      StorageContainer   — blob container name

.PARAMETER ReportsPeriod
    Graph Reports API period. Default D7. Options: D7, D30, D90, D180.

.PARAMETER DebugLogs
    When $true, emit verbose DEBUG-level lines (token acquisition, SAS generation, notification trigger).
    Default $false — only INFO / WARN / ERROR / SUCCESS lines are written, reducing job output volume.
#>

param(
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$ReportsPeriod = 'D7',
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

function ConvertTo-MB {
    param([object]$Size)
    if ($null -eq $Size) { return 0 }
    # Numeric bytes (REST API returns integer)
    if ($Size -is [long] -or $Size -is [int] -or $Size -is [double]) {
        return [math]::Round([double]$Size / 1MB, 2)
    }
    # String format "X GB (N bytes)" — EXO quota format
    if ($Size.ToString() -match '\(([0-9,]+) bytes\)') {
        return [math]::Round([long]($Matches[1] -replace ',', '') / 1MB, 2)
    }
    # Pure numeric string
    $parsed = 0L
    if ([long]::TryParse($Size.ToString(), [ref]$parsed)) {
        return [math]::Round($parsed / 1MB, 2)
    }
    return 0
}

function Get-UsagePercent {
    param([double]$UsedMB, [double]$QuotaMB)
    if ($QuotaMB -le 0) { return 0 }
    return [math]::Round(($UsedMB / $QuotaMB) * 100, 2)
}

function Get-JsonProp {
    # Null-safe property read: a 2xx response with an empty body parses to $null via
    # ConvertFrom-Json, and Set-StrictMode throws on ANY property access against $null
    # (including .PSObject.Properties itself) — this centralizes the null/missing check.
    param([object]$Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    if ($Obj.PSObject.Properties[$Name]) { return $Obj.$Name }
    return $null
}

function ConvertFrom-StatsJson {
    # Parses an EXO Admin REST response body and unwraps the OData 'value' array if
    # present. Returns $null for an empty/whitespace body instead of throwing.
    param([string]$Json)
    if ([string]::IsNullOrWhiteSpace($Json)) { return $null }
    $parsed = ConvertFrom-Json $Json
    $item = Get-JsonProp $parsed 'value'
    if ($item) { return ($item | Select-Object -First 1) }
    return $parsed
}

function Invoke-WithRetry {
    param([scriptblock]$ScriptBlock, [int]$MaxRetries = 5)
    $attempt = 0
    while ($true) {
        try { return & $ScriptBlock }
        catch {
            $msg = $_.Exception.Message
            $attempt++
            if ($attempt -ge $MaxRetries) { throw }
            if ($msg -match '429|throttl|too many|exceeded|transient') {
                $wait = [math]::Pow(2, $attempt) + (Get-Random -Minimum 1 -Maximum 5)
                Write-Log "Transient error — waiting $wait seconds before retry $attempt/$MaxRetries..." "WARN"
                Start-Sleep -Seconds $wait
            } else { throw }
        }
    }
}

function Invoke-ThrottledGet {
    <#
        Issues many GET requests concurrently through a single shared HttpClient,
        bounded to MaxConcurrency requests in flight at a time, using .NET's async
        Task APIs directly (Task.WaitAny over the in-flight set). Everything runs on
        the calling runspace — no extra PowerShell runspaces are created, so memory
        stays flat no matter how many items are queued. This replaces
        ForEach-Object -Parallel, which allocates one runspace per concurrent item
        and exhausts the Azure Automation sandbox's memory ceiling at scale.

        -Items must be objects with .Key (a unique string used to look up the result)
        and .Uri properties. Returns a Dictionary<string,object> keyed by .Key, where
        each value is @{ Success; StatusCode; Body; Error }.
    #>
    param(
        [Parameter(Mandatory)][object[]]$Items,
        [Parameter(Mandatory)][string]$BearerToken,
        [int]$MaxConcurrency = 8,
        [int]$MaxRetries = 3,
        [string]$ProgressActivity = $null,
        [int]$ProgressEvery = 1000
    )

    $results = [System.Collections.Generic.Dictionary[string, object]]::new()
    if ($Items.Count -eq 0) { return $results }

    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromSeconds(100)
    $client.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $BearerToken)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        $queue = [System.Collections.Generic.Queue[object]]::new()
        foreach ($it in $Items) { $queue.Enqueue([PSCustomObject]@{ Key = $it.Key; Uri = $it.Uri; Attempt = 0 }) }
        $total = $queue.Count
        $done  = 0

        $inFlight  = [System.Collections.Generic.Dictionary[object, object]]::new()
        $startNext = {
            if ($queue.Count -gt 0) {
                $work = $queue.Dequeue()
                $inFlight[$client.GetAsync($work.Uri)] = $work
            }
        }
        for ($i = 0; $i -lt $MaxConcurrency; $i++) { & $startNext }

        while ($inFlight.Count -gt 0) {
            [System.Threading.Tasks.Task[]]$tasksArr = @($inFlight.Keys)
            $idx      = [System.Threading.Tasks.Task]::WaitAny($tasksArr)
            $doneTask = $tasksArr[$idx]
            $work     = $inFlight[$doneTask]
            [void]$inFlight.Remove($doneTask)

            $statusCode     = -1
            $body           = $null
            $errMsg         = $null
            $transportError = $false
            try {
                $resp       = $doneTask.GetAwaiter().GetResult()
                $statusCode = [int]$resp.StatusCode
                $body       = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                $resp.Dispose()
            } catch {
                $errMsg         = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
                $transportError = $true
            }

            $finish = $false
            if (-not $transportError -and $statusCode -ge 200 -and $statusCode -lt 300) {
                $results[$work.Key] = [PSCustomObject]@{ Success = $true; StatusCode = $statusCode; Body = $body; Error = $null }
                $finish = $true
            } elseif (-not $transportError -and $statusCode -eq 404) {
                $results[$work.Key] = [PSCustomObject]@{ Success = $false; StatusCode = 404; Body = $null; Error = 'NotFound' }
                $finish = $true
            } else {
                if (-not $errMsg) { $errMsg = "HTTP $statusCode" }
                $retryable = $transportError -or $statusCode -eq 429 -or $statusCode -ge 500
                if ($retryable -and $work.Attempt -lt $MaxRetries) {
                    $work.Attempt++
                    Start-Sleep -Seconds ([math]::Min(([math]::Pow(2, $work.Attempt) + (Get-Random -Minimum 1 -Maximum 5)), 20))
                    $inFlight[$client.GetAsync($work.Uri)] = $work
                } else {
                    $results[$work.Key] = [PSCustomObject]@{ Success = $false; StatusCode = $statusCode; Body = $null; Error = "(attempt $($work.Attempt + 1)) $errMsg" }
                    $finish = $true
                }
            }

            if ($finish) {
                & $startNext
                $done++
                if ($ProgressActivity -and ($done % $ProgressEvery -eq 0 -or $done -eq $total)) {
                    $tsNow = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Write-Host "[$tsNow] [INFO] $ProgressActivity - $done / $total ($([math]::Round($sw.Elapsed.TotalSeconds))s elapsed)"
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
    Write-Log "=== Generate-MailboxReport Started ==="

    # Authenticate Az cmdlets using the System-Assigned Managed Identity
    Connect-AzAccount -Identity | Out-Null

    # Read configuration from Automation Variables
    $tenantId         = Get-AutomationVariable -Name "TenantId"
    $organization     = Get-AutomationVariable -Name "Organization"
    $storageAcctName  = Get-AutomationVariable -Name "StorageAccountName"
    $storageContainer = Get-AutomationVariable -Name "StorageContainer"
    $repAccountRG     = Get-AutomationVariable -Name "AutomationAccountRG"
    $repAccountName   = Get-AutomationVariable -Name "AutomationAccountName"

    Write-Log "Organization     : $organization"
    Write-Log "Storage account  : $storageAcctName / $storageContainer"
    Write-Log "Reports period   : $ReportsPeriod"

    # Initialise storage context (OAuth — Managed Identity)
    $storageCtx = New-AzStorageContext -StorageAccountName $storageAcctName -UseConnectedAccount
    Write-Log "Storage context initialised for $storageAcctName." "DEBUG"

    $exportPath = Join-Path $env:TEMP "MailboxOverview_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    # Acquire tokens via Managed Identity
    Write-Log "Acquiring Graph token (Reports.Read.All) via Managed Identity..." "DEBUG"
    $graphToken = Get-ManagedIdentityToken -Resource "https://graph.microsoft.com/"
    Write-Log "Graph token acquired." "DEBUG"

    Write-Log "Acquiring EXO admin token (Exchange.ManageAsApp) via Managed Identity..." "DEBUG"
    $exoToken        = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
    $exoHeaders      = @{ Authorization = "Bearer $exoToken" }
    $tokenAcquiredAt = Get-Date
    Write-Log "EXO admin token acquired." "DEBUG"

    # Org-level AutoExpandingArchiveEnabled: when enabled at org level, the mailbox-level
    # property is read-only and always returns false regardless of effective state.
    # Fetch once here so Phase 1 can compute the correct effective value.
    $orgAutoExpand = $false
    try {
        $orgCfg = Invoke-RestMethod -Uri "https://outlook.office365.com/adminapi/beta/$tenantId/OrganizationConfig" -Headers $exoHeaders -Method GET
        if ($orgCfg.PSObject.Properties['AutoExpandingArchiveEnabled']) {
            $orgAutoExpand = [bool]$orgCfg.AutoExpandingArchiveEnabled
        }
        if ($orgAutoExpand) { Write-Log "Org-level AutoExpandingArchiveEnabled is ON — effective value will be True for all mailboxes." }
    } catch {
        Write-Log "OrganizationConfig unavailable — AutoExpandingArchiveEnabled reflects mailbox-level value only. Detail: $_" "DEBUG"
    }

    # ── Phase 1: EXO metadata via Admin REST API ───────────────────────────────
    Write-Log "Retrieving mailbox metadata via EXO Admin REST API..."
    $rawMailboxes = [System.Collections.Generic.List[object]]::new()
    $uri = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox" +
           "?`$filter=RecipientTypeDetails eq 'UserMailbox'" +
           "&`$select=ExternalDirectoryObjectId,UserPrincipalName,DisplayName," +
           "ProhibitSendQuota,ArchiveStatus,ArchiveQuota,AutoExpandingArchiveEnabled," +
           "RetentionPolicy,RetentionHoldEnabled,LitigationHoldEnabled,LitigationHoldDuration,ArchiveGuid," +
           "WhenMailboxCreated" +
           "&`$top=1000"
    do {
        $resp = Invoke-WithRetry { Invoke-RestMethod -Uri $uri -Headers $exoHeaders -Method GET }
        foreach ($m in $resp.value) { $rawMailboxes.Add($m) }
        $uri = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    } while ($uri)
    Write-Log "Found $($rawMailboxes.Count) mailboxes." "SUCCESS"

    $allMailboxes = @($rawMailboxes | ForEach-Object {
        $qmb  = if ($_.ProhibitSendQuota -like 'Unlimited') { 0 } else { ConvertTo-MB $_.ProhibitSendQuota }
        $aqmb = if ($_.PSObject.Properties['ArchiveQuota'] -and $_.ArchiveQuota -and $_.ArchiveQuota -notlike '*Unlimited*') { ConvertTo-MB $_.ArchiveQuota } else { 0 }
        [PSCustomObject]@{
            EOID                        = $_.ExternalDirectoryObjectId
            UserPrincipalName           = $_.UserPrincipalName
            DisplayName                 = $_.DisplayName
            QuotaMBComputed             = $qmb
            ArchiveQuotaMBComputed      = $aqmb
            ArchiveStatus               = $_.ArchiveStatus
            AutoExpandingArchiveEnabled = $orgAutoExpand -or ($_.PSObject.Properties['AutoExpandingArchiveEnabled'] -and [bool]$_.AutoExpandingArchiveEnabled)
            RetentionPolicy             = if ($_.PSObject.Properties['RetentionPolicy']) { $_.RetentionPolicy } else { $null }
            RetentionHoldEnabled        = if ($_.PSObject.Properties['RetentionHoldEnabled']) { [bool]$_.RetentionHoldEnabled } else { $false }
            LitigationHoldEnabled       = if ($_.PSObject.Properties['LitigationHoldEnabled']) { [bool]$_.LitigationHoldEnabled } else { $false }
            LitigationHoldDuration      = if ($_.PSObject.Properties['LitigationHoldDuration'] -and $_.LitigationHoldDuration) {
                                              $raw = $_.LitigationHoldDuration.ToString()
                                              if ($raw -match '^(\d+)\.') { [int]$Matches[1] } else { $raw }
                                          } else { $null }
            ArchiveGuid                 = if ($_.PSObject.Properties['ArchiveGuid'] -and $_.ArchiveGuid) { $_.ArchiveGuid.ToString() } else { $null }
            WhenMailboxCreated          = if ($_.PSObject.Properties['WhenMailboxCreated'] -and $_.WhenMailboxCreated) { $_.WhenMailboxCreated.ToString() } else { $null }
        }
    })

    $archiveGuidByEOID = @{}
    foreach ($mbx in $allMailboxes) {
        if ($mbx.ArchiveGuid) { $archiveGuidByEOID[$mbx.EOID] = $mbx.ArchiveGuid }
    }

    # ── Phase 2: Reports API usage stats ──────────────────────────────────────
    Write-Log "Fetching mailbox usage report via Graph Reports API (period=$ReportsPeriod, data 24-48h stale)..." "WARN"
    $tempReport = [System.IO.Path]::GetTempFileName()
    try {
        $reportBytes = Invoke-RestMethod `
            -Uri     "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$ReportsPeriod')" `
            -Method  GET `
            -Headers @{ Authorization = "Bearer $graphToken" }
        [System.IO.File]::WriteAllText($tempReport, $reportBytes, [System.Text.Encoding]::UTF8)
        $reportRows = Import-Csv -Path $tempReport -Encoding UTF8
    } finally {
        Remove-Item $tempReport -Force -ErrorAction SilentlyContinue
    }
    Write-Log "Report downloaded — $($reportRows.Count) rows." "SUCCESS"

    if ($reportRows.Count -gt 0 -and -not ($reportRows[0].PSObject.Properties['User Principal Name'])) {
        Write-Log "UPN column missing — tenant privacy settings may be hiding user details (M365 Admin > Org Settings > Reports)." "WARN"
    }

    $reportLookup = @{}
    foreach ($row in $reportRows) {
        if ($row.'Is Deleted' -eq 'True') { continue }
        $upn = $row.'User Principal Name'
        if ($upn) { $reportLookup[$upn.ToLower()] = $row }
    }

    # ── Phase 3: Join EXO metadata + Reports usage ────────────────────────────
    $ts      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $results = [System.Collections.Generic.List[object]]::new()
    $gaps    = [System.Collections.Generic.List[object]]::new()

    foreach ($mbx in $allMailboxes) {
        $reportRow = $reportLookup[$mbx.UserPrincipalName.ToLower()]
        if ($reportRow) {
            $usedMB  = [math]::Round([long]$reportRow.'Storage Used (Byte)' / 1MB, 2)
            $quotaMB = if ($mbx.QuotaMBComputed -gt 0) { $mbx.QuotaMBComputed } else {
                if ([long]$reportRow.'Prohibit Send Quota (Byte)' -gt 0) {
                    [math]::Round([long]$reportRow.'Prohibit Send Quota (Byte)' / 1MB, 2)
                } else { 0 }
            }
            $pct = Get-UsagePercent -UsedMB $usedMB -QuotaMB $quotaMB
            $results.Add([PSCustomObject]@{
                EOID                        = $mbx.EOID
                DisplayName                 = $mbx.DisplayName
                UPN                         = $mbx.UserPrincipalName
                UsedMB                      = $usedMB
                UsedGB                      = [math]::Round($usedMB  / 1024, 3)
                QuotaMB                     = $quotaMB
                QuotaGB                     = [math]::Round($quotaMB / 1024, 3)
                UsagePercent                = $pct
                ItemCount                   = if ($reportRow.'Item Count') { [long]$reportRow.'Item Count' } else { 0 }
                DeletedItemCount            = if ($reportRow.'Deleted Item Count') { [long]$reportRow.'Deleted Item Count' } else { 0 }
                DeletedItemSizeMB           = if ($reportRow.'Deleted Item Size (Byte)') { [math]::Round([long]$reportRow.'Deleted Item Size (Byte)' / 1MB, 2) } else { 0 }
                LastActivityDate            = $reportRow.'Last Activity Date'
                ArchiveStatus               = $mbx.ArchiveStatus
                ArchiveEnabled              = ($mbx.ArchiveStatus -eq 'Active')
                AutoExpandingArchiveEnabled = $mbx.AutoExpandingArchiveEnabled
                ArchiveUsedMB               = $null
                ArchiveUsedGB               = $null
                ArchiveQuotaMB              = $mbx.ArchiveQuotaMBComputed
                ArchiveQuotaGB              = if ($mbx.ArchiveQuotaMBComputed -gt 0) { [math]::Round($mbx.ArchiveQuotaMBComputed / 1024, 3) } else { 0 }
                RetentionPolicy             = $mbx.RetentionPolicy
                RetentionHoldEnabled        = $mbx.RetentionHoldEnabled
                LitigationHoldEnabled       = $mbx.LitigationHoldEnabled
                LitigationHoldDuration      = $mbx.LitigationHoldDuration
                WhenMailboxCreated          = $mbx.WhenMailboxCreated
                Status                      = if ($pct -ge 96) { 'CRITICAL' } elseif ($pct -ge 90) { 'HIGH' } elseif ($pct -ge 80) { 'WARNING' } else { 'OK' }
                Timestamp                   = $ts
            })
        } else {
            $gaps.Add($mbx)
        }
    }

    # ── Phase 4: Gap fill via live EXO Admin REST API (async, concurrency=8) ──
    if ($gaps.Count -gt 0) {
        Write-Log "$($gaps.Count) mailbox(es) not in Reports API (new <48h or privacy-hidden) — fetching live stats via EXO REST (async, concurrency=8)..." "WARN"

        # Refresh EXO token if it has been running long (Phase 1-3 can take time at scale)
        if (((Get-Date) - $tokenAcquiredAt).TotalMinutes -ge 40) {
            Write-Log "Refreshing EXO token before gap-fill collection..."
            $exoToken        = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
            $tokenAcquiredAt = Get-Date
        }

        $mainItems = foreach ($mbx in $gaps) {
            [PSCustomObject]@{ Key = $mbx.EOID; Uri = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$($mbx.EOID)')/Exchange.GetMailboxStatistics" }
        }
        $mainStats = Invoke-ThrottledGet -Items $mainItems -BearerToken $exoToken -MaxConcurrency 8 -ProgressActivity "Phase 4 gap-fill mailbox stats"

        $archiveCandidates = @($gaps | Where-Object { $_.ArchiveStatus -eq 'Active' -and $_.ArchiveGuid })
        $archStats = @{}
        if ($archiveCandidates.Count -gt 0) {
            $archItems = foreach ($mbx in $archiveCandidates) {
                [PSCustomObject]@{ Key = $mbx.EOID; Uri = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$($mbx.ArchiveGuid)')/Exchange.GetMailboxStatistics" }
            }
            $archStats = Invoke-ThrottledGet -Items $archItems -BearerToken $exoToken -MaxConcurrency 8 -ProgressActivity "Phase 4 gap-fill archive stats"
        }

        $gapErrors = 0
        foreach ($mbx in $gaps) {
            $main = $mainStats[$mbx.EOID]
            if (-not $main -or -not $main.Success) {
                $gapErrors++
                $results.Add([PSCustomObject]@{
                    EOID = $mbx.EOID; DisplayName = $mbx.DisplayName; UPN = $mbx.UserPrincipalName
                    UsedMB = -1; UsedGB = -1; QuotaMB = -1; QuotaGB = -1; UsagePercent = -1
                    ItemCount = $null; DeletedItemCount = $null; DeletedItemSizeMB = $null; LastActivityDate = $null
                    ArchiveStatus = $mbx.ArchiveStatus; ArchiveEnabled = $false; AutoExpandingArchiveEnabled = $false
                    ArchiveUsedMB = $null; ArchiveUsedGB = $null; ArchiveQuotaMB = 0; ArchiveQuotaGB = 0
                    RetentionPolicy = $null; RetentionHoldEnabled = $false
                    LitigationHoldEnabled = $false; LitigationHoldDuration = $null
                    WhenMailboxCreated = $mbx.WhenMailboxCreated
                    Status = 'ERROR'; Timestamp = $ts
                })
                continue
            }

            $stats       = ConvertFrom-StatsJson $main.Body
            $usedMB      = ConvertTo-MB (Get-JsonProp $stats 'TotalItemSize')
            $quotaMB     = $mbx.QuotaMBComputed
            $pct         = Get-UsagePercent -UsedMB $usedMB -QuotaMB $quotaMB
            $archEnabled = ($mbx.ArchiveStatus -eq 'Active')

            $archUsedMB = $null; $archUsedGB = $null
            if ($archEnabled) {
                if (-not $mbx.ArchiveGuid) {
                    $archUsedMB = -1; $archUsedGB = -1
                } else {
                    $a = $archStats[$mbx.EOID]
                    if ($a -and $a.Success) {
                        $aItem      = ConvertFrom-StatsJson $a.Body
                        $archUsedMB = ConvertTo-MB (Get-JsonProp $aItem 'TotalItemSize')
                        $archUsedGB = [math]::Round($archUsedMB / 1024, 3)
                    } elseif ($a -and $a.StatusCode -eq 404) {
                        $archUsedMB = 0; $archUsedGB = 0
                    } else {
                        $archUsedMB = -1; $archUsedGB = -1
                    }
                }
            }

            $results.Add([PSCustomObject]@{
                EOID                        = $mbx.EOID
                DisplayName                 = $mbx.DisplayName
                UPN                         = $mbx.UserPrincipalName
                UsedMB                      = $usedMB
                UsedGB                      = [math]::Round($usedMB  / 1024, 3)
                QuotaMB                     = $quotaMB
                QuotaGB                     = [math]::Round($quotaMB / 1024, 3)
                UsagePercent                = $pct
                ItemCount                   = $(if ($null -ne (Get-JsonProp $stats 'ItemCount')) { Get-JsonProp $stats 'ItemCount' } else { 0 })
                DeletedItemCount            = $(if ($null -ne (Get-JsonProp $stats 'DeletedItemCount')) { Get-JsonProp $stats 'DeletedItemCount' } else { 0 })
                DeletedItemSizeMB           = ConvertTo-MB (Get-JsonProp $stats 'TotalDeletedItemSize')
                LastActivityDate            = Get-JsonProp $stats 'LastLogonTime'
                ArchiveStatus               = $mbx.ArchiveStatus
                ArchiveEnabled              = $archEnabled
                AutoExpandingArchiveEnabled = $mbx.AutoExpandingArchiveEnabled
                ArchiveUsedMB               = $archUsedMB
                ArchiveUsedGB               = $archUsedGB
                ArchiveQuotaMB              = $mbx.ArchiveQuotaMBComputed
                ArchiveQuotaGB              = if ($mbx.ArchiveQuotaMBComputed -gt 0) { [math]::Round($mbx.ArchiveQuotaMBComputed / 1024, 3) } else { 0 }
                RetentionPolicy             = $mbx.RetentionPolicy
                RetentionHoldEnabled        = $mbx.RetentionHoldEnabled
                LitigationHoldEnabled       = $mbx.LitigationHoldEnabled
                LitigationHoldDuration      = $mbx.LitigationHoldDuration
                WhenMailboxCreated          = $mbx.WhenMailboxCreated
                Status                      = if ($pct -ge 96) { 'CRITICAL' } elseif ($pct -ge 90) { 'HIGH' } elseif ($pct -ge 80) { 'WARNING' } else { 'OK' }
                Timestamp                   = $ts
            })
        }

        if ($gapErrors -gt 0) {
            Write-Log "$gapErrors of $($gaps.Count) gap-fill mailbox(es) failed (Status = ERROR)." "WARN"
        }
    }

    # ── Phase 5: Archive stats for report-matched mailboxes (async, concurrency=8) ─
    $archiveRows = @($results | Where-Object { $_.ArchiveEnabled -eq $true -and $null -eq $_.ArchiveUsedMB })
    if ($archiveRows.Count -gt 0) {
        Write-Log "Fetching archive stats for $($archiveRows.Count) mailbox(es) with active archives (async, concurrency=8)..."

        # Refresh EXO token if it has been running long (Phase 1-4 can take time at scale)
        if (((Get-Date) - $tokenAcquiredAt).TotalMinutes -ge 40) {
            Write-Log "Refreshing EXO token before archive stats collection..."
            $exoToken = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
        }

        $archItems = [System.Collections.Generic.List[object]]::new()
        foreach ($row in $archiveRows) {
            $archGuid = $archiveGuidByEOID[$row.EOID]
            if ($archGuid) {
                $archItems.Add([PSCustomObject]@{ Key = $row.EOID; Uri = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$archGuid')/Exchange.GetMailboxStatistics" })
            }
        }
        $archiveStats = if ($archItems.Count -gt 0) { Invoke-ThrottledGet -Items $archItems -BearerToken $exoToken -MaxConcurrency 8 -ProgressActivity "Phase 5 archive stats" } else { @{} }

        $archiveErrors = 0
        foreach ($row in $archiveRows) {
            if (-not $archiveGuidByEOID[$row.EOID]) {
                $row.ArchiveUsedMB = -1; $row.ArchiveUsedGB = -1
                $archiveErrors++
                Write-Log "Archive stats failed for $($row.EOID): No ArchiveGuid in Phase 1 data" "WARN"
                continue
            }
            $s = $archiveStats[$row.EOID]
            if ($s -and $s.Success) {
                $item = ConvertFrom-StatsJson $s.Body
                $row.ArchiveUsedMB = ConvertTo-MB (Get-JsonProp $item 'TotalItemSize')
                $row.ArchiveUsedGB = [math]::Round($row.ArchiveUsedMB / 1024, 3)
            } elseif ($s -and $s.StatusCode -eq 404) {
                $row.ArchiveUsedMB = 0; $row.ArchiveUsedGB = 0
            } else {
                $row.ArchiveUsedMB = -1; $row.ArchiveUsedGB = -1
                $archiveErrors++
                Write-Log "Archive stats failed for $($row.EOID): $(if ($s) { $s.Error } else { 'no response' })" "WARN"
            }
        }
        if ($archiveErrors -gt 0) {
            Write-Log "$archiveErrors of $($archiveRows.Count) archive stat fetch(es) failed (ArchiveUsedMB = -1)." "WARN"
        }
    }

    $fromReport = $results.Count - $gaps.Count
    Write-Log "Collection complete: $fromReport from Reports API + $($gaps.Count) live fallback = $($results.Count) total." "SUCCESS"

    # ── Export CSV ────────────────────────────────────────────────────────────
    $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSV written: $exportPath ($($results.Count) rows)." "SUCCESS"

    # ── Upload to Blob Storage ────────────────────────────────────────────────
    Write-Log "Uploading to Blob Storage ($storageAcctName / $storageContainer)..."
    $blobName = "MailboxOverview_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Set-AzStorageBlobContent -Container $storageContainer -File $exportPath -Blob $blobName -Context $storageCtx -Force | Out-Null
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

    # ── Trigger Send-ReportNotification ───────────────────────────────────────
    Write-Log "Triggering Send-ReportNotification..." "DEBUG"
    Start-AzAutomationRunbook `
        -ResourceGroupName     $repAccountRG `
        -AutomationAccountName $repAccountName `
        -Name                  "Send-ReportNotification" `
        -Parameters @{
            PortalUrl        = if ($SendAsAttachment) { '' } else { $sasUrl }
            BlobName         = $blobName
            RowCount         = $results.Count
            SendAsAttachment = $SendAsAttachment
        } | Out-Null
    Write-Log "Send-ReportNotification triggered." "DEBUG"

} catch {
    Write-Log "Script failed: $_" "ERROR"
    throw
} finally {
    Write-Log "=== Generate-MailboxReport Finished ==="
}
