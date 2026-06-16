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
      Phase 4 — Gap fill: live EXO Admin REST for mailboxes missing from report
      Phase 5 — Archive stats for report-matched mailboxes with active archives
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
#>

param(
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$ReportsPeriod = 'D7',
    [bool]$SendAsAttachment = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Helpers ───────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
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
    Write-Log "Storage context initialised for $storageAcctName."

    $exportPath = Join-Path $env:TEMP "MailboxOverview_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    # Acquire tokens via Managed Identity
    Write-Log "Acquiring Graph token (Reports.Read.All) via Managed Identity..."
    $graphToken = Get-ManagedIdentityToken -Resource "https://graph.microsoft.com/"
    Write-Log "Graph token acquired." "SUCCESS"

    Write-Log "Acquiring EXO admin token (Exchange.ManageAsApp) via Managed Identity..."
    $exoToken        = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
    $exoHeaders      = @{ Authorization = "Bearer $exoToken" }
    $tokenAcquiredAt = Get-Date
    Write-Log "EXO admin token acquired." "SUCCESS"

    # ── Phase 1: EXO metadata via Admin REST API ───────────────────────────────
    Write-Log "Retrieving mailbox metadata via EXO Admin REST API..."
    $rawMailboxes = [System.Collections.Generic.List[object]]::new()
    $uri = "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox" +
           "?`$filter=RecipientTypeDetails eq 'UserMailbox'" +
           "&`$select=ExternalDirectoryObjectId,UserPrincipalName,DisplayName," +
           "ProhibitSendQuota,ArchiveStatus,ArchiveQuota,AutoExpandingArchiveEnabled," +
           "RetentionPolicy,RetentionHoldEnabled,LitigationHoldEnabled,LitigationHoldDuration" +
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
            AutoExpandingArchiveEnabled = if ($_.PSObject.Properties['AutoExpandingArchiveEnabled']) { [bool]$_.AutoExpandingArchiveEnabled } else { $false }
            RetentionPolicy             = if ($_.PSObject.Properties['RetentionPolicy']) { $_.RetentionPolicy } else { $null }
            RetentionHoldEnabled        = if ($_.PSObject.Properties['RetentionHoldEnabled']) { [bool]$_.RetentionHoldEnabled } else { $false }
            LitigationHoldEnabled       = if ($_.PSObject.Properties['LitigationHoldEnabled']) { [bool]$_.LitigationHoldEnabled } else { $false }
            LitigationHoldDuration      = if ($_.PSObject.Properties['LitigationHoldDuration'] -and $_.LitigationHoldDuration) {
                                              $raw = $_.LitigationHoldDuration.ToString()
                                              if ($raw -match '^(\d+)\.') { [int]$Matches[1] } else { $raw }
                                          } else { $null }
        }
    })

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
                Status                      = if ($pct -ge 96) { 'CRITICAL' } elseif ($pct -ge 90) { 'HIGH' } elseif ($pct -ge 80) { 'WARNING' } else { 'OK' }
                Timestamp                   = $ts
            })
        } else {
            $gaps.Add($mbx)
        }
    }

    # ── Phase 4: Gap fill via live EXO Admin REST API ─────────────────────────
    if ($gaps.Count -gt 0) {
        Write-Log "$($gaps.Count) mailbox(es) not in Reports API (new <48h or privacy-hidden) — fetching live stats via EXO REST..." "WARN"
        foreach ($mbx in $gaps) {
            try {
                $stats       = Invoke-WithRetry { Invoke-RestMethod -Uri "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$($mbx.EOID)')/MailboxStatistics" -Headers $exoHeaders -Method GET }
                $usedMB      = ConvertTo-MB $stats.TotalItemSize
                $quotaMB     = $mbx.QuotaMBComputed
                $pct         = Get-UsagePercent -UsedMB $usedMB -QuotaMB $quotaMB
                $archEnabled = ($mbx.ArchiveStatus -eq 'Active')
                $archUsedMB  = $null; $archUsedGB = $null
                if ($archEnabled) {
                    try {
                        $archStats  = Invoke-WithRetry { Invoke-RestMethod -Uri "https://outlook.office365.com/adminapi/beta/$tenantId/Mailbox('$($mbx.EOID)')/ArchiveMailboxStatistics" -Headers $exoHeaders -Method GET }
                        $archUsedMB = ConvertTo-MB $archStats.TotalItemSize
                        $archUsedGB = [math]::Round($archUsedMB / 1024, 3)
                    } catch { $archUsedMB = -1; $archUsedGB = -1 }
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
                    ItemCount                   = if ($stats.PSObject.Properties['ItemCount']) { $stats.ItemCount } else { 0 }
                    DeletedItemCount            = if ($stats.PSObject.Properties['DeletedItemCount']) { $stats.DeletedItemCount } else { 0 }
                    DeletedItemSizeMB           = ConvertTo-MB $stats.TotalDeletedItemSize
                    LastActivityDate            = if ($stats.PSObject.Properties['LastLogonTime']) { $stats.LastLogonTime } else { $null }
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
                    Status                      = if ($pct -ge 96) { 'CRITICAL' } elseif ($pct -ge 90) { 'HIGH' } elseif ($pct -ge 80) { 'WARNING' } else { 'OK' }
                    Timestamp                   = $ts
                })
            } catch {
                $results.Add([PSCustomObject]@{
                    EOID = $mbx.EOID; DisplayName = $mbx.DisplayName; UPN = $mbx.UserPrincipalName
                    UsedMB = -1; UsedGB = -1; QuotaMB = -1; QuotaGB = -1; UsagePercent = -1
                    ItemCount = $null; DeletedItemCount = $null; DeletedItemSizeMB = $null; LastActivityDate = $null
                    ArchiveStatus = $mbx.ArchiveStatus; ArchiveEnabled = $false; AutoExpandingArchiveEnabled = $false
                    ArchiveUsedMB = $null; ArchiveUsedGB = $null; ArchiveQuotaMB = 0; ArchiveQuotaGB = 0
                    RetentionPolicy = $null; RetentionHoldEnabled = $false
                    LitigationHoldEnabled = $false; LitigationHoldDuration = $null
                    Status = 'ERROR'; Timestamp = $ts
                })
            }
        }
    }

    # ── Phase 5: Archive stats for report-matched mailboxes (parallel) ────────
    $archiveRows = @($results | Where-Object { $_.ArchiveEnabled -eq $true -and $null -eq $_.ArchiveUsedMB })
    if ($archiveRows.Count -gt 0) {
        Write-Log "Fetching archive stats for $($archiveRows.Count) mailbox(es) with active archives (parallel, ThrottleLimit=8)..."

        # Refresh EXO token if it has been running long (Phase 1-4 can take time at scale)
        if (((Get-Date) - $tokenAcquiredAt).TotalMinutes -ge 40) {
            Write-Log "Refreshing EXO token before archive stats collection..."
            $exoToken = Get-ManagedIdentityToken -Resource "https://outlook.office365.com/"
        }

        $p_token    = $exoToken
        $p_tenantId = $tenantId

        # Parallel fetch — ForEach-Object -Parallel returns new objects; mutating $_ inside
        # the block does not propagate back. Collect (EOID, stats) then join in parent scope.
        $archiveStatsList = $archiveRows | ForEach-Object -Parallel {
            $eoid  = $_.EOID
            $token = $using:p_token
            $tid   = $using:p_tenantId

            $usedMB = -1; $usedGB = -1
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    $resp = Invoke-RestMethod `
                        -Uri     "https://outlook.office365.com/adminapi/beta/$tid/Mailbox('$eoid')/ArchiveMailboxStatistics" `
                        -Headers @{ Authorization = "Bearer $token" } `
                        -Method  GET
                    $s = $resp.TotalItemSize
                    $usedMB = if     ($null -eq $s)                                            { 0 }
                              elseif ($s -is [long] -or $s -is [int] -or $s -is [double])     { [math]::Round([double]$s / 1MB, 2) }
                              elseif ($s.ToString() -match '\(([0-9,]+) bytes\)')              { [math]::Round([long]($Matches[1] -replace ',', '') / 1MB, 2) }
                              else   { $p = 0L; if ([long]::TryParse($s.ToString(), [ref]$p)) { [math]::Round($p / 1MB, 2) } else { 0 } }
                    $usedGB = [math]::Round($usedMB / 1024, 3)
                    break
                } catch {
                    if ($attempt -ge 3 -or $_.Exception.Message -notmatch '429|throttl') { break }
                    Start-Sleep -Seconds ([math]::Pow(2, $attempt))
                }
            }
            [PSCustomObject]@{ EOID = $eoid; ArchiveUsedMB = $usedMB; ArchiveUsedGB = $usedGB }
        } -ThrottleLimit 8

        # Join stats back to $results by EOID (parent-scope object refs are still valid here)
        $archiveLookup = @{}
        foreach ($s in $archiveStatsList) { if ($s) { $archiveLookup[$s.EOID] = $s } }
        foreach ($row in $archiveRows) {
            $s = $archiveLookup[$row.EOID]
            if ($s) { $row.ArchiveUsedMB = $s.ArchiveUsedMB; $row.ArchiveUsedGB = $s.ArchiveUsedGB }
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
    Write-Log "SAS URL (24h): generated." "SUCCESS"

    # ── Trigger Send-ReportNotification ───────────────────────────────────────
    Write-Log "Triggering Send-ReportNotification..."
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
    Write-Log "Send-ReportNotification triggered." "SUCCESS"

} catch {
    Write-Log "Script failed: $_" "ERROR"
    throw
} finally {
    Write-Log "=== Generate-MailboxReport Finished ==="
}
