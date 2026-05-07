#Requires -Version 7.4
<#
.SYNOPSIS
    Azure Automation runbook: sends a report notification email via Microsoft Graph.
    Called by Generate-MailboxReport and Invoke-ScheduledMailboxProvisioning via
    Start-AzAutomationRunbook (cross-account).

.DESCRIPTION
    Reads recipients from the ReportRecipients Automation Variable (JSON array).
    Sends an HTML email from the configured sender mailbox using the Managed Identity
    Mail.Send permission scoped by Application Access Policy.

    When SasUrl is provided the body contains a download link (reporting use case).
    When Body is provided directly it is used as-is (provisioning notification use case).
    Access to the report requires Entra authentication — no anonymous SAS links are used.

.PARAMETER PortalUrl
    Azure portal link to the blob storage container. Requires Entra sign-in to access —
    no anonymous access. Omit for provisioning completion notifications.

.PARAMETER BlobName
    Blob file name shown in the email body. Omit when PortalUrl is not used.

.PARAMETER RowCount
    Number of rows in the report — shown in the email body summary line.

.PARAMETER Subject
    Email subject line. Defaults to "Mailbox Overview Report — <date>".

.PARAMETER Body
    Plain HTML body. Used when PortalUrl and SendAsAttachment are both not set.

.PARAMETER SendAsAttachment
    When true, downloads the blob and attaches the CSV to the email instead of including a URL.
    The blob is always uploaded regardless of this parameter — attachment vs URL is email-only.

.NOTES
    Automation Variables required (in the reporting account):
      SenderEmail        — licensed mailbox UPN used as FROM address
      ReportRecipients   — JSON array of recipient addresses
      StorageAccountName — blob storage account (read when SendAsAttachment is true)
      StorageContainer   — blob container name (read when SendAsAttachment is true)
#>

param(
    [string]$PortalUrl       = "",
    [string]$BlobName        = "",
    [int]$RowCount           = 0,
    [string]$Subject         = "",
    [string]$Body            = "",
    [bool]$SendAsAttachment  = $false
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
    (Get-AzAccessToken -ResourceUrl $Resource).Token
}

# ── Entry point ───────────────────────────────────────────────────────────────

try {
    Write-Log "=== Send-ReportNotification Started ==="

    Connect-AzAccount -Identity | Out-Null

    $senderEmail = Get-AutomationVariable -Name "SenderEmail"
    $recipients  = (Get-AutomationVariable -Name "ReportRecipients") | ConvertFrom-Json
    $graphToken  = Get-ManagedIdentityToken -Resource "https://graph.microsoft.com/"

    Write-Log "Sender     : $senderEmail"
    Write-Log "Recipients : $($recipients -join ', ')"

    $emailSubject = if ($Subject) { $Subject } else {
        "Mailbox Overview Report — $(Get-Date -Format 'yyyy-MM-dd')"
    }

    # ── Download blob for attachment (once, before the recipient loop) ─────────
    $attachmentBase64 = $null
    if ($SendAsAttachment -and $BlobName) {
        $storageAcctName  = Get-AutomationVariable -Name "StorageAccountName"
        $storageContainer = Get-AutomationVariable -Name "StorageContainer"
        $storageCtx       = New-AzStorageContext -StorageAccountName $storageAcctName -UseConnectedAccount
        $tempFile         = Join-Path $env:TEMP $BlobName
        Get-AzStorageBlobContent -Container $storageContainer -Blob $BlobName -Destination $tempFile -Context $storageCtx -Force | Out-Null
        $attachmentBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($tempFile))
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        Write-Log "Blob downloaded for attachment: $BlobName ($([math]::Round($attachmentBase64.Length * 3 / 4 / 1KB, 1)) KB)." "SUCCESS"
    }

    $emailBody = if ($SendAsAttachment -and $BlobName) {
        @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#333'>
<h2 style='color:#0078d4'>Report Ready</h2>
<p>The report (<strong>$RowCount rows</strong>) is attached as <code>$BlobName</code>.</p>
<p style='color:#666;font-size:11px'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') UTC</p>
</body></html>
"@
    } elseif ($PortalUrl) {
        @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#333'>
<h2 style='color:#0078d4'>Mailbox Overview Report</h2>
<p>The weekly mailbox usage report (<strong>$RowCount rows</strong>) is ready in Blob Storage.</p>
<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;min-width:300px'>
  <tr style='background:#0078d4;color:white'><th align='left'>File</th><th align='left'>Location</th></tr>
  <tr><td><code>$BlobName</code></td><td><a href='$PortalUrl' style='color:#0078d4'>Download (link valid 24 hours)</a></td></tr>
</table>
<p style='margin-top:12px;color:#666;font-size:11px'>Link expires in 24 hours. Historical reports are accessible via the Azure portal storage browser.</p>
<p style='color:#666;font-size:11px'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') UTC</p>
</body></html>
"@
    } else {
        $Body
    }

    foreach ($recipient in $recipients) {
        $message = @{
            subject      = $emailSubject
            body         = @{ contentType = "HTML"; content = $emailBody }
            toRecipients = @(@{ emailAddress = @{ address = $recipient } })
        }
        if ($attachmentBase64) {
            $message['attachments'] = @(
                @{
                    '@odata.type' = '#microsoft.graph.fileAttachment'
                    name          = $BlobName
                    contentType   = 'text/csv'
                    contentBytes  = $attachmentBase64
                }
            )
        }
        $payload = @{
            message         = $message
            saveToSentItems = $false
        } | ConvertTo-Json -Depth 8

        Invoke-RestMethod `
            -Uri     "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($senderEmail))/sendMail" `
            -Method  POST `
            -Headers @{ Authorization = "Bearer $graphToken"; "Content-Type" = "application/json" } `
            -Body    $payload | Out-Null

        Write-Log "Email sent to $recipient." "SUCCESS"
    }

} catch {
    Write-Log "Script failed: $_" "ERROR"
    throw
} finally {
    Write-Log "=== Send-ReportNotification Finished ==="
}
