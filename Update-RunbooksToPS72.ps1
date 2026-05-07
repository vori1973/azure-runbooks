<#
.SYNOPSIS
    Migrates all runbooks in both Automation Accounts from PowerShell 5.1 to PowerShell 7.2.

.DESCRIPTION
    The runbookType property is immutable after creation. This script:
      1. Exports each runbook's current content
      2. Deletes the runbook
      3. Re-imports it with -Type PowerShell72 (the correct PS 7.2 type)
    Schedules linked to a runbook are NOT deleted — they reattach automatically
    when a runbook with the same name is republished.

.EXAMPLE
    Connect-AzAccount
    .\azure-runbooks\Update-RunbooksToPS72.ps1
#>

#region — Config (edit if your names differ from the defaults)
$rg              = "rg-exo-automation"
$provAccountName = "aa-exo-provisioning"
$repAccountName  = "aa-exo-reporting"
#endregion

$ErrorActionPreference = "Stop"

$ctx = Get-AzContext
if (-not $ctx) { throw "Not connected. Run Connect-AzAccount first." }
Write-Host "Subscription : $($ctx.Subscription.Name) ($($ctx.Subscription.Id))"
Write-Host "Resource group: $rg`n"

$tmpDir = Join-Path $env:TEMP "runbook-ps72-migration"
New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

foreach ($accountName in @($provAccountName, $repAccountName)) {
    Write-Host "=== $accountName ===" -ForegroundColor Cyan

    $runbooks = Get-AzAutomationRunbook -ResourceGroupName $rg -AutomationAccountName $accountName |
        Where-Object { $_.RunbookType -eq "PowerShell" }   # only PS 5.1 runbooks

    if (-not $runbooks) {
        Write-Host "  No PowerShell 5.1 runbooks found — skipping."
        continue
    }

    foreach ($rb in $runbooks) {
        Write-Host "  [$($rb.Name)]"

        # 1 — Export current content
        Write-Host "    Exporting..." -NoNewline
        Export-AzAutomationRunbook `
            -ResourceGroupName     $rg `
            -AutomationAccountName $accountName `
            -Name                  $rb.Name `
            -OutputFolder          $tmpDir `
            -Force | Out-Null

        $scriptPath = Join-Path $tmpDir "$($rb.Name).ps1"
        if (-not (Test-Path $scriptPath)) {
            Write-Host " FAILED (no export file)" -ForegroundColor Red
            continue
        }
        Write-Host " OK" -ForegroundColor Green

        # 2 — Delete (required because runbookType is immutable)
        Write-Host "    Deleting..." -NoNewline
        Remove-AzAutomationRunbook `
            -ResourceGroupName     $rg `
            -AutomationAccountName $accountName `
            -Name                  $rb.Name `
            -Force
        Write-Host " OK" -ForegroundColor Green

        # 3 — Re-create as PowerShell72 and publish
        Write-Host "    Importing as PowerShell72..." -NoNewline
        Import-AzAutomationRunbook `
            -ResourceGroupName     $rg `
            -AutomationAccountName $accountName `
            -Name                  $rb.Name `
            -Path                  $scriptPath `
            -Type                  PowerShell72 `
            -Published `
            -Force | Out-Null
        Write-Host " OK" -ForegroundColor Green

        Write-Host "    Done → PowerShell 7.2" -ForegroundColor Green
    }
}

Remove-Item -Path $tmpDir -Recurse -Force
Write-Host "`nDone. Verify in the portal: Automation Account → Runbooks → select a runbook → Runtime version should show 7.2."
