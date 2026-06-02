# EXO Automation Runbooks — Deployment Guide

Two Azure Automation Accounts running PowerShell 7.4 runbooks with System-Assigned Managed Identity.
No app registrations, no certificates, no secrets.

---

## Solution Files

```
azure-runbooks/
  Generate-MailboxReport.ps1           — weekly mailbox usage report → Blob Storage + email
  Invoke-ScheduledMailboxProvisioning.ps1 — daily archive, retention policy, and litigation hold enforcement
  Send-ReportNotification.ps1          — shared email notification runbook (called by both above)
  Set-AutomationVariables.ps1          — one-time setup: writes all Automation Variables
  README.md                            — this file
```

---

## Architecture

```
  aa-exo-provisioning (Write)          aa-exo-reporting (Read)
  ┌─────────────────────────────┐      ┌──────────────────────────────────────────────┐
  │  Schedule: Daily 02:00 UTC  │      │  Schedule: Weekly Monday 06:00 UTC           │
  │                             │      │                                              │
  │  Invoke-Scheduled           │      │  Generate-MailboxReport                      │
  │  MailboxProvisioning        │      │    Reports.Read.All  ──► Microsoft Graph     │
  │    User.Read.All ──► Graph  │      │    Exchange.ManageAsApp ──► EXO REST API     │
  │    Exchange.ManageAsApp     │      │    upload CSV ──────────► Blob Storage       │
  │         ──► EXO REST API    │      │         │                                    │
  │             │               │      │         ▼                                    │
  │             │ cross-account │      │  Send-ReportNotification                     │
  │             └───────────────┼──────►    Mail.Send ─────────► report recipients   │
  └─────────────────────────────┘      └──────────────────────────────────────────────┘
                                                  ▲
                                       IT Reviewers AD group
                                       (Storage Blob Data Reader on reports container)
```

### Why two accounts

Each account has its own execution boundary, billing scope, and audit trail.
The provisioning identity never touches email — `Mail.Send` lives only on the reporting account.
The reporting identity never writes mailbox configuration — `Exchange.ManageAsApp` is present on both
but the runbooks are scoped to their respective operations.

### Why two runbooks in the reporting account

`Send-ReportNotification` is shared: the provisioning account calls it cross-account via
`Start-AzAutomationRunbook`. This lets both accounts reuse the same email delivery logic
without duplicating `Mail.Send` permission. Recipients are stored as an Automation Variable
(JSON array) — add or remove a recipient without touching any runbook.

---

## Prerequisites

```powershell
# Required PowerShell modules (local machine, run once)
Install-Module Az                        -Scope CurrentUser -Force
Install-Module Microsoft.Graph           -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement  -Scope CurrentUser -Force
```

Required access:
- Azure subscription — Owner or Contributor on the target subscription
- Entra ID — Global Administrator or Privileged Role Administrator (to grant app roles to MIs)
- Exchange Online Administrator (to create Application Access Policy)

---

## Deployment

### Step 0 — Connect and set shared variables

Run once at the start of your session. All subsequent steps reference these variables.

```powershell
Connect-AzAccount
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All","Application.Read.All","RoleManagement.ReadWrite.Directory"

$sub              = (Get-AzContext).Subscription.Id
$rg               = "rg-exo-automation"
$loc              = "eastus"

$provAccountName  = "aa-exo-provisioning"
$repAccountName   = "aa-exo-reporting"
$storageAcctName  = "stexoreports"        # must be globally unique — append initials if taken
$storageContainer = "reports"

Write-Host "Subscription: $((Get-AzContext).Subscription.Name) ($sub)"
```

---

### Step 1 — Create resource group

```powershell
New-AzResourceGroup -Name $rg -Location $loc
```

---

### Step 2 — Create Automation Accounts and enable Managed Identity

```powershell
New-AzAutomationAccount -ResourceGroupName $rg -Name $provAccountName -Location $loc -Plan Basic
New-AzAutomationAccount -ResourceGroupName $rg -Name $repAccountName  -Location $loc -Plan Basic

Set-AzAutomationAccount -ResourceGroupName $rg -Name $provAccountName -AssignSystemIdentity
Set-AzAutomationAccount -ResourceGroupName $rg -Name $repAccountName  -AssignSystemIdentity
```

Enabling Managed Identity auto-creates an Enterprise Application entry in Entra for each account.
No manual app registration is needed.

Capture the Object IDs — required for all permission grants in Steps 4–8:

```powershell
$provMiId = (Get-AzAutomationAccount -ResourceGroupName $rg -Name $provAccountName).Identity.PrincipalId
$repMiId  = (Get-AzAutomationAccount -ResourceGroupName $rg -Name $repAccountName).Identity.PrincipalId

Write-Host "Provisioning MI Object ID : $provMiId"
Write-Host "Reporting MI Object ID    : $repMiId"
```

Save these values — you will need them in Steps 4–8.

---

### Step 3 — (Production only) Set up PIM for Automation Contributor

> Skip for simulation — your Owner access covers everything.

Grants eligible (not active) access so team members can self-activate when editing runbooks.

**Portal method:**
1. Resource group `rg-exo-automation` → **Access control (IAM)** → **Add role assignment**
2. Role: **Automation Contributor** — Assignment type: **PIM Eligible**
3. Members: add each team member
4. Settings: max activation 4 hours, require justification, require MFA on activation

**PowerShell method:**
```powershell
$roleDefId = (Get-AzRoleDefinition -Name "Automation Contributor").Id
$rgScope   = (Get-AzResourceGroup -Name $rg).ResourceId

$members = @("admin@contoso.com")   # add team members here

foreach ($upn in $members) {
    $userId = (Get-MgUser -Filter "userPrincipalName eq '$upn'").Id
    New-AzRoleEligibilityScheduleRequest `
        -Scope $rgScope `
        -PrincipalId $userId `
        -RoleDefinitionId $roleDefId `
        -RequestType AdminAssign `
        -ScheduleInfoStartDateTime (Get-Date) `
        -ExpirationEndDateTime (Get-Date).AddYears(1) `
        -ExpirationType AfterDateTime `
        -Justification "Automation team — PIM eligible for runbook management"
}
```

After PIM is in place, remove standing Owner/Contributor at the RG scope (production only).

---

### Step 4 — Grant Microsoft Graph permissions

Run as Global Administrator. There is no portal blade for this — PowerShell only.

```powershell
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$graph   = Get-MgServicePrincipal -ServicePrincipalId $graphSp.Id

function Grant-GraphRole {
    param([string]$MiId, [string]$RoleName)
    $role = $graph.AppRoles | Where-Object { $_.Value -eq $RoleName }
    if (-not $role) { throw "Role '$RoleName' not found on Microsoft Graph service principal" }
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $MiId `
        -PrincipalId $MiId `
        -ResourceId $graph.Id `
        -AppRoleId $role.Id
}

# Provisioning MI: enumerate licensed users only
Grant-GraphRole -MiId $provMiId -RoleName "User.Read.All"

# Reporting MI: pull usage data and send email
Grant-GraphRole -MiId $repMiId -RoleName "Reports.Read.All"
Grant-GraphRole -MiId $repMiId -RoleName "Mail.Send"
```

`Mail.Send` is intentionally NOT granted to the provisioning MI — email is delegated to the
reporting account via `Send-ReportNotification`.

---

### Step 5 — Grant Exchange.ManageAsApp to both MIs

```powershell
$exoSp   = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'"
$exo     = Get-MgServicePrincipal -ServicePrincipalId $exoSp.Id
$exoRole = $exo.AppRoles | Where-Object { $_.Value -eq "Exchange.ManageAsApp" }

foreach ($miId in @($provMiId, $repMiId)) {
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $miId `
        -PrincipalId $miId `
        -ResourceId $exo.Id `
        -AppRoleId $exoRole.Id
}
```

Exchange Online has no read-only equivalent of `Exchange.ManageAsApp`. Both accounts carry this
permission; separation is enforced at the runbook level.

---

### Step 6 — Assign Exchange Recipient Administrator directory role

Required by Exchange Online for any app-only access, including read operations.

```powershell
$role = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq "Exchange Recipient Administrator" }

if (-not $role) {
    $template = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq "Exchange Recipient Administrator" }
    $role = New-MgDirectoryRole -RoleTemplateId $template.Id
}

foreach ($miId in @($provMiId, $repMiId)) {
    Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($role.Id)/members/`$ref" `
        -Body @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$miId" }
}
```

---

### Step 7 — Grant Automation Operator on the reporting account

Both MIs need **Automation Operator** on `aa-exo-reporting` to call `Start-AzAutomationRunbook`.

> **Note:** **Automation Job Operator** is not sufficient. `Start-AzAutomationRunbook` internally
> reads the runbook before creating the job (`runbooks/read`), which only **Automation Operator**
> and above provides.

```powershell
$reportingAccountScope = "/subscriptions/$sub/resourceGroups/$rg/providers/Microsoft.Automation/automationAccounts/$repAccountName"

# Provisioning MI — cross-account call (aa-exo-provisioning → Send-ReportNotification)
New-AzRoleAssignment `
    -ObjectId            $provMiId `
    -RoleDefinitionName  "Automation Operator" `
    -Scope               $reportingAccountScope

# Reporting MI — self-account call (Generate-MailboxReport → Send-ReportNotification)
New-AzRoleAssignment `
    -ObjectId            $repMiId `
    -RoleDefinitionName  "Automation Operator" `
    -Scope               $reportingAccountScope
```

---

### Step 8 — Create Blob Storage account and container

```powershell
$storageAccount = New-AzStorageAccount `
    -ResourceGroupName   $rg `
    -Name                $storageAcctName `
    -Location            $loc `
    -SkuName             Standard_LRS `
    -Kind                StorageV2 `
    -AllowBlobPublicAccess $false `
    -MinimumTlsVersion   TLS1_2

$ctx = $storageAccount.Context
New-AzStorageContainer -Name $storageContainer -Context $ctx -Permission Off
```

Grant both MIs upload access (provisioning runbook also uploads its log CSV):

```powershell
$storageScope   = "/subscriptions/$sub/resourceGroups/$rg/providers/Microsoft.Storage/storageAccounts/$storageAcctName"
$containerScope = "$storageScope/blobServices/default/containers/$storageContainer"

New-AzRoleAssignment `
    -ObjectId           $repMiId `
    -RoleDefinitionName "Storage Blob Data Contributor" `
    -Scope              $containerScope

New-AzRoleAssignment `
    -ObjectId           $provMiId `
    -RoleDefinitionName "Storage Blob Data Contributor" `
    -Scope              $containerScope
```

Grant both MIs permission to generate User Delegation SAS tokens. Must be at the **storage account**
scope — container scope is not sufficient for this role:

```powershell
New-AzRoleAssignment `
    -ObjectId           $repMiId `
    -RoleDefinitionName "Storage Blob Delegator" `
    -Scope              $storageScope

New-AzRoleAssignment `
    -ObjectId           $provMiId `
    -RoleDefinitionName "Storage Blob Delegator" `
    -Scope              $storageScope
```

Grant IT Reviewers read access via an AD group:

```powershell
$reviewersGroupId = "<object-id-of-IT-Reviewers-AD-group>"

New-AzRoleAssignment `
    -ObjectId           $reviewersGroupId `
    -RoleDefinitionName "Storage Blob Data Reader" `
    -Scope              $containerScope
```

> **Simulation:** Use your own account as the reviewer:
> ```powershell
> $myId = (Get-AzADUser -UserPrincipalName (Get-AzContext).Account.Id).Id
> New-AzRoleAssignment -ObjectId $myId -RoleDefinitionName "Storage Blob Data Reader" -Scope $containerScope
> ```

---

### Step 9 — Scope Mail.Send with Application Access Policy

Restricts the reporting MI to sending email only from the designated sender mailbox.

```powershell
Connect-ExchangeOnline

New-DistributionGroup -Name "AutomationSenders" -Type Security
Add-DistributionGroupMember -Identity "AutomationSenders" -Member "reports@contoso.com"

$repAppId = (Get-MgServicePrincipal -ServicePrincipalId $repMiId).AppId

New-ApplicationAccessPolicy `
    -AppId            $repAppId `
    -PolicyScopeGroupId "AutomationSenders" `
    -AccessRight      RestrictAccess `
    -Description      "Restrict reporting MI Mail.Send to sender mailbox only"

# Verify (allow a few minutes for policy to propagate)
Test-ApplicationAccessPolicy -AppId $repAppId -Identity "reports@contoso.com"
```

---

### Step 10 — Verify PowerShell modules

The full `Az` module set is pre-installed in all PS 7.x Automation Accounts.
**No additional modules are required.** All Exchange Online operations use the EXO Admin REST API
directly with a Managed Identity token — the `ExchangeOnlineManagement` PowerShell module is
not imported or used by any runbook.

Verify the built-in modules are present if needed:

```powershell
foreach ($account in @($provAccountName, $repAccountName)) {
    Write-Host "`n$account"
    Get-AzAutomationModule -ResourceGroupName $rg -AutomationAccountName $account |
        Where-Object { $_.Name -in @("Az.Storage","Az.Automation","Az.Accounts") } |
        Select-Object Name, ProvisioningState
}
```

---

### Step 11 — Set Automation Variables

Open `Set-AutomationVariables.ps1` and fill in:

- **Section 1** — infrastructure names (edit only if you used different names than the defaults)
- **Section 2** — tenant-specific values: tenant ID, organization domain, sender/recipient email,
  retention policy name, report recipients

Then run it:

```powershell
.\azure-runbooks\Set-AutomationVariables.ps1
```

The script confirms your Azure context before writing and prints all variables at the end.
It is safe to re-run — existing values are overwritten.

#### Automation Variables reference

**aa-exo-provisioning**

| Variable | Description |
|---|---|
| `TenantId` | Azure AD directory ID (Entra ID → Overview → Tenant ID) |
| `Organization` | `.onmicrosoft.com` domain |
| `SenderEmail` | Licensed mailbox UPN used as FROM address |
| `RecipientEmail` | TO address for provisioning completion notification |
| `RetentionPolicyName` | Exact name of the EXO retention policy to assign |
| `ReportingAccountRG` | Resource group of `aa-exo-reporting` |
| `ReportingAccountName` | Name of the reporting Automation Account |
| `StorageAccountName` | Blob Storage account name |
| `StorageContainer` | Container name (`reports`) |
| `LitigationHoldDuration` | **(Optional)** Hold duration in days (e.g. `2555` = 7 years) or `Unlimited`. If absent, defaults to Unlimited. |

**aa-exo-reporting**

| Variable | Description |
|---|---|
| `TenantId` | Azure AD directory ID |
| `Organization` | `.onmicrosoft.com` domain |
| `SenderEmail` | Licensed mailbox UPN used as FROM address |
| `StorageAccountName` | Blob Storage account name |
| `StorageContainer` | Container name (`reports`) |
| `ReportRecipients` | JSON array of recipient addresses: `["a@x.com","b@x.com"]` |
| `AutomationAccountRG` | Resource group of `aa-exo-reporting` (self-reference for runbook invocation) |
| `AutomationAccountName` | Name of `aa-exo-reporting` (self-reference) |

To update recipients later without re-running the full script:

```powershell
Set-AzAutomationVariable -ResourceGroupName $rg -AutomationAccountName $repAccountName `
    -Name "ReportRecipients" -Encrypted $false `
    -Value '["it-team@contoso.com","manager@contoso.com"]'
```

---

### Step 12 — Upload and publish runbooks

```powershell
Import-AzAutomationRunbook `
    -ResourceGroupName $rg -AutomationAccountName $provAccountName `
    -Name "Invoke-ScheduledMailboxProvisioning" `
    -Path ".\azure-runbooks\Invoke-ScheduledMailboxProvisioning.ps1" `
    -Type PowerShell72 -Published

Import-AzAutomationRunbook `
    -ResourceGroupName $rg -AutomationAccountName $repAccountName `
    -Name "Generate-MailboxReport" `
    -Path ".\azure-runbooks\Generate-MailboxReport.ps1" `
    -Type PowerShell72 -Published

Import-AzAutomationRunbook `
    -ResourceGroupName $rg -AutomationAccountName $repAccountName `
    -Name "Send-ReportNotification" `
    -Path ".\azure-runbooks\Send-ReportNotification.ps1" `
    -Type PowerShell72 -Published
```

**Editing runbooks after upload:**
- **Portal:** Automation Account → Runbooks → select → Edit → Publish (requires PIM activation in production)
- **VS Code:** Install the [Azure Automation extension](https://marketplace.visualstudio.com/items?itemName=azure-automation.vscode-azureautomation), PIM-activate, sign in to the Azure extension, open and push from the Azure pane

---

### Step 13 — Test each runbook manually

Portal: **Automation Account → Runbooks → select → Test pane → Start**

Test in this order:

1. **`Send-ReportNotification`** (reporting account) — pass a dummy `PortalUrl` (any string) and
   `BlobName`; verify email is received at all configured recipients with a download link in the body.

2. **`Generate-MailboxReport`** (reporting account) — no parameters required; verify:
   - Status: Completed
   - CSV blob visible in the storage container
   - `SAS URL (24h): generated.` logged in output (actual URL is not logged — it is a bearer token)
   - `Send-ReportNotification` job appears in the reporting account Jobs list
   - Email received with a working 24-hour download link

3. **`Invoke-ScheduledMailboxProvisioning`** (provisioning account) — no parameters required; verify:
   - Status: Completed
   - Mailbox counts logged in output (applied / no-op / errors)
   - CSV blob visible in the storage container
   - `Send-ReportNotification` job appears in the reporting account Jobs list
   - Email received with a working 24-hour download link

> Verify blob upload directly if email recipients aren't configured yet:
> ```powershell
> Get-AzStorageBlob -Container $storageContainer `
>     -Context (Get-AzStorageAccount -ResourceGroupName $rg -Name $storageAcctName).Context |
>     Select-Object Name, LastModified, Length
> ```

---

### Step 14 — Create schedules

```powershell
# Provisioning: daily at 02:00 UTC
New-AzAutomationSchedule `
    -ResourceGroupName $rg -AutomationAccountName $provAccountName `
    -Name "Daily-0200-UTC" `
    -StartTime ([datetime]::Today.AddDays(1).AddHours(2)) `
    -DayInterval 1 -TimeZone "UTC"

Register-AzAutomationScheduledRunbook `
    -ResourceGroupName $rg -AutomationAccountName $provAccountName `
    -RunbookName "Invoke-ScheduledMailboxProvisioning" -ScheduleName "Daily-0200-UTC" `
    -Parameters @{ SendAsAttachment = $false }   # set to $true to receive CSV as email attachment

# Reporting: weekly Monday at 06:00 UTC
New-AzAutomationSchedule `
    -ResourceGroupName $rg -AutomationAccountName $repAccountName `
    -Name "Weekly-Mon-0600-UTC" `
    -StartTime ([datetime]::Today.AddDays(1).AddHours(6)) `
    -WeekInterval 1 -DaysOfWeek Monday -TimeZone "UTC"

Register-AzAutomationScheduledRunbook `
    -ResourceGroupName $rg -AutomationAccountName $repAccountName `
    -RunbookName "Generate-MailboxReport" -ScheduleName "Weekly-Mon-0600-UTC" `
    -Parameters @{ SendAsAttachment = $false }   # set to $true to receive CSV as email attachment
```

---

### Step 15 — Configure failure alerting

Portal: **Automation Account → Alerts → New alert rule** — repeat for both accounts.

- Signal: `Total Jobs`
- Condition: `Status = Failed`, count > 0, evaluation every 5 minutes
- Action group: email or Teams webhook to the ops team

---

## Permission Model Summary

### aa-exo-provisioning

| Resource | Permission | Purpose |
|---|---|---|
| Microsoft Graph | `User.Read.All` | Enumerate licensed member users |
| Office 365 Exchange Online | `Exchange.ManageAsApp` | Enable archive, auto-expand, set retention via EXO REST |
| Entra Directory | Exchange Recipient Administrator | Required by EXO for app-only access |
| Storage container (Azure RBAC) | Storage Blob Data Contributor | Upload provisioning log CSV |
| Storage account (Azure RBAC) | Storage Blob Delegator | Generate 24-hour User Delegation SAS URL |
| aa-exo-reporting (Azure RBAC) | Automation Operator | Cross-account `Start-AzAutomationRunbook` |

### aa-exo-reporting

| Resource | Permission | Purpose |
|---|---|---|
| Microsoft Graph | `Reports.Read.All` | Pull mailbox usage data from Graph Reports API |
| Microsoft Graph | `Mail.Send` | Send notification email from sender mailbox |
| Office 365 Exchange Online | `Exchange.ManageAsApp` | EXO Admin REST API for live mailbox statistics (no module required) |
| Entra Directory | Exchange Recipient Administrator | Required by EXO for app-only REST API access |
| Storage container (Azure RBAC) | Storage Blob Data Contributor | Upload CSV from runbook |
| Storage account (Azure RBAC) | Storage Blob Delegator | Generate 24-hour User Delegation SAS URL |
| aa-exo-reporting (Azure RBAC) | Automation Operator | Self-account `Start-AzAutomationRunbook` |

`Mail.Send` is additionally restricted by an Exchange **Application Access Policy** scoped to the
single sender mailbox — the reporting identity cannot send as any other user in the tenant.

---

## Runbook Parameters

### Generate-MailboxReport.ps1

| Parameter | Default | Description |
|---|---|---|
| `ReportsPeriod` | `D7` | Graph Reports API period: `D7`, `D30`, `D90`, `D180` |
| `SendAsAttachment` | `$false` | Attach CSV to email instead of including a blob URL. Blob is always uploaded regardless. |

### Invoke-ScheduledMailboxProvisioning.ps1

| Parameter | Default | Description |
|---|---|---|
| `SkipArchive` | off | Skip enabling archive mailboxes |
| `SkipAutoExpand` | off | Skip enabling auto-expanding archive |
| `SkipRetentionPolicy` | off | Skip assigning the retention policy |
| `SkipLitigationHold` | off | Skip enabling Litigation Hold |
| `SendAsAttachment` | `$false` | Attach CSV to email instead of including a blob URL. Blob is always uploaded regardless. |

### Send-ReportNotification.ps1

| Parameter | Description |
|---|---|
| `PortalUrl` | Azure portal link to the blob container (reporting use case) |
| `BlobName` | File name shown in the email body |
| `RowCount` | Row count shown in the email summary |
| `Subject` | Email subject (defaults to "Mailbox Overview Report — \<date\>") |
| `Body` | HTML body fallback (used when both PortalUrl and SendAsAttachment are not set) |
| `SendAsAttachment` | When `$true`, downloads the blob and attaches the CSV. Blob is always uploaded regardless. |

---

## Deployment Checklist

### Infrastructure
- [ ] Resource group `rg-exo-automation` created
- [ ] Both Automation Accounts created with System-Assigned MI enabled
- [ ] MI Object IDs captured: `$provMiId` / `$repMiId`
- [ ] (Production) PIM-eligible Automation Contributor assigned to team on `rg-exo-automation`
- [ ] (Production) Standing Owner/Contributor removed from RG scope after PIM is in place

### Permissions
- [ ] Graph `User.Read.All` → provisioning MI
- [ ] Graph `Reports.Read.All` + `Mail.Send` → reporting MI
- [ ] EXO `Exchange.ManageAsApp` → both MIs
- [ ] Entra Exchange Recipient Administrator → both MIs
- [ ] Azure RBAC **Automation Operator** → provisioning MI on reporting account
- [ ] Azure RBAC **Automation Operator** → reporting MI on reporting account
- [ ] Azure RBAC Storage Blob Data Contributor → reporting MI on reports container
- [ ] Azure RBAC Storage Blob Data Contributor → provisioning MI on reports container
- [ ] Azure RBAC Storage Blob Delegator → reporting MI on storage account
- [ ] Azure RBAC Storage Blob Delegator → provisioning MI on storage account
- [ ] Azure RBAC Storage Blob Data Reader → IT Reviewers AD group on reports container
- [ ] Exchange Application Access Policy scoping Mail.Send to sender mailbox

### Configuration
- [ ] Blob Storage account and `reports` container created (public access off, TLS 1.2)
- [ ] `Set-AutomationVariables.ps1` run — all placeholder values replaced, output verified

### Deployment and validation
- [ ] All three runbooks uploaded and published
- [ ] `Send-ReportNotification` tested — email received at all recipients
- [ ] `Generate-MailboxReport` tested — Completed, CSV in Blob, email delivered
- [ ] `Invoke-ScheduledMailboxProvisioning` tested — Completed, mailboxes processed, notification received
- [ ] Schedules created and registered on both accounts
- [ ] Failure alerting configured on both accounts
- [ ] First scheduled run observed and validated

---

## Troubleshooting

### 403 AuthorizationPermissionMismatch on Blob Storage

**Symptom:** Runbook fails with:

```
This request is not authorized to perform this operation using this permission.
HTTP Status Code: 403 — ErrorCode: AuthorizationPermissionMismatch
```

**Cause:** The Automation Account's Managed Identity is missing one or both Azure RBAC roles on the
storage account. Both runbooks use `-UseConnectedAccount` (Entra ID auth — no storage key), so RBAC
roles are mandatory. The two required roles and their minimum scope are:

| Role | Scope | Required for |
|---|---|---|
| `Storage Blob Data Contributor` | Container | `Set-AzStorageBlobContent` (upload CSV) |
| `Storage Blob Delegator` | Storage account | `New-AzStorageBlobSASToken` (User Delegation SAS) |

**Fix:** Run the block below for the affected account. It re-uses the same shared variables from
Step 0 — connect first if your session has expired.

```powershell
# Identify which MI is missing the roles
# $provMiId — aa-exo-provisioning  (run if Invoke-ScheduledMailboxProvisioning failed)
# $repMiId  — aa-exo-reporting     (run if Generate-MailboxReport failed)

$miId           = $provMiId   # or $repMiId
$storageScope   = "/subscriptions/$sub/resourceGroups/$rg/providers/Microsoft.Storage/storageAccounts/$storageAcctName"
$containerScope = "$storageScope/blobServices/default/containers/$storageContainer"

New-AzRoleAssignment -ObjectId $miId -RoleDefinitionName "Storage Blob Data Contributor" -Scope $containerScope
New-AzRoleAssignment -ObjectId $miId -RoleDefinitionName "Storage Blob Delegator"        -Scope $storageScope
```

> RBAC propagation takes 1–5 minutes. Re-run the failed job after waiting.

**Why `Generate-MailboxReport` may work while `Invoke-ScheduledMailboxProvisioning` fails:**
The two runbooks run under different Managed Identities (different Automation Accounts). The
reporting MI may already have both roles from a prior deployment while the provisioning MI does not.
