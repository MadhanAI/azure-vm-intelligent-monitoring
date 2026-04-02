# Azure VM Performance Report — Windows Server Deployment Guide
**Version 2 · Lighthouse + Microsoft Graph · Windows Task Scheduler**

---

## Architecture

```
Windows Server (your MSP / managing org)
├── Python Scripts  C:\VM-Reports\scripts\
│     ├── main.py             ← entrypoint (called by Task Scheduler)
│     ├── config.py           ← reads ALL config from Windows env vars
│     ├── collect_metrics.py  ← Lighthouse multi-tenant metrics collection
│     ├── generate_report.py  ← .docx with borders, charts, findings
│     └── graph_mailer.py     ← sends report via Microsoft Graph
│
├── Windows System Environment Variables  (set by setup_env.ps1)
│     ├── AZURE_TENANT_ID / CLIENT_ID / CLIENT_SECRET  ← Service Principal
│     ├── GRAPH_CLIENT_ID / CLIENT_SECRET / SENDER_UPN ← App Registration
│     └── TARGET_SUBSCRIPTION_IDS_JSON                 ← explicit allowlist
│
└── Windows Task Scheduler
      └── Azure Monitoring \ Azure VM Performance Report - Monthly
            Trigger:  1st of each month, 06:00 AM
            Action:   run_report.bat → python main.py
            Account:  SYSTEM  (reads Machine-scope env vars)

Azure (read-only via Lighthouse delegation)
├── Customer Tenant A — Subscription LFG Prod
│     VMs + Azure Monitor Metrics + Log Analytics
├── Customer Tenant B — Subscription Client B
│     VMs + Azure Monitor Metrics + Log Analytics
└── ... (only subscriptions in TARGET_SUBSCRIPTION_IDS_JSON are touched)

Microsoft 365 (managing org)
└── App Registration → Mail.Send → Sends .docx report to per-client recipients
```

---

## Key Design Decisions

| Topic | Decision |
|---|---|
| Scheduling | Windows Task Scheduler (SYSTEM account) — no Azure Automation |
| Secrets | Windows System Environment Variables — no Key Vault, no Managed Identity |
| Auth (Azure) | Service Principal with client_credentials grant |
| Auth (Email) | App Registration, Mail.Send application permission, client_credentials |
| Subscription filter | **Explicit allowlist only** — empty list is a hard error |
| Report period | **Exact calendar month** — day 1 00:00:00 to last-day 23:59:59 |
| February | `calendar.monthrange(year, month)` — returns 28 or 29 correctly |
| Page borders | XML injection into `w:pgBorders` in every `sectPr` — all pages |

---

## Prerequisites

| What | Where | Notes |
|---|---|---|
| Windows Server 2019/2022 | Your infrastructure | Any edition |
| Python 3.10+ | python.org | Add to PATH during install |
| Azure Service Principal | Managing tenant | Permissions via Lighthouse |
| Azure Lighthouse delegations | Each customer subscription | See Part 1 |
| App Registration | Managing tenant | Mail.Send + admin consent |
| Exchange Online mailbox | Managing tenant M365 | Sender UPN must exist |

---

## Part 1 — Azure Lighthouse Delegation

### Roles to delegate per customer subscription

| Role | Definition ID | Purpose |
|---|---|---|
| Reader | `acdd72a7-3385-48ef-bd42-f606fba81ae7` | VM discovery, ARM API |
| Monitoring Reader | `43d0d8ad-25c7-4714-9337-8ba259a9fe05` | Azure Monitor Metrics |
| Log Analytics Reader | `73c42c96-874c-492b-b04d-ab87d138a893` | KQL disk utilisation |

### ARM template — deploy on each customer subscription

```json
{
  "$schema": "https://schema.management.azure.com/schemas/2018-06-01/subscriptionDeploymentTemplate.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
    "mspTenantId": { "type": "string" },
    "spObjectId":  { "type": "string" }
  },
  "resources": [{
    "type": "Microsoft.ManagedServices/registrationDefinitions",
    "apiVersion": "2020-02-01-preview",
    "name": "[guid(subscription().subscriptionId, 'vm-report')]",
    "properties": {
      "registrationDefinitionName": "Azure VM Monitoring (MSP Read-Only)",
      "managedByTenantId": "[parameters('mspTenantId')]",
      "authorizations": [
        { "principalId": "[parameters('spObjectId')]", "principalIdDisplayName": "VM Report SP",
          "roleDefinitionId": "acdd72a7-3385-48ef-bd42-f606fba81ae7" },
        { "principalId": "[parameters('spObjectId')]", "principalIdDisplayName": "VM Report SP",
          "roleDefinitionId": "43d0d8ad-25c7-4714-9337-8ba259a9fe05" },
        { "principalId": "[parameters('spObjectId')]", "principalIdDisplayName": "VM Report SP",
          "roleDefinitionId": "73c42c96-874c-492b-b04d-ab87d138a893" }
      ]
    }
  }]
}
```

Deploy (run from the customer subscription context):

```powershell
$SP_OID = (Get-AzADServicePrincipal -ApplicationId "<YOUR_CLIENT_ID>").Id
az deployment sub create `
  --subscription "<CUSTOMER-SUB-ID>" `
  --location "eastus" `
  --template-file lighthouse_offer.json `
  --parameters mspTenantId="<YOUR-TENANT-ID>" spObjectId="$SP_OID"
```

---

## Part 2 — Service Principal

```powershell
$SP = az ad sp create-for-rbac --name "vm-report-sp" --skip-assignment --output json | ConvertFrom-Json
Write-Host "AZURE_TENANT_ID:     $(az account show --query tenantId -o tsv)"
Write-Host "AZURE_CLIENT_ID:     $($SP.appId)"
Write-Host "AZURE_CLIENT_SECRET: $($SP.password)"
Write-Host "SP Object ID:        $(az ad sp show --id $SP.appId --query id -o tsv)"
```

---

## Part 3 — App Registration (Microsoft Graph / Mail.Send)

```powershell
$APP_ID = az ad app create --display-name "VM-Report-Mailer" --query appId -o tsv
az ad sp create --id $APP_ID
az ad app permission add --id $APP_ID `
  --api 00000003-0000-0000-c000-000000000000 `
  --api-permissions b633e1c5-b582-4048-a93e-9f11b44c7e96=Role
az ad app permission admin-consent --id $APP_ID   # Requires Global Admin
$SECRET = az ad app credential reset --id $APP_ID --years 2 --query password -o tsv
Write-Host "GRAPH_CLIENT_ID:     $APP_ID"
Write-Host "GRAPH_CLIENT_SECRET: $SECRET"
```

---

## Part 4 — Windows Server Setup

### Step 1 — Install Python

Download Python 3.11+ from python.org. During install tick: **Add Python to PATH** and **Install for all users**.

### Step 2 — Copy scripts

```
C:\VM-Reports\scripts\   ← copy all .py, .ps1, .bat, .txt files here
```

### Step 3 — Install packages

```cmd
cd C:\VM-Reports\scripts
pip install -r requirements.txt
```

### Step 4 — Configure environment variables

Edit `setup_env.ps1`, fill in all values in the `CONFIGURE THESE` blocks, then run:

```powershell
# As Administrator
Set-ExecutionPolicy RemoteSigned -Scope Process
.\setup_env.ps1
```

Restart your terminal after running.

### Step 5 — Test

```cmd
cd C:\VM-Reports\scripts
python main.py
```

### Step 6 — Schedule

```powershell
# As Administrator
.\create_scheduled_task.ps1
```

Task created at: **Task Scheduler → Task Scheduler Library → Azure Monitoring**

---

## Part 5 — CLIENT_CONFIGS_JSON Schema

```json
{
  "<SUBSCRIPTION-ID>": {
    "client_name": "CSPL LFG PROD",
    "to_recipients": ["ops@client.com"],
    "cc_recipients": ["mgr@yourmsp.com"],
    "log_analytics_workspace_id": "<WORKSPACE-ID-OR-EMPTY>"
  }
}
```

Every subscription ID listed in `TARGET_SUBSCRIPTION_IDS_JSON` must have a matching entry here — the script hard-errors on startup if any are missing.

---

## Part 6 — Report Period Logic

```
Run date        →  Report month
Any day Mar     →  Feb 01 00:00:00 to Feb 28/29 23:59:59
Any day Oct     →  Sep 01 00:00:00 to Sep 30 23:59:59
Any day Jan     →  Dec 01 00:00:00 to Dec 31 23:59:59

February:
  calendar.monthrange(2025, 2) → (5, 28)   non-leap
  calendar.monthrange(2024, 2) → (3, 29)   leap year

Backfill:
  set REPORT_YEAR=2025
  set REPORT_MONTH=9
  python main.py
```

---

## Part 7 — Environment Variables

### Required

| Variable | Description |
|---|---|
| `AZURE_TENANT_ID` | Your MSP's Azure AD tenant ID |
| `AZURE_CLIENT_ID` | Service Principal application ID |
| `AZURE_CLIENT_SECRET` | Service Principal secret |
| `GRAPH_CLIENT_ID` | App Registration client ID |
| `GRAPH_CLIENT_SECRET` | App Registration secret |
| `GRAPH_SENDER_UPN` | Sending mailbox UPN |
| `TARGET_SUBSCRIPTION_IDS_JSON` | `["sub-id-1","sub-id-2"]` — must be non-empty |
| `CLIENT_CONFIGS_JSON` | Per-client config object (see Part 5) |

### Optional

| Variable | Default | Description |
|---|---|---|
| `GRAPH_TENANT_ID` | Same as `AZURE_TENANT_ID` | If M365 is in a separate tenant |
| `TARGET_RESOURCE_GROUPS_JSON` | `{}` | `{"sub-id": ["rg1","rg2"]}` |
| `REPORT_YEAR` / `REPORT_MONTH` | Auto (previous month) | Backfill override |
| `REPORT_OUTPUT_DIR` | `<scripts>\reports` | Where .docx files are saved |
| `REPORT_LOG_DIR` | `<scripts>\logs` | Where run logs are saved |
| `INTERNAL_CC_EMAILS` | *(empty)* | Comma-separated CC on all emails |
| `CPU_ALERT_THRESHOLD` | `80` | % CPU threshold |
| `MEMORY_MIN_THRESHOLD` | `10` | % free memory threshold |
| `DISK_UTIL_THRESHOLD` | `85` | % disk used threshold |
| `STORAGE_ACCOUNT_NAME` | *(empty)* | Azure Blob for archiving (optional) |
| `TEAMS_WEBHOOK_URL` | *(empty)* | Teams notifications (optional) |

---

## Part 8 — Troubleshooting

| Symptom | Fix |
|---|---|
| `EnvironmentError: ... not set` | Re-run `setup_env.ps1` as Admin; restart terminal |
| `TARGET_SUBSCRIPTION_IDS_JSON must be non-empty` | Add at least one subscription ID |
| `has no entry in CLIENT_CONFIGS_JSON` | Add entry for that subscription ID |
| HTTP 401 on ARM | SP secret expired — regenerate and update env var |
| `[WARN] subscription not found` | Customer hasn't accepted Lighthouse delegation yet |
| HTTP 403 on Monitor API | Add Monitoring Reader to Lighthouse offer |
| HTTP 403 on Log Analytics | Add Log Analytics Reader to Lighthouse offer |
| Graph HTTP 401 | Wrong `GRAPH_CLIENT_ID` or `GRAPH_CLIENT_SECRET` |
| Graph HTTP 403 | Run `az ad app permission admin-consent --id <APP_ID>` |
| Email not received | `GRAPH_SENDER_UPN` mailbox doesn't exist in M365 |
| Disk chart missing | Azure Monitor Agent not deployed on that VM |
| Task Scheduler exits code 1 | Check log file in `REPORT_LOG_DIR` |
| Task Scheduler can't read env vars | Must be **Machine-scope** — re-run `setup_env.ps1` as Admin |
