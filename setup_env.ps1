#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Sets Windows System Environment Variables for the Azure VM Report.

.DESCRIPTION
    Only CREDENTIALS are stored as environment variables (all short strings,
    well within Windows' 1,024-character limit per variable).

    Large JSON configurations (client list, subscription list) are stored as
    JSON FILES in the config directory. Edit them with Notepad - no size limit.

    Config files location (VM_REPORT_CONFIG_DIR):
        C:\VM-Reports\config\subscriptions.json     ← list of subscription IDs
        C:\VM-Reports\config\client_configs.json    ← per-client settings

    Run this script ONCE as Administrator.
    After running, restart any open terminals or the Task Scheduler service.

.NOTES
    Fill in ALL values in the CONFIGURE THESE sections before running.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Set-SysEnv($Name, $Value) {
    [System.Environment]::SetEnvironmentVariable($Name, $Value, "Machine")
    Write-Host "  [SET] $Name" -ForegroundColor Green
}
function Set-SysEnvSecret($Name, $Value) {
    [System.Environment]::SetEnvironmentVariable($Name, $Value, "Machine")
    Write-Host "  [SET] $Name  (secret)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  Azure VM Report - Environment Variable Setup"  -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  NOTE: JSON configs are stored as FILES, not env vars."
Write-Host "  No more 1,024-character limit issues."
Write-Host ""


# ══════════════════════════════════════════════════════════════════
# CONFIGURE THESE - Azure Service Principal (ARM + Metrics)
# ══════════════════════════════════════════════════════════════════
$AzureTenantId     = "YOUR-MANAGING-TENANT"
$AzureClientId     = "YOUR-SERVICE-PRINCIPAL"
$AzureClientSecret = "YOUR-SERVICE-PRINCIPAL"

# ══════════════════════════════════════════════════════════════════
# CONFIGURE THESE - Microsoft Graph (Mail.Send)
# ══════════════════════════════════════════════════════════════════
$GraphClientId     = "YOUR-GRAPH-APP"
$GraphClientSecret = "YOUR-GRAPH-APP"
$GraphSenderUpn    = "sender@domain.com"
$GraphTenantId     = ""    # Leave blank if same tenant as Azure above

# ══════════════════════════════════════════════════════════════════
# CONFIGURE THESE - Paths
# ══════════════════════════════════════════════════════════════════
$ScriptsDir      = "C:\Scripts\VM-Reports"     # Where the .py files live
$ConfigDir       = "C:\Scripts\VM-Reports\config"      # Where the JSON config files live
$ReportOutputDir = "C:\Scripts\VM-Reports\output"      # Where .docx reports are saved
$ReportLogDir    = "C:\Scripts\VM-Reports\logs"        # Where run logs are saved

# ══════════════════════════════════════════════════════════════════
# OPTIONAL - Alert thresholds and delivery settings
# ══════════════════════════════════════════════════════════════════
$InternalCcEmails   = "user@domain.com"          # Comma-separated CC on all emails (optional)
$StorageAccountName = ""          # Azure Blob Storage account name (optional)
$StorageContainer   = "vm-reports"
$TeamsWebhookUrl    = ""          # Teams incoming webhook (optional)
$CpuAlertThreshold  = "80"
$MemMinThreshold    = "10"
$DiskUtilThreshold  = "80"


# ── Validate placeholders ─────────────────────────────────────────
$problems = @()
if ($AzureTenantId     -match "YOUR-MANAGING-TENANT")         { $problems += "AZURE_TENANT_ID" }
if ($AzureClientId     -match "YOUR-SERVICE-PRINCIPAL")        { $problems += "AZURE_CLIENT_ID" }
if ($AzureClientSecret -match "YOUR-SERVICE-PRINCIPAL")        { $problems += "AZURE_CLIENT_SECRET" }
if ($GraphClientId     -match "YOUR-GRAPH-APP")               { $problems += "GRAPH_CLIENT_ID" }
if ($GraphClientSecret -match "YOUR-GRAPH-APP")               { $problems += "GRAPH_CLIENT_SECRET" }
if ($GraphSenderUpn    -match "yourdomain\.com")              { $problems += "GRAPH_SENDER_UPN" }

if ($problems.Count -gt 0) {
    Write-Host ""
    Write-Host "[ERROR] The following variables still have placeholder values:" -ForegroundColor Red
    $problems | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Write-Host ""
    Write-Host "Edit setup_env.ps1 and replace placeholders before running." -ForegroundColor Red
    exit 1
}


# ── Write credential env vars ─────────────────────────────────────
Write-Host "Writing Azure credential variables..."
Set-SysEnv       "AZURE_TENANT_ID"    $AzureTenantId
Set-SysEnv       "AZURE_CLIENT_ID"    $AzureClientId
Set-SysEnvSecret "AZURE_CLIENT_SECRET" $AzureClientSecret

Write-Host "`nWriting Microsoft Graph credential variables..."
Set-SysEnv       "GRAPH_CLIENT_ID"     $GraphClientId
Set-SysEnvSecret "GRAPH_CLIENT_SECRET" $GraphClientSecret
Set-SysEnv       "GRAPH_SENDER_UPN"    $GraphSenderUpn
if ($GraphTenantId) {
    Set-SysEnv   "GRAPH_TENANT_ID"     $GraphTenantId
}

Write-Host "`nWriting path variables..."
Set-SysEnv "VM_REPORT_CONFIG_DIR"  $ConfigDir
Set-SysEnv "REPORT_OUTPUT_DIR"     $ReportOutputDir
Set-SysEnv "REPORT_LOG_DIR"        $ReportLogDir

Write-Host "`nWriting optional settings..."
Set-SysEnv "INTERNAL_CC_EMAILS"   $InternalCcEmails
Set-SysEnv "STORAGE_ACCOUNT_NAME" $StorageAccountName
Set-SysEnv "STORAGE_CONTAINER"    $StorageContainer
Set-SysEnv "TEAMS_WEBHOOK_URL"    $TeamsWebhookUrl
Set-SysEnv "CPU_ALERT_THRESHOLD"  $CpuAlertThreshold
Set-SysEnv "MEMORY_MIN_THRESHOLD" $MemMinThreshold
Set-SysEnv "DISK_UTIL_THRESHOLD"  $DiskUtilThreshold


# ── Create directories ────────────────────────────────────────────
Write-Host "`nCreating directories..."
foreach ($dir in @($ScriptsDir, $ConfigDir, $ReportOutputDir, $ReportLogDir)) {
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "  [CREATED] $dir" -ForegroundColor Green
    } else {
        Write-Host "  [EXISTS]  $dir" -ForegroundColor Gray
    }
}


# ── Check config files exist ──────────────────────────────────────
Write-Host "`nChecking config files..."
$subsFile    = Join-Path $ConfigDir "subscriptions.json"
$clientsFile = Join-Path $ConfigDir "client_configs.json"

foreach ($f in @($subsFile, $clientsFile)) {
    if (Test-Path $f) {
        # Validate JSON
        try {
            $null = Get-Content $f -Raw | ConvertFrom-Json
            Write-Host "  [OK]      $f" -ForegroundColor Green
        } catch {
            Write-Host "  [ERROR]   $f - invalid JSON: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "  [MISSING] $f" -ForegroundColor Yellow
        Write-Host "            Copy the provided template file to this location." -ForegroundColor Yellow
    }
}


# ── Python check ──────────────────────────────────────────────────
Write-Host "`nChecking Python..."
try {
    $v = & python --version 2>&1
    Write-Host "  [OK] $v" -ForegroundColor Green
} catch {
    Write-Host "  [WARN] Python not found. Install Python 3.10+ and add to PATH." -ForegroundColor Yellow
}


Write-Host ""
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  Setup complete." -ForegroundColor Cyan
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor White
Write-Host "  1. Copy subscriptions.json and client_configs.json to:"
Write-Host "     $ConfigDir"
Write-Host "  2. Restart any open terminals"
Write-Host "  3. pip install -r requirements.txt"
Write-Host "  4. python main.py   (test run)"
Write-Host "  5. .\create_scheduled_task.ps1   (as Administrator)"
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""
