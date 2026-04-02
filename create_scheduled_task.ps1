#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Creates a monthly Windows Task Scheduler task for the Azure VM Report.

.DESCRIPTION
    Registers a scheduled task that:
      - Runs on the 1st of every month at 06:00 AM (local server time)
      - Executes run_report.bat (which calls main.py and captures logs)
      - Runs under the SYSTEM account (has access to Machine env vars)
      - Retries once after 30 minutes if the run fails
      - Timeout after 3 hours (report should finish in < 30 min normally)

    Task visible in: Task Scheduler Library → Azure Monitoring

.NOTES
    Run as Administrator.
    Edit $ScriptsDir if your scripts are in a different location.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Configuration ──────────────────────────────────────────────────────
$TaskName    = "Azure VM Performance Report - Monthly"
$TaskFolder  = "Azure Monitoring"
$ScriptsDir  = "C:\VM-Reports\scripts"          # Adjust if needed
$BatchFile   = Join-Path $ScriptsDir "run_report.bat"
$Description = (
    "Generates and distributes the monthly Azure VM Performance Report. " +
    "Reads metrics via Azure Lighthouse (Read-Only). Sends via Microsoft Graph."
)

# Run on the 1st of every month at 06:00 local time
$RunHour    = 6
$RunMinute  = 0
$RunDay     = 1   # day of month

# ── Verify the batch file exists ──────────────────────────────────────
if (!(Test-Path $BatchFile)) {
    Write-Host "[ERROR] Batch file not found: $BatchFile" -ForegroundColor Red
    Write-Host "        Copy run_report.bat to $ScriptsDir and try again." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  Creating Azure VM Report Scheduled Task"         -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

# ── Create task folder if it doesn't exist ────────────────────────────
$scheduler = New-Object -ComObject "Schedule.Service"
$scheduler.Connect()
$rootFolder = $scheduler.GetFolder("\")
try {
    $rootFolder.GetFolder($TaskFolder) | Out-Null
    Write-Host "[OK]  Task folder '\$TaskFolder' already exists." -ForegroundColor Gray
} catch {
    $rootFolder.CreateFolder($TaskFolder) | Out-Null
    Write-Host "[CREATED] Task folder '\$TaskFolder'" -ForegroundColor Green
}

# ── Build trigger: monthly on day 1 at 06:00 ─────────────────────────
$startTime = (Get-Date -Day $RunDay -Hour $RunHour -Minute $RunMinute -Second 0)
$trigger   = New-ScheduledTaskTrigger `
    -Monthly `
    -DaysOfMonth $RunDay `
    -At $startTime

# ── Build action: cmd.exe /C run_report.bat ───────────────────────────
# Using cmd.exe so environment variables are loaded from the registry
$action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/C `"$BatchFile`"" `
    -WorkingDirectory $ScriptsDir

# ── Settings ─────────────────────────────────────────────────────────
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3) `
    -MultipleInstances IgnoreNew `
    -RestartCount 1 `
    -RestartInterval (New-TimeSpan -Minutes 30)

# ── Principal: SYSTEM account (reads Machine-scope env vars) ──────────
$principal = New-ScheduledTaskPrincipal `
    -UserId "NT AUTHORITY\SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

# ── Register the task ─────────────────────────────────────────────────
$taskPath = "\$TaskFolder\"

# Remove existing task with the same name if present
try {
    Unregister-ScheduledTask -TaskName $TaskName -TaskPath $taskPath -Confirm:$false
    Write-Host "[REMOVED] Existing task with same name (will re-create)" -ForegroundColor Yellow
} catch { }

Register-ScheduledTask `
    -TaskName   $TaskName `
    -TaskPath   $taskPath `
    -Trigger    $trigger `
    -Action     $action `
    -Settings   $settings `
    -Principal  $principal `
    -Description $Description `
    -Force | Out-Null

Write-Host "[OK] Task registered: \$TaskFolder\$TaskName" -ForegroundColor Green

# ── Verify registration ───────────────────────────────────────────────
$task = Get-ScheduledTask -TaskName $TaskName -TaskPath $taskPath
Write-Host ""
Write-Host "Task Summary:" -ForegroundColor White
Write-Host "  Name      : $($task.TaskName)"
Write-Host "  Path      : $($task.TaskPath)"
Write-Host "  Next run  : $(($task | Get-ScheduledTaskInfo).NextRunTime)"
Write-Host "  State     : $($task.State)"
Write-Host ""

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  Scheduled task created successfully."            -ForegroundColor Cyan
Write-Host ""
Write-Host "  Verify in: Task Scheduler → Task Scheduler Library"
Write-Host "             → Azure Monitoring → $TaskName"
Write-Host ""
Write-Host "  To run immediately for testing:"
Write-Host "  Start-ScheduledTask -TaskName '$TaskName' -TaskPath '$taskPath'"
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""
