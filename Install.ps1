<#
.SYNOPSIS
    Installs CleanPaste as a background startup task for the current user.

.DESCRIPTION
    - Copies CleanPaste.ps1 to a persistent location (~\.cleanpaste\)
    - Creates a Windows Scheduled Task that runs at user logon
    - Optionally starts the monitor immediately

.PARAMETER Uninstall
    Removes the scheduled task and installed files.
#>

param(
    [switch]$Uninstall
)

$installDir  = Join-Path $env:USERPROFILE ".cleanpaste"
$scriptDest  = Join-Path $installDir "CleanPaste.ps1"
$taskName    = "CleanPaste-ClipboardMonitor"
$scriptSrc   = Join-Path $PSScriptRoot "CleanPaste.ps1"

function Write-Banner {
    Write-Host ""
    Write-Host "  CleanPaste Installer" -ForegroundColor Cyan
    Write-Host "  ────────────────────" -ForegroundColor DarkCyan
    Write-Host ""
}

# ── Uninstall ────────────────────────────────────────────────────────────────
if ($Uninstall) {
    Write-Banner
    Write-Host "  Uninstalling CleanPaste..." -ForegroundColor Yellow

    # Remove scheduled task
    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "  ✓ Removed scheduled task" -ForegroundColor Green
    }

    # Stop any running instance
    $running = Get-Process -Name "powershell", "pwsh" -ErrorAction SilentlyContinue |
        Where-Object {
            try {
                $_.MainModule.FileName -and
                (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue).CommandLine -like "*CleanPaste*"
            } catch { $false }
        }
    foreach ($proc in $running) {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        Write-Host "  ✓ Stopped running instance (PID $($proc.Id))" -ForegroundColor Green
    }

    # Remove install directory
    if (Test-Path $installDir) {
        Remove-Item $installDir -Recurse -Force
        Write-Host "  ✓ Removed $installDir" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "  CleanPaste has been uninstalled." -ForegroundColor Green
    Write-Host ""
    return
}

# ── Install ──────────────────────────────────────────────────────────────────
Write-Banner

if (-not (Test-Path $scriptSrc)) {
    Write-Host "  ERROR: CleanPaste.ps1 not found in $PSScriptRoot" -ForegroundColor Red
    Write-Host "  Make sure Install.ps1 and CleanPaste.ps1 are in the same folder." -ForegroundColor Yellow
    return
}

# Copy files
Write-Host "  Installing to $installDir ..."
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Copy-Item $scriptSrc $scriptDest -Force
Write-Host "  ✓ Copied CleanPaste.ps1" -ForegroundColor Green

# Create scheduled task
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

$action  = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptDest`""

$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "CleanPaste clipboard monitor - auto-cleans terminal output on copy" `
    -Force | Out-Null

Write-Host "  ✓ Created startup task '$taskName'" -ForegroundColor Green

# Start immediately
Write-Host "  Starting CleanPaste now..."
Start-ScheduledTask -TaskName $taskName
Start-Sleep -Seconds 1

$taskInfo = Get-ScheduledTask -TaskName $taskName
if ($taskInfo.State -eq 'Running') {
    Write-Host "  ✓ CleanPaste is running!" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Task registered but may not have started. Check Task Scheduler." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  ┌─────────────────────────────────────────────────┐" -ForegroundColor Cyan
Write-Host "  │  CleanPaste is installed and will start          │" -ForegroundColor Cyan
Write-Host "  │  automatically when you log in.                  │" -ForegroundColor Cyan
Write-Host "  │                                                  │" -ForegroundColor Cyan
Write-Host "  │  To uninstall:                                   │" -ForegroundColor Cyan
Write-Host "  │    .\Install.ps1 -Uninstall                      │" -ForegroundColor Cyan
Write-Host "  └─────────────────────────────────────────────────┘" -ForegroundColor Cyan
Write-Host ""
