<#
.SYNOPSIS
  Registers (or re-registers) the two ServiceStatusMonitor scheduled tasks:
    ServiceStatusWatchdog     - every 30 minutes, alert email only on changes
    ServiceStatusMorningDigest - daily 6:30 AM full status email

  Both run as the current user, interactive logon (required for Outlook COM),
  via the wscript.exe .vbs wrapper so no console window flashes.
  Re-run any time to update; safe to run repeatedly.
#>
$ErrorActionPreference = 'Stop'

$Root    = $PSScriptRoot
$Vbs     = Join-Path $Root 'Run-ServiceStatusMonitor.vbs'
$WScript = Join-Path $env:SystemRoot 'System32\wscript.exe'
$User    = "$env:USERDOMAIN\$env:USERNAME"

if (-not (Test-Path $Vbs)) { throw "Wrapper not found: $Vbs" }

$principal = New-ScheduledTaskPrincipal -UserId $User -LogonType Interactive -RunLevel Limited
$settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
             -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

# --- Watchdog: every 30 minutes, forever -------------------------------------
$watchStart   = (Get-Date).AddMinutes(2)
$watchTrigger = New-ScheduledTaskTrigger -Once -At $watchStart `
                -RepetitionInterval (New-TimeSpan -Minutes 30) `
                -RepetitionDuration (New-TimeSpan -Days 3650)
$watchAction  = New-ScheduledTaskAction -Execute $WScript -Argument ('"' + $Vbs + '" Watchdog')

Register-ScheduledTask -TaskName 'ServiceStatusWatchdog' `
    -Action $watchAction -Trigger $watchTrigger -Principal $principal -Settings $settings `
    -Description 'Polls vendor status feeds every 30 min; emails you only when a service goes DOWN or recovers. Script: ServiceStatusMonitor' `
    -Force | Out-Null
Write-Host 'Registered: ServiceStatusWatchdog (every 30 minutes)'

# --- Morning digest: daily 6:30 AM -------------------------------------------
$digestTrigger = New-ScheduledTaskTrigger -Daily -At '06:30'
$digestAction  = New-ScheduledTaskAction -Execute $WScript -Argument ('"' + $Vbs + '" Digest')

Register-ScheduledTask -TaskName 'ServiceStatusMorningDigest' `
    -Action $digestAction -Trigger $digestTrigger -Principal $principal -Settings $settings `
    -Description 'Daily 6:30 AM full service-status digest email. Script: ServiceStatusMonitor' `
    -Force | Out-Null
Write-Host 'Registered: ServiceStatusMorningDigest (daily 6:30 AM)'

Write-Host 'Done. Remove with: Unregister-ScheduledTask -TaskName ServiceStatusWatchdog, ServiceStatusMorningDigest'
