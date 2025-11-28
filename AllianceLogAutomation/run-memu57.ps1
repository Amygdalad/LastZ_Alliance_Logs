$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

.\memu-report-reader.ps1 `
    -Indexes 57 `
    -AfterLaunchDelaySeconds 60 `
    -TapPauseMs 5500 `
    -TapX 40 -TapY 900 `
    -AllianceTapX 500 -AllianceTapY 725 `
    -AllianceLogsTapX 400 -AllianceLogsTapY 725 `
    -AllianceToLogsDelaySeconds 5 `
    -PreAllianceTapCount 0 `
    -PreAllianceCloseCount 0 `
    -PreAllianceBackCount 0 `
    -ForceOcrForLogs `
    -UseOcrForAllianceButton `
    -TapWhileWaitingForAlliance `
    -PostLogsDelaySeconds 3 `
    -LeaveRunning `
    -Verbose
