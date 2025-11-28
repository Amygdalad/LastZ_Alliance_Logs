$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

.\memu-report-reader.ps1 `
    -Indexes 173 `
    -AfterLaunchDelaySeconds 60 `
    -RestartAttempts 4 `
    -PostStartWaitSeconds 25 `
    -InitialStartWaitSeconds 45 `
    -ResetAdbBeforeStart `
    -ForceStopBeforeStart `
    -TapPauseMs 5500 `
    -TapX 40 -TapY 900 `
    -AllianceButtonTemplatePath 'templates\nonexistent.png' `
    -AllianceTapX 506 -AllianceTapY 750 `
    -SearchRegionFraction 0.3 `
    -MaxTapWaitSeconds 5 `
    -AllianceLogsTapX 400 -AllianceLogsTapY 750 `
    -AllianceToLogsDelaySeconds 5 `
    -LogsCropXRatio 0.0 `
    -LogsCropSlices 5 `
    -LogsCropSliceOverlapRatio 0.15 `
    -UseTesseract `
    -DismissBottomBrightX:$true `
    -PreAllianceTapCount 0 `
    -PreAllianceCloseCount 0 `
    -PreAllianceBackCount 0 `
    -ForceOcrForLogs `
    -UseOcrForAllianceButton `
    -TapWhileWaitingForAlliance `
    -PostLogsDelaySeconds 3 `
    -LeaveRunning `
    -Verbose
