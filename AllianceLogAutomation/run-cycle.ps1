$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

while ($true) {
    Write-Host "=== Starting Cycle: 173 -> 58 -> 59 ===" -ForegroundColor Magenta

    # --- Instance 173 ---
    Write-Host "Running Instance 173..." -ForegroundColor Cyan
    try {
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
            -LeaveRunning:$false `
            -Verbose
    }
    catch {
        Write-Warning "Instance 173 failed: $_"
    }
    Start-Sleep -Seconds 5

    # --- Instance 58 ---
    Write-Host "Running Instance 58..." -ForegroundColor Cyan
    try {
        .\memu-report-reader.ps1 `
            -Indexes 58 `
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
            -LeaveRunning:$false `
            -Verbose
    }
    catch {
        Write-Warning "Instance 58 failed: $_"
    }
    Start-Sleep -Seconds 5

    # --- Instance 59 ---
    Write-Host "Running Instance 59..." -ForegroundColor Cyan
    try {
        .\memu-report-reader.ps1 `
            -Indexes 59 `
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
            -LeaveRunning:$false `
            -Verbose
    }
    catch {
        Write-Warning "Instance 59 failed: $_"
    }
    Start-Sleep -Seconds 5
}
