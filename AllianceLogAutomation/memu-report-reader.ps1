#requires -Version 5.1
<#
.SYNOPSIS
    Launch specific MEmu instances, open a "Z" app, capture screenshots, and OCR the text.

.DESCRIPTION
    Runs one emulator at a time (default indexes 59, 57, 173), wakes the device, launches a target
    package (explicitly provided or the last package that matches a Z-pattern), takes a screenshot,
    and extracts text using Windows' built-in OCR. Reports (screenshots + OCR text) are written
    under a local reports/ folder by default.

.PARAMETER Indexes
    MEmu instance indexes to process.

.PARAMETER InstanceRoot
    Path to the MEmu instance folders (contains MEmu_<n> subfolders).

.PARAMETER MEmuRoot
    Path to the MEmu install (must contain adb.exe and memuc.exe).

.PARAMETER ReportRoot
    Optional output root. Defaults to a "reports" subfolder next to this script.

.PARAMETER TargetPackage
    Package name to launch. If omitted, the script discovers the last installed package matching ZAppPattern.

.PARAMETER ZAppPattern
    Regex used to find the "Z" app when TargetPackage is not provided. Default matches packages starting with "com.z" or containing ".z".

.PARAMETER AfterLaunchDelaySeconds
    Delay after launching the app before the screenshot is taken.

.PARAMETER DeviceBootTimeoutSeconds
    How long to wait for adb to report "device" after starting the instance.

.PARAMETER LeaveRunning
    If set, the emulator is left running after the capture; otherwise it is stopped.

.PARAMETER DisableOcr
    Skip OCR if set; only screenshots are saved.
#>
[CmdletBinding()]
param(
    [int[]]$Indexes = @(59, 57, 173),
    [string]$InstanceRoot = 'Z:\Program Files\Microvirt\MEmu\MemuHyperv VMs',
    [string]$MEmuRoot = 'Z:\Program Files\Microvirt\MEmu',
    [string]$ReportRoot,
    [string]$TargetPackage,
    [string]$ZAppPattern = '[Zz]',
    [ValidateRange(1,300)]
    [int]$AfterLaunchDelaySeconds = 8,
    [ValidateRange(30,600)]
    [int]$DeviceBootTimeoutSeconds = 180,
    [ValidateRange(1,10)]
    [int]$RestartAttempts = 3,
    [ValidateRange(0,60)]
    [int]$PostStartWaitSeconds = 10,
    [ValidateRange(10,180)]
    [int]$InitialStartWaitSeconds = 45,
    [switch]$ResetAdbBeforeStart,
    [switch]$ForceStopBeforeStart,
    [ValidateRange(1,300)]
    [int]$TapIntervalSeconds = 5,
    [ValidateRange(1,900)]
    [int]$MaxTapWaitSeconds = 180,
    [ValidateRange(0.1,1.0)]
    [double]$SearchRegionFraction = 0.6,
    [ValidateRange(0,60000)]
    [int]$TapPauseMs = 700,
    [int]$TapX,
    [int]$TapY,
    [string]$MailIconTemplatePath,
    [ValidateRange(0,765)]
    [int]$ColorTolerance = 80,
    [ValidateRange(0.1,1.0)]
    [double]$MatchThreshold = 0.78,
    [ValidateRange(1,20)]
    [int]$TemplateSampleStep = 2,
    [switch]$WatchForMailIcon = $false,
    [switch]$WatchAllianceFlow = $true,
    [string]$AllianceButtonTemplatePath,
    [string]$AllianceLogsTemplatePath,
    [int]$AllianceTapX,
    [int]$AllianceTapY,
    [int]$AllianceLogsTapX,
    [int]$AllianceLogsTapY,
    [switch]$TapWhileWaitingForAlliance = $true,
    [int]$PreAllianceTapCount = 6,
    [int]$PreAllianceTapDelayMs = 400,
    [int]$PreAllianceCloseX = 220,
    [int]$PreAllianceCloseY = 700,
    [int]$PreAllianceCloseCount = 2,
    [int]$PreAllianceBackCount = 4,
    [int]$PreAllianceBackDelayMs = 300,
    [switch]$UseOcrForAllianceButton = $true,
    [switch]$ForceOcrForLogs,
    [switch]$DismissWhiteX = $true,
    [int]$BrightSpotMinSum = 720,
    [switch]$DismissBottomBrightX = $true,
    [double]$BottomBrightLeftRatio = 0.25,
    [double]$BottomBrightTopRatio = 0.85,
    [double]$BottomBrightRightRatio = 0.75,
    [double]$BottomBrightBottomRatio = 0.98,
    [int]$BottomBrightMinSum = 680,
    [switch]$TapBackOnEmptyLogs = $true,
    [double]$BackXRatio = 0.08,
    [double]$BackYRatio = 0.94,
    [ValidateRange(0.1,1.0)]
    [double]$AllianceLogsXRatio = 0.8,
    [ValidateRange(0.1,1.0)]
    [double]$AllianceLogsYRatio = 0.82,
    [ValidateRange(0,10)]
    [double]$AllianceToLogsDelaySeconds = 2,
    [ValidateRange(0,10)]
    [double]$PostLogsDelaySeconds = 2,
    [switch]$UseLogsCrop = $true,
    [ValidateRange(0.0,1.0)]
    [double]$LogsCropXRatio = 0.04,
    [ValidateRange(0.0,1.0)]
    [double]$LogsCropYRatio = 0.18,
    [ValidateRange(0.0,1.0)]
    [double]$LogsCropWidthRatio = 0.92,
    [ValidateRange(0.0,1.0)]
    [double]$LogsCropHeightRatio = 0.80,
    [string]$ExistingLogsImagePath,
    [switch]$UseLogsEnhance = $true,
    [ValidateRange(0,255)]
    [int]$LogsEnhanceThreshold = 170,
    [ValidateRange(0.1,3.0)]
    [double]$LogsEnhanceGain = 1.1,
    [ValidateRange(1,20)]
    [int]$LogsCropSlices = 4,
    [ValidateRange(0.0,0.9)]
    [double]$LogsCropSliceOverlapRatio = 0.10,
    [switch]$UseTesseract,
    [string]$TesseractPath = 'C:\Program Files\Tesseract-OCR\tesseract.exe',
    [switch]$UseOcrForAllianceLogs = $true,
    [switch]$EmitLogsCsv = $true,
    [switch]$LeaveRunning,
    [switch]$DisableOcr
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if (-not $ReportRoot) {
    $ReportRoot = Join-Path -Path $scriptRoot -ChildPath 'reports'
}
$screenshotRoot = Join-Path -Path $ReportRoot -ChildPath 'screenshots'
$ocrRoot = Join-Path -Path $ReportRoot -ChildPath 'ocr'
foreach ($dir in @($ReportRoot, $screenshotRoot, $ocrRoot)) {
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

$templatesRoot = Join-Path -Path $scriptRoot -ChildPath 'templates'

function Resolve-TemplatePath {
    param(
        [string]$ExplicitPath,
        [string[]]$CandidateRelativePaths,
        [string]$Label
    )

    if ($ExplicitPath) { return $ExplicitPath }
    foreach ($candidate in $CandidateRelativePaths) {
        $full = if ([System.IO.Path]::IsPathRooted($candidate)) { $candidate } else { Join-Path -Path $scriptRoot -ChildPath $candidate }
        if (Test-Path -LiteralPath $full) {
            return $full
        }
    }
    Write-Verbose ("No template resolved for {0}; candidates: {1}" -f $Label, ($CandidateRelativePaths -join ', '))
    return $null
}

$MailIconTemplatePath = Resolve-TemplatePath -ExplicitPath $MailIconTemplatePath -CandidateRelativePaths @('templates\mail-icon.png') -Label 'Mail'
$AllianceButtonTemplatePath = Resolve-TemplatePath -ExplicitPath $AllianceButtonTemplatePath -CandidateRelativePaths @('templates\alliance-button.png','templates\alliance.png','templates\alliance icon.png','templates\allaince icon.png','templates\alliance-btn.png') -Label 'Alliance button'
$AllianceLogsTemplatePath = Resolve-TemplatePath -ExplicitPath $AllianceLogsTemplatePath -CandidateRelativePaths @('templates\alliance-logs.png','templates\alliance logs.png','templates\alliance log.png','templates\alliance-logs-icon.png','templates\logs.png','templates\allaince logs.png','templates\allaince log.png') -Label 'Alliance logs'

$adbExe = Join-Path $MEmuRoot 'adb.exe'
$memucExe = Join-Path $MEmuRoot 'memuc.exe'
$memuConsoleExe = Join-Path $MEmuRoot 'MEmuConsole.exe'

function Assert-ToolPath {
    param(
        [string]$Path,
        [string]$FriendlyName
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$FriendlyName not found at $Path"
    }
}

Assert-ToolPath -Path $adbExe -FriendlyName 'adb.exe'
Assert-ToolPath -Path $memucExe -FriendlyName 'memuc.exe'
Assert-ToolPath -Path $memuConsoleExe -FriendlyName 'MEmuConsole.exe'

function Invoke-Memuc {
    param(
        [string[]]$Arguments,
        [string]$ErrorContext = 'memuc command failed'
    )

    Write-Verbose ("memuc {0}" -f ($Arguments -join ' '))
    $output = & $memucExe @Arguments 2>&1
    $exit = $LASTEXITCODE
    if ($exit -ne 0) {
        throw "{0} (exit {1})" -f $ErrorContext, $exit
    }
    return $output
}

function Get-AdbPort {
    param([string]$MemuFile)

    if (-not (Test-Path -LiteralPath $MemuFile)) {
        throw "Memu file not found: $MemuFile"
    }
    [xml]$doc = Get-Content -LiteralPath $MemuFile
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('m', 'http://www.memuhyperv.org/')
    $node = $doc.SelectSingleNode("//m:Forwarding[@name='ADB']", $ns)
    if (-not $node) {
        throw "ADB forwarding entry missing in $MemuFile"
    }
    return [int]$node.hostport
}

function Invoke-Adb {
    param(
        [string]$Serial,
        [string[]]$Arguments
    )

    Write-Verbose ("adb -s {0} {1}" -f $Serial, ($Arguments -join ' '))
    $output = & $adbExe '-s' $Serial @Arguments
    $exit = $LASTEXITCODE
    if ($exit -ne 0) {
        throw "adb failed (exit $exit) while running: $($Arguments -join ' ')"
    }
    return $output
}

function Get-AdbState {
    param([string]$Serial)

    try {
        $output = Invoke-Adb -Serial $Serial -Arguments @('get-state')
        $state = ($output | Select-Object -First 1).Trim().ToLowerInvariant()
        if (-not $state) { return 'unknown' }
        return $state
    }
    catch {
        return 'absent'
    }
}

function Wait-ForDevice {
    param(
        [string]$Serial,
        [int]$TimeoutSeconds = 120
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        if (Get-AdbState -Serial $Serial -eq 'device') {
            return $true
        }
        Start-Sleep -Seconds 3
    } while ((Get-Date) -lt $deadline)
    return $false
}

function Wait-ForDeviceOffline {
    param(
        [string]$Serial,
        [int]$TimeoutSeconds = 60
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        if (Get-AdbState -Serial $Serial -ne 'device') {
            return $true
        }
        Start-Sleep -Seconds 3
    } while ((Get-Date) -lt $deadline)
    return $false
}

function Start-MEmuInstance {
    param(
        [int]$Index,
        [string]$InstanceName,
        [string]$Serial,
        [int]$TimeoutSeconds = 120,
        [int]$RestartAttempts = 2,
        [int]$PostStartWaitSeconds = 5,
        [int]$InitialStartWaitSeconds = 45,
        [switch]$ResetAdbBeforeStart,
        [switch]$ForceStopBeforeStart
    )

    for ($attempt = 1; $attempt -le $RestartAttempts; $attempt++) {
        if ($ForceStopBeforeStart -and $attempt -eq 1) {
            Write-Host ("[{0}] Forcing stop for clean start." -f $InstanceName) -ForegroundColor DarkYellow
            Stop-MEmuInstance -Index $Index -Serial $Serial -InstanceName $InstanceName
        }

        if (Get-AdbState -Serial $Serial -eq 'device') {
            $pingOk = $false
            try {
                Invoke-Adb -Serial $Serial -Arguments @('shell', 'true') | Out-Null
                $pingOk = $true
            }
            catch {
                Write-Verbose ("{0} reported online but adb ping failed; restarting instance." -f $InstanceName)
            }
            if ($pingOk) {
                Write-Verbose ("{0} already online at {1}; skipping start." -f $InstanceName, $Serial)
                return
            }
        }

        if ($ResetAdbBeforeStart -and $attempt -eq 1) {
            Write-Host ("[{0}] Resetting adb server before start." -f $InstanceName) -ForegroundColor DarkGray
            try { & $adbExe 'kill-server' | Out-Null } catch { Write-Verbose ("adb kill-server failed: {0}" -f $_.Exception.Message) }
            try { & $adbExe 'start-server' | Out-Null } catch { Write-Verbose ("adb start-server failed: {0}" -f $_.Exception.Message) }
        }

        if ($attempt -gt 1) {
            Write-Host ("[{0}] Pre-stop for clean start (attempt {1}/{2})." -f $InstanceName, $attempt, $RestartAttempts) -ForegroundColor DarkYellow
            Stop-MEmuInstance -Index $Index -Serial $Serial -InstanceName $InstanceName
        }

        Write-Host ("[{0}] Starting via MEmuConsole (attempt {1}/{2})" -f $InstanceName, $attempt, $RestartAttempts) -ForegroundColor Cyan
        try {
            Start-Process -FilePath $memuConsoleExe -ArgumentList $InstanceName -WorkingDirectory $MEmuRoot | Out-Null
        }
        catch {
            Write-Warning ("MEmuConsole start failed for {0}: {1}" -f $InstanceName, $_.Exception.Message)
        }

        try { & $adbExe 'connect' $Serial | Out-Null } catch { Write-Verbose ("adb connect attempt failed for {0}: {1}" -f $Serial, $_.Exception.Message) }

        if ($PostStartWaitSeconds -gt 0) { Start-Sleep -Seconds $PostStartWaitSeconds }

        $initialWait = [Math]::Min($TimeoutSeconds, $InitialStartWaitSeconds)
        $finalWait = [Math]::Max(0, $TimeoutSeconds - $initialWait)

        Write-Host ("[{0}] Waiting up to {1}s for adb device {2} after start (attempt {3}/{4})." -f $InstanceName, $initialWait, $Serial, $attempt, $RestartAttempts) -ForegroundColor DarkGray
        if (Wait-ForDevice -Serial $Serial -TimeoutSeconds $initialWait) {
            if (Wait-ForBootCompletion -Serial $Serial) {
                return
            }
            Write-Warning ("{0} device found but boot completion timed out." -f $InstanceName)
        }
        
        Write-Warning ("{0} still offline or incomplete boot after start attempt {1}/{2}." -f $InstanceName, $attempt, $RestartAttempts)
        try {
            Start-Process -FilePath $memuConsoleExe -ArgumentList $InstanceName -WorkingDirectory $MEmuRoot | Out-Null
        }
        catch {
            Write-Warning ("Fallback start via MEmuConsole failed for {0}: {1}" -f $InstanceName, $_.Exception.Message)
        }

        if ($PostStartWaitSeconds -gt 0) { Start-Sleep -Seconds $PostStartWaitSeconds }

        if ($finalWait -gt 0) {
            Write-Host ("[{0}] Waiting up to {1}s for adb device {2} after console launch (attempt {3}/{4})." -f $InstanceName, $finalWait, $Serial, $attempt, $RestartAttempts) -ForegroundColor DarkGray
            if (Wait-ForDevice -Serial $Serial -TimeoutSeconds $finalWait) {
                if (Wait-ForBootCompletion -Serial $Serial) {
                    return
                }
                Write-Warning ("{0} device found but boot completion timed out (fallback)." -f $InstanceName)
            }
        }

        Write-Warning ("{0} not online after attempt {1}/{2}; will retry." -f $InstanceName, $attempt, $RestartAttempts)
    }

    throw "Timed out waiting for $InstanceName ($Serial) after $RestartAttempts attempts."
}

function Stop-MEmuInstance {
    param(
        [int]$Index,
        [string]$Serial,
        [string]$InstanceName
    )

    # Try graceful ADB shutdown first if possible
    if ($Serial -and (Get-AdbState -Serial $Serial) -eq 'device') {
        try {
            Write-Verbose "Attempting adb emu kill for $InstanceName..."
            & $adbExe '-s' $Serial 'emu' 'kill' | Out-Null
            # Give it a moment to shut down
            Start-Sleep -Seconds 2
        }
        catch {
            Write-Verbose ("adb emu kill failed for {0}: {1}" -f $InstanceName, $_.Exception.Message)
        }
    }

    Write-Host ("Stopping MEmu instance {0}..." -f $Index) -ForegroundColor DarkYellow
    
    # Force kill MEmuHeadless/MEmu processes for this instance if possible, or just rely on memuc
    # But first, try memuc stop
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $memucExe
        $psi.Arguments = "stop -i $Index"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        
        $proc = [System.Diagnostics.Process]::Start($psi)
        if (-not $proc.WaitForExit(15000)) { # 15 second timeout
            Write-Warning ("memuc stop -i {0} timed out." -f $Index)
        }
    }
    catch {
        Write-Warning ("memuc stop failed: {0}" -f $_.Exception.Message)
    }
    
    # Double check and force kill via taskkill if it's still running (by window title or similar is hard, but we can try memuc again or just wait)
    Start-Sleep -Seconds 2

    if ($Serial) {
        [void](Wait-ForDeviceOffline -Serial $Serial -TimeoutSeconds 10)
    }
}

function Unlock-Device {
    param([string]$Serial)

    try {
        Invoke-Adb -Serial $Serial -Arguments @('shell', 'input', 'keyevent', '224') | Out-Null # wake
        Start-Sleep -Milliseconds 300
        Invoke-Adb -Serial $Serial -Arguments @('shell', 'input', 'keyevent', '82') | Out-Null  # unlock/menu
        Start-Sleep -Milliseconds 300
        Invoke-Adb -Serial $Serial -Arguments @('shell', 'input', 'keyevent', '3') | Out-Null   # home
    }
    catch {
        Write-Warning ("Unable to send wake/unlock events: {0}" -f $_.Exception.Message)
    }
}

function Get-ZAppPackage {
    param([string]$Serial)

    if ($TargetPackage) { return $TargetPackage }

    $packageLists = @()
    $thirdParty = @()
    
    # Retry loop for package listing (sometimes fails immediately after boot)
    for ($i = 1; $i -le 5; $i++) {
        try {
            $thirdParty = @(Invoke-Adb -Serial $Serial -Arguments @('shell', 'pm', 'list', 'packages', '-3'))
            $packageLists += $thirdParty
            
            $allPkgs = @(Invoke-Adb -Serial $Serial -Arguments @('shell', 'pm', 'list', 'packages'))
            $packageLists += $allPkgs
            
            if ($packageLists.Count -gt 0) {
                break
            }
        }
        catch {
            Write-Verbose ("Package list attempt $i failed: {0}" -f $_.Exception.Message)
        }
        Start-Sleep -Seconds 2
    }

    $thirdPartyClean = @(
        $thirdParty |
        ForEach-Object { ($_ -replace '^package:', '').Trim() } |
        Where-Object { $_ } |
        Sort-Object -Unique
    )

    $candidates = @(
        $packageLists |
        Where-Object { $_ } |
        ForEach-Object { $_ } |
        ForEach-Object { ($_ -replace '^package:', '').Trim() } |
        Where-Object { $_ } |
        Sort-Object -Unique
    )

    if (-not $candidates -or $candidates.Count -eq 0) {
        return $null
    }

    if ($thirdPartyClean -and $thirdPartyClean.Count -gt 0) {
        if ($ZAppPattern) {
            $tpMatches = $thirdPartyClean | Where-Object { $_ -match $ZAppPattern }
            if ($tpMatches -and $tpMatches.Count -gt 0) {
                return ($tpMatches | Select-Object -Last 1)
            }
        }
        # No Z-pattern match among third-party: use last third-party
        return ($thirdPartyClean | Select-Object -Last 1)
    }

    if ($ZAppPattern) {
        $allMatches = $candidates | Where-Object { $_ -match $ZAppPattern }
        if ($allMatches -and $allMatches.Count -gt 0) {
            return ($allMatches | Select-Object -Last 1)
        }
    }

    # Fallback: last package alphabetically
    return ($candidates | Select-Object -Last 1)
}

function Wait-ForBootCompletion {
    param(
        [string]$Serial,
        [int]$TimeoutSeconds = 120
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $boot = Invoke-Adb -Serial $Serial -Arguments @('shell', 'getprop', 'sys.boot_completed')
            if ($boot -and $boot.Trim() -eq '1') {
                Write-Verbose "Device $Serial boot completed."
                return $true
            }
        } catch {}
        Start-Sleep -Seconds 2
    }
    Write-Warning "Timeout waiting for boot completion on $Serial."
    return $false
}

function Wait-ForAppFocus {
    param(
        [string]$Serial,
        [string]$Package,
        [int]$TimeoutSeconds = 60
    )
    
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            # Check mCurrentFocus or mFocusedApp
            $output = Invoke-Adb -Serial $Serial -Arguments @('shell', 'dumpsys', 'window', 'windows')
            # Join output to a single string for easier matching
            $text = $output -join "`n"
            
            if ($text -match "mCurrentFocus=.*$Package" -or $text -match "mFocusedApp=.*$Package") {
                Write-Verbose "App $Package is in focus."
                return $true
            }
        }
        catch {
            Write-Verbose "Error checking app focus: $_"
        }
        Start-Sleep -Seconds 2
    }
    
    Write-Warning "Timeout waiting for app $Package to be in focus."
    return $false
}

function Launch-App {
    param(
        [string]$Serial,
        [string]$Package
    )

    if (-not $Package) {
        Write-Warning "No matching Z app found; skipping app launch."
        return
    }

    $args = @('shell', 'monkey', '-p', $Package, '-c', 'android.intent.category.LAUNCHER', '1')
    try {
        Invoke-Adb -Serial $Serial -Arguments $args | Out-Null
        Write-Verbose ("Launched package {0}" -f $Package)
    }
    catch {
        Write-Warning ("Failed to launch {0}: {1}" -f $Package, $_.Exception.Message)
    }
}

function Capture-Screenshot {
    param(
        [string]$Serial,
        [string]$Path
    )

    $cmd = '"{0}" -s "{1}" exec-out "screencap -p" > "{2}"' -f $adbExe, $Serial, $Path
    cmd.exe /c $cmd | Out-Null
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Screenshot failed; file not created at $Path"
    }
    return $Path
}

function Get-ScreenSize {
    param([string]$Serial)

    try {
        $output = Invoke-Adb -Serial $Serial -Arguments @('shell', 'wm', 'size')
        $line = ($output | Where-Object { $_ -match 'Physical size' } | Select-Object -First 1)
        if (-not $line) { $line = $output | Select-Object -First 1 }
        if ($line -match '(?<w>\d+)\s*x\s*(?<h>\d+)') {
            return [pscustomobject]@{
                Width  = [int]$Matches.w
                Height = [int]$Matches.h
            }
        }
    }
    catch {
        Write-Verbose ("wm size failed: {0}" -f $_.Exception.Message)
    }
    return $null
}

function Invoke-AdbTap {
    param(
        [string]$Serial,
        [int]$X,
        [int]$Y
    )
    try {
        Invoke-Adb -Serial $Serial -Arguments @('shell', 'input', 'tap', $X, $Y) | Out-Null
        Write-Verbose ("Tapped {0} at {1},{2}" -f $Serial, $X, $Y)
        if ($TapPauseMs -gt 0) {
            Start-Sleep -Milliseconds $TapPauseMs
        }
    }
    catch {
        Write-Warning ("tap failed at {0},{1}: {2}" -f $X, $Y, $_.Exception.Message)
    }
}

function Get-ColorDistance {
    param(
        [System.Drawing.Color]$A,
        [System.Drawing.Color]$B
    )
    $dr = [int]$A.R - [int]$B.R
    $dg = [int]$A.G - [int]$B.G
    $db = [int]$A.B - [int]$B.B
    return [Math]::Abs($dr) + [Math]::Abs($dg) + [Math]::Abs($db)
}

function Find-TemplateMatch {
    param(
        [string]$ImagePath,
        [string]$TemplatePath,
        [double]$RegionFraction = 0.6,
        [int]$Tolerance = 28,
        [double]$Threshold = 0.78,
        [int]$SampleStep = 2
    )

    Add-Type -AssemblyName System.Drawing

    if (-not (Test-Path -LiteralPath $ImagePath)) {
        throw "Image not found: $ImagePath"
    }
    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        throw "Template not found: $TemplatePath"
    }

    $img = [System.Drawing.Bitmap]::FromFile($ImagePath)
    $tpl = [System.Drawing.Bitmap]::FromFile($TemplatePath)
    try {
        if ($tpl.Width -ge $img.Width -or $tpl.Height -ge $img.Height) {
            return [pscustomobject]@{ Found = $false }
        }

        $startX = [Math]::Max([int]($img.Width * (1 - $RegionFraction)), 0)
        $startY = [Math]::Max([int]($img.Height * (1 - $RegionFraction)), 0)
        $endX = $img.Width - $tpl.Width
        $endY = $img.Height - $tpl.Height
        $step = [Math]::Max(1, [int]$SampleStep)

        for ($y = $startY; $y -le $endY; $y += $step) {
            for ($x = $startX; $x -le $endX; $x += $step) {
                $matched = 0
                $total = 0
                for ($ty = 0; $ty -lt $tpl.Height; $ty += $step) {
                    for ($tx = 0; $tx -lt $tpl.Width; $tx += $step) {
                        $cTpl = $tpl.GetPixel($tx, $ty)
                        $cImg = $img.GetPixel($x + $tx, $y + $ty)
                        $total++
                        if (Get-ColorDistance -A $cTpl -B $cImg -le $Tolerance) {
                            $matched++
                        }
                    }
                }
                if ($total -gt 0) {
                    $ratio = $matched / $total
                    if ($ratio -ge $Threshold) {
                        $centerX = [int]($x + ($tpl.Width / 2))
                        $centerY = [int]($y + ($tpl.Height / 2))
                        return [pscustomobject]@{
                            Found = $true
                            X     = $centerX
                            Y     = $centerY
                            Score = $ratio
                        }
                    }
                }
            }
        }
        return [pscustomobject]@{ Found = $false }
    }
    finally {
        $img.Dispose()
        $tpl.Dispose()
    }
}

function Find-BrightestPixel {
    param(
        [string]$ImagePath,
        [ValidateRange(0.0,1.0)][double]$LeftRatio = 0.0,
        [ValidateRange(0.0,1.0)][double]$TopRatio = 0.0,
        [ValidateRange(0.0,1.0)][double]$RightRatio = 1.0,
        [ValidateRange(0.0,1.0)][double]$BottomRatio = 1.0,
        [int]$MinX = 0
    )

    Add-Type -AssemblyName System.Drawing

    if (-not (Test-Path -LiteralPath $ImagePath)) {
        throw "Image not found: $ImagePath"
    }

    $bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
    try {
        $x0 = [int]([Math]::Max(0, $bmp.Width * $LeftRatio))
        $y0 = [int]([Math]::Max(0, $bmp.Height * $TopRatio))
        $x1 = [int]([Math]::Min($bmp.Width, $bmp.Width * $RightRatio))
        $y1 = [int]([Math]::Min($bmp.Height, $bmp.Height * $BottomRatio))

        # Enforce MinX
        if ($x0 -lt $MinX) { $x0 = $MinX }

        if ($x1 -le $x0 -or $y1 -le $y0) {
            return [pscustomobject]@{ Found = $false }
        }

        $maxSum = -1
        $maxX = 0
        $maxY = 0
        
        # Iterate to find the brightest pixel.
        # We use -ge to prefer the bottom-right-most pixel in case of ties (helps with bottom buttons vs text)
        for ($y = $y0; $y -lt $y1; $y++) {
            for ($x = $x0; $x -lt $x1; $x++) {
                $c = $bmp.GetPixel($x, $y)
                $sum = [int]$c.R + [int]$c.G + [int]$c.B
                if ($sum -ge $maxSum) {
                    $maxSum = $sum
                    $maxX = $x
                    $maxY = $y
                }
            }
        }

        return [pscustomobject]@{
            Found      = ($maxSum -ge 0)
            X          = $maxX
            Y          = $maxY
            Brightness = $maxSum
        }
    }
    finally {
        $bmp.Dispose()
    }
}

function Wait-ForIcon {
    param(
        [string]$Serial,
        [string]$TemplatePath,
        [int]$TapInterval,
        [int]$MaxWait,
        [double]$RegionFraction,
        [int]$TapXOverride,
        [int]$TapYOverride,
        [int]$Tolerance,
        [double]$Threshold,
        [int]$SampleStep,
        [string]$Label = 'Icon',
        [switch]$TapEachInterval = $true,
        [switch]$TapOnDetection = $false
    )

    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        Write-Warning ("{0} template missing at {1}; skipping watch loop." -f $Label, $TemplatePath)
        return $null
    }

    $size = Get-ScreenSize -Serial $Serial
    $tapXCoord = if ($TapXOverride) { $TapXOverride } else { 30 }
    $tapYCoord = if ($TapYOverride) { $TapYOverride } elseif ($size) { [Math]::Max($size.Height - 30, 5) } else { 700 }

    $deadline = (Get-Date).AddSeconds($MaxWait)
    $lastShot = $null
    $safeLabel = ($Label -replace '[^a-zA-Z0-9_-]', '-')

    while ((Get-Date) -lt $deadline) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmssfff'
        $shotPath = Join-Path $screenshotRoot ("memu-watch-{0}-{1}-{2}.png" -f ($Serial -replace '[:\.]', '_'), $safeLabel, $stamp)
        $lastShot = $shotPath
        Capture-Screenshot -Serial $Serial -Path $shotPath | Out-Null

        try {
            $match = Find-TemplateMatch -ImagePath $shotPath -TemplatePath $TemplatePath -RegionFraction $RegionFraction -Tolerance $Tolerance -Threshold $Threshold -SampleStep $SampleStep
            if ($match.Found) {
                Write-Host ("{0} detected; screenshot saved to {1}" -f $Label, $shotPath) -ForegroundColor Green
                if ($TapOnDetection) {
                    Write-Host ("Tapping {0} at {1},{2}" -f $Label, $match.X, $match.Y) -ForegroundColor DarkCyan
                    Invoke-AdbTap -Serial $Serial -X $match.X -Y $match.Y
                }
                return @{
                    Found      = $true
                    Screenshot = $shotPath
                    X          = $match.X
                    Y          = $match.Y
                    Score      = $match.Score
                }
            }
        }
        catch {
            Write-Warning ("{0} detection failed on {1}: {2}" -f $Label, $shotPath, $_.Exception.Message)
        }

        if ($TapEachInterval) {
            Invoke-AdbTap -Serial $Serial -X $tapXCoord -Y $tapYCoord
        }
        Start-Sleep -Seconds $TapInterval
    }

    Write-Warning ("{0} not detected before timeout." -f $Label)
    return @{
        Found      = $false
        Screenshot = $lastShot
    }
}

function Get-AsyncResult {
    param(
        $AsyncOperation,
        [Type]$ResultType
    )

    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 } |
        Select-Object -First 1
    if (-not $method) {
        throw "AsTask extension method not found; cannot bridge WinRT async."
    }
    $generic = $method.MakeGenericMethod($ResultType)
    $task = $generic.Invoke($null, @($AsyncOperation))
    $task.Wait()
    return $task.Result
}

function Get-OcrResult {
    param([string]$ImagePath)

    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    # Touch WinRT types to ensure they are loaded
    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]

    $file = Get-AsyncResult -AsyncOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)) -ResultType ([Windows.Storage.StorageFile])
    $stream = Get-AsyncResult -AsyncOperation ($file.OpenReadAsync()) -ResultType ([Windows.Storage.Streams.IRandomAccessStreamWithContentType])
    $decoder = Get-AsyncResult -AsyncOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) -ResultType ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Get-AsyncResult -AsyncOperation ($decoder.GetSoftwareBitmapAsync()) -ResultType ([Windows.Graphics.Imaging.SoftwareBitmap])

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if (-not $engine) {
        throw "Windows OCR engine not available on this machine."
    }
    return Get-AsyncResult -AsyncOperation ($engine.RecognizeAsync($bitmap)) -ResultType ([Windows.Media.Ocr.OcrResult])
}

function Invoke-Ocr {
    param([string]$ImagePath)
    $result = Get-OcrResult -ImagePath $ImagePath
    return $result.Text
}

function Find-OcrPhraseCenter {
    param(
        [string]$ImagePath,
        [string]$Phrase
    )

    $result = Get-OcrResult -ImagePath $ImagePath
    if (-not $result) { return [pscustomobject]@{ Found = $false } }

    $phraseWords = @(
        ($Phrase -split '\s+') |
        ForEach-Object { $_.Trim().ToLowerInvariant() } |
        Where-Object { $_ }
    )
    if (-not $phraseWords -or $phraseWords.Count -eq 0) {
        return [pscustomobject]@{ Found = $false }
    }

    foreach ($line in $result.Lines) {
        $lineWords = @()
        foreach ($word in $line.Words) {
            $text = "$($word.Text)".Trim().ToLowerInvariant()
            if (-not $text) { continue }
            $rect = $word.BoundingRect
            $lineWords += [pscustomobject]@{
                Text  = $text
                Rect  = $rect
                X     = [double]$rect.X
                Y     = [double]$rect.Y
                X2    = [double]($rect.X + $rect.Width)
                Y2    = [double]($rect.Y + $rect.Height)
            }
        }

        for ($i = 0; $i -le $lineWords.Count - $phraseWords.Count; $i++) {
            $allMatch = $true
            for ($j = 0; $j -lt $phraseWords.Count; $j++) {
                if ($lineWords[$i + $j].Text -ne $phraseWords[$j]) {
                    $allMatch = $false
                    break
                }
            }
            if ($allMatch) {
                $x1 = [double]::MaxValue
                $y1 = [double]::MaxValue
                $x2 = 0.0
                $y2 = 0.0
                for ($k = 0; $k -lt $phraseWords.Count; $k++) {
                    $w = $lineWords[$i + $k]
                    $x1 = [Math]::Min($x1, $w.X)
                    $y1 = [Math]::Min($y1, $w.Y)
                    $x2 = [Math]::Max($x2, $w.X2)
                    $y2 = [Math]::Max($y2, $w.Y2)
                }
                $cx = [int](($x1 + $x2) / 2)
                $cy = [int](($y1 + $y2) / 2)
                return [pscustomobject]@{
                    Found = $true
                    X     = $cx
                    Y     = $cy
                    Words = ($lineWords[$i..($i + $phraseWords.Count - 1)])
                }
            }
        }
    }

    return [pscustomobject]@{ Found = $false }
}

function Crop-ImageFraction {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [double]$XRatio,
        [double]$YRatio,
        [double]$WidthRatio,
        [double]$HeightRatio
    )

    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($InputPath)
    try {
        $x = [int]([double]$bmp.Width * $XRatio)
        $y = [int]([double]$bmp.Height * $YRatio)
        $w = [int]([double]$bmp.Width * $WidthRatio)
        $h = [int]([double]$bmp.Height * $HeightRatio)
        $w = [Math]::Min($w, $bmp.Width - $x)
        $h = [Math]::Min($h, $bmp.Height - $y)
        $rect = New-Object System.Drawing.Rectangle $x, $y, $w, $h
        $crop = $bmp.Clone($rect, $bmp.PixelFormat)
        $crop.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $crop.Dispose()
    }
    finally {
        $bmp.Dispose()
    }
    return $OutputPath
}
function Crop-ImageSlices {
    param(
        [string]$InputPath,
        [string]$OutputPrefix,
        [int]$Slices,
        [double]$OverlapRatio
    )

    Add-Type -AssemblyName System.Drawing
    $paths = New-Object System.Collections.Generic.List[string]
    $bmp = [System.Drawing.Bitmap]::FromFile($InputPath)
    try {
        $h = $bmp.Height
        $sliceHeight = [int]([Math]::Max(10, $h / [Math]::Max(1,$Slices)))
        $overlap = [int]($sliceHeight * [Math]::Min([Math]::Max($OverlapRatio,0.0),0.9))
        $y = 0
        $idx = 0
        while ($y -lt $h) {
            $height = [Math]::Min($sliceHeight, $h - $y)
            $rect = New-Object System.Drawing.Rectangle 0, $y, $bmp.Width, $height
            $crop = $bmp.Clone($rect, $bmp.PixelFormat)
            $out = "{0}-slice{1}.png" -f $OutputPrefix, $idx
            $crop.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
            $crop.Dispose()
            $paths.Add($out) | Out-Null
            $idx++
            if ($height -eq 0) { break }
            $y += ($sliceHeight - $overlap)
        }
    }
    finally {
        $bmp.Dispose()
    }
    return $paths
}

function Scale-Image {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [double]$ScaleFactor = 2.0
    )

    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($InputPath)
    try {
        $newW = [int]([Math]::Max(1, $bmp.Width * $ScaleFactor))
        $newH = [int]([Math]::Max(1, $bmp.Height * $ScaleFactor))
        $scaled = New-Object System.Drawing.Bitmap $newW, $newH
        $g = [System.Drawing.Graphics]::FromImage($scaled)
        try {
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
            $g.DrawImage($bmp, 0, 0, $newW, $newH)
            $scaled.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $g.Dispose()
            $scaled.Dispose()
        }
    }
    finally {
        $bmp.Dispose()
    }
    return $OutputPath
}

function Enhance-ImageForOcr {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [int]$Threshold,
        [double]$Gain
    )

    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($InputPath)
    try {
        for ($y = 0; $y -lt $bmp.Height; $y++) {
            for ($x = 0; $x -lt $bmp.Width; $x++) {
                $p = $bmp.GetPixel($x, $y)
                $g = [int](($p.R * 0.299) + ($p.G * 0.587) + ($p.B * 0.114))
                $g = [int]([Math]::Min(255, [Math]::Max(0, $g * $Gain)))
                $val = if ($g -ge $Threshold) { 255 } else { 0 }
                $bmp.SetPixel($x, $y, [System.Drawing.Color]::FromArgb($val, $val, $val))
            }
        }
        $bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bmp.Dispose()
    }
    return $OutputPath
}

function Invoke-TesseractOcr {
    param(
        [string]$ImagePath,
        [string]$TesseractExe
    )

    if (-not (Test-Path -LiteralPath $TesseractExe)) {
        throw "tesseract.exe not found at $TesseractExe"
    }
    $tmpBase = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), $null)
    $tmpTxt = "${tmpBase}.txt"
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $TesseractExe
    $psi.Arguments = "`"$ImagePath`" `"$tmpBase`" -l eng --psm 6"
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.WaitForExit()
    if ($proc.ExitCode -ne 0) {
        $err = $proc.StandardError.ReadToEnd()
        throw "tesseract failed (exit $($proc.ExitCode)): $err"
    }
    $text = ''
    if (Test-Path -LiteralPath $tmpTxt) {
        $text = Get-Content -LiteralPath $tmpTxt -Raw
        Remove-Item -LiteralPath $tmpTxt -Force -ErrorAction SilentlyContinue
    }
    return $text
}

function Invoke-AnyOcr {
    param([string]$ImagePath)

    if ($UseTesseract) {
        try {
            return Invoke-TesseractOcr -ImagePath $ImagePath -TesseractExe $TesseractPath
        }
        catch {
            Write-Warning ("Tesseract OCR failed: {0}; falling back to Windows OCR." -f $_.Exception.Message)
        }
    }
    return Invoke-Ocr -ImagePath $ImagePath
}

function Parse-AttackLogs {
    param(
        [string]$Text,
        [string]$InitialTime = ''
    )

    $clean = -join ($Text.ToCharArray() | ForEach-Object { if ([char]::IsControl($_)) { ' ' } else { $_ } })
    $clean = $clean -replace '[^\u0020-\u007E]', ' '
    $clean = $clean -replace '\s+', ' '

    $logDate = ''
    $dateMatch = [regex]::Match($clean, '\d{4}-\d{2}-\d{2}')
    if ($dateMatch.Success) { $logDate = $dateMatch.Value }

    $entries = New-Object System.Collections.Generic.List[pscustomobject]

    $attackPattern = '(?:[\u0000-\u001f·•◆▶▷►]*\s*)?(?:(?<time>\d{2}:\d{2}:\s*\d{2})\s*)?(?:Logs\s+)?(?<farm>[A-Za-z0-9\s]+?)\s+was attacked by\s+(?<attacker>[A-Za-z0-9\s]+?)\s*\(X[:：]?[^\d]{0,3}(?<x>\d+)\s*Y[:：]?[^\d]{0,3}(?<y>\d+)\)'
    $matches = [regex]::Matches($clean, $attackPattern, 'IgnoreCase')
    
    $lastTime = $InitialTime
    
    foreach ($m in $matches) {
        $timeVal = $m.Groups['time'].Value
        if ($timeVal) { 
            $timeVal = $timeVal -replace '\s+', '' 
            $lastTime = $timeVal
        } elseif ($lastTime) {
            $timeVal = $lastTime
        }

        $finalTime = $timeVal
        if ($logDate) {
            if ($finalTime) { $finalTime = "$logDate $finalTime" } else { $finalTime = $logDate }
        }
        $entries.Add([pscustomobject]@{
            Time     = $finalTime
            Farm     = $m.Groups['farm'].Value.Trim()
            Attacker = $m.Groups['attacker'].Value.Trim()
            X        = $m.Groups['x'].Value.Trim()
            Y        = $m.Groups['y'].Value.Trim()
            CaptureDate = (Get-Date -Format 'yyyy-MM-dd')
        }) | Out-Null
    }

    return [pscustomobject]@{
        Records = $entries
        LastTime = $lastTime
    }
}

function Normalize-AttackRecords {
    param([System.Collections.IEnumerable]$Records)

    $normalized = New-Object System.Collections.Generic.List[pscustomobject]
    foreach ($r in $Records) {
        $farm = ($r.Farm -replace '[^A-Za-z0-9\s]', ' ') -replace '\s+', ' '
        $attacker = ($r.Attacker -replace '[^A-Za-z0-9\s]', ' ') -replace '\s+', ' '

        $normalized.Add([pscustomobject]@{
            Time     = "$($r.Time)".Trim()
            Farm     = $farm.Trim()
            Attacker = $attacker.Trim()
            X        = "$($r.X)".Trim()
            Y        = "$($r.Y)".Trim()
            CaptureDate = "$($r.CaptureDate)".Trim()
        }) | Out-Null
    }

    return $normalized
}

function Collapse-AttackRecords {
    param([System.Collections.IEnumerable]$Records)

    $byFarmCoord = $Records | Group-Object -Property Farm,X,Y
    $result = New-Object System.Collections.Generic.List[pscustomobject]
    foreach ($g in $byFarmCoord) {
        $rows = $g.Group
        $timed = @($rows | Where-Object { $_.Time -match '\d{2}:\d{2}:\d{2}' })
        if ($timed.Count -gt 0) {
            $rowsToUse = $timed | Sort-Object {
                try {
                    return [DateTime]::Parse($_.Time)
                }
                catch {
                    return [DateTime]::MinValue
                }
            } -Descending
            $result.Add(($rowsToUse | Select-Object -First 1)) | Out-Null
        }
        else {
            $result.Add(($rows | Select-Object -First 1)) | Out-Null
        }
    }

    $result | Sort-Object @{
        Expression={
            $dt=[DateTime]::MinValue
            if ([DateTime]::TryParse($_.Time,[ref]$dt)) { return $dt } else { return [DateTime]::MinValue }
        }; Descending=$true
    }, @{Expression={$_.Farm}} | ForEach-Object { $_ }
}

function Get-AttackRecordDate {
    param([psobject]$Record)

    if (-not $Record) { return $null }
    $dt = [DateTime]::MinValue
    if ($Record.Time -and [DateTime]::TryParse("$($Record.Time)", [ref]$dt)) {
        return $dt.Date
    }
    $m = [regex]::Match("$($Record.Time)", '\d{4}-\d{2}-\d{2}')
    if ($m.Success -and [DateTime]::TryParse($m.Value, [ref]$dt)) {
        return $dt.Date
    }
    return $null
}

function Update-DailyAttackLog {
    param(
        [string]$OutputRoot,
        [System.Collections.IEnumerable]$Records
    )

    if (-not $Records) { return }

    $byDate = @{}
    foreach ($r in $Records) {
        $d = Get-AttackRecordDate -Record $r
        if (-not $d) { continue }
        $k = $d.ToString('yyyyMMdd')
        if (-not $byDate.ContainsKey($k)) {
            $byDate[$k] = New-Object System.Collections.Generic.List[pscustomobject]
        }
        $byDate[$k].Add($r) | Out-Null
    }

    foreach ($k in $byDate.Keys) {
        $dailyPath = Join-Path $OutputRoot ("memu-logs-daily-{0}.csv" -f $k)
        $existing = @()
        if (Test-Path -LiteralPath $dailyPath) {
            try { $existing = @(Import-Csv -Path $dailyPath) } catch { $existing = @() }
        }

        $combined = @($existing) + @($byDate[$k])
        $dedup = $combined | Group-Object Time,Farm,Attacker,X,Y | ForEach-Object { $_.Group | Select-Object -First 1 }
        $ordered = $dedup | Sort-Object @{
            Expression={
                $dt=[DateTime]::MinValue
                if ([DateTime]::TryParse($_.Time,[ref]$dt)) { return $dt } else { return [DateTime]::MinValue }
            }
        }, @{Expression={$_.Farm}}, @{Expression={$_.Attacker}}

        $ordered | Export-Csv -Path $dailyPath -NoTypeInformation -Encoding UTF8
        $newCount = $ordered.Count - $existing.Count
        if ($newCount -gt 0) {
            Write-Host ("Daily log updated: {0} (Total: {1}, Added: {2})" -f $dailyPath, $ordered.Count, $newCount) -ForegroundColor Green
        } else {
            Write-Verbose ("Daily log checked: {0} (Total: {1}, No new records)" -f $dailyPath, $ordered.Count)
        }
    }
}

function Write-InstanceReport {
    param(
        [int]$Index,
        [string]$Serial,
        [string]$Package,
        [string]$ScreenshotPath,
        [string]$OcrText,
        [string]$ReportPath,
        [string]$BrightSpotNote = '(none)'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $lines = @(
        "Timestamp: $timestamp"
        "Index: $Index"
        "ADB: $Serial"
        "Package: $Package"
        "Screenshot: $ScreenshotPath"
        "DismissedBrightSpot: $BrightSpotNote"
        ""
        "---- OCR ----"
        $OcrText
    )
    Set-Content -Path $ReportPath -Value $lines -Encoding UTF8
}

foreach ($index in $Indexes) {
    $instanceName = "MEmu_{0}" -f $index
    if ($ExistingLogsImagePath) {
        $finalShot = $ExistingLogsImagePath
        $logsOcrPath = $finalShot
        if ($UseLogsCrop -and -not (Split-Path $finalShot -Leaf).StartsWith('memu-logs-crop')) {
            $logsOcrPath = Join-Path $screenshotRoot ("memu-logs-crop-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
            try {
                Crop-ImageFraction -InputPath $finalShot -OutputPath $logsOcrPath -XRatio $LogsCropXRatio -YRatio $LogsCropYRatio -WidthRatio $LogsCropWidthRatio -HeightRatio $LogsCropHeightRatio | Out-Null
                Write-Host ("Cropped logs region saved to {0}" -f $logsOcrPath) -ForegroundColor DarkGray
            }
            catch {
                Write-Warning ("Logs crop failed; using full screenshot. {0}" -f $_.Exception.Message)
                $logsOcrPath = $finalShot
            }
        }

        $ocrSources = New-Object System.Collections.Generic.List[string]
        $ocrSources.Add($logsOcrPath) | Out-Null

        # Upscale to help OCR on small text
        $scaledPath = Join-Path $screenshotRoot ("memu-logs-scale-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
        try {
            Scale-Image -InputPath $logsOcrPath -OutputPath $scaledPath -ScaleFactor 2.0 | Out-Null
            $ocrSources.Add($scaledPath) | Out-Null
        }
        catch {
            Write-Verbose ("Scale failed: {0}" -f $_.Exception.Message)
        }

        # Add vertical slices to increase OCR hits
        $slicePrefix = Join-Path $screenshotRoot ("memu-logs-slice-{0}-{1}" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
        try {
            $slicePaths = Crop-ImageSlices -InputPath $logsOcrPath -OutputPrefix $slicePrefix -Slices $LogsCropSlices -OverlapRatio $LogsCropSliceOverlapRatio
            foreach ($s in $slicePaths) { $ocrSources.Add($s) | Out-Null }
        }
        catch {
            Write-Verbose ("Slice crop failed: {0}" -f $_.Exception.Message)
        }

        if ($UseLogsEnhance) {
            $baseEnhance = Join-Path $screenshotRoot ("memu-logs-enhanced-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
            try {
                Enhance-ImageForOcr -InputPath $logsOcrPath -OutputPath $baseEnhance -Threshold $LogsEnhanceThreshold -Gain $LogsEnhanceGain | Out-Null
                $ocrSources.Add($baseEnhance) | Out-Null
                Write-Host ("Enhanced logs image saved to {0}" -f $baseEnhance) -ForegroundColor DarkGray
            }
            catch {
                Write-Warning ("Logs enhancement failed; using crop/full image. {0}" -f $_.Exception.Message)
            }

            $alt1 = Join-Path $screenshotRoot ("memu-logs-enhanced-alt1-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
            try {
                Enhance-ImageForOcr -InputPath $logsOcrPath -OutputPath $alt1 -Threshold 140 -Gain 1.2 | Out-Null
                $ocrSources.Add($alt1) | Out-Null
            }
            catch { }

            $alt2 = Join-Path $screenshotRoot ("memu-logs-enhanced-alt2-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
            try {
                Enhance-ImageForOcr -InputPath $logsOcrPath -OutputPath $alt2 -Threshold 110 -Gain 1.3 | Out-Null
                $ocrSources.Add($alt2) | Out-Null
            }
            catch { }
        }

        if (-not $DisableOcr) {
            try {
                $allTexts = New-Object System.Collections.Generic.List[string]
                foreach ($src in ($ocrSources | Sort-Object -Unique)) {
                    try {
                        $allTexts.Add((Invoke-Ocr -ImagePath $src)) | Out-Null
                    }
                    catch {
                        Write-Verbose ("OCR failed on {0}: {1}" -f $src, $_.Exception.Message)
                    }
                }
                $logsText = ($allTexts -join "`n")
                $logsReport = Join-Path $ocrRoot ("memu-logs-{0}-{1}.txt" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                $lines = @(
                    "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                    "Index: $index"
                    "ADB: (skipped)"
                    "Screenshot: $finalShot"
                    "OCRSource: $($ocrSources -join ', ')"
                    ""
                    "---- Alliance Logs OCR ----"
                    $logsText
                )
                Set-Content -Path $logsReport -Value $lines -Encoding UTF8
                Write-Host ("Alliance logs OCR saved to {0}" -f $logsReport) -ForegroundColor Green

                if ($EmitLogsCsv) {
                    $records = Normalize-AttackRecords -Records (Parse-AttackLogs -Text $logsText)
                    if ($records -and $records.Count -gt 0) {
                        $csvPath = Join-Path $ocrRoot ("memu-logs-{0}-{1}.csv" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                        $records | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                        Update-DailyAttackLog -OutputRoot $ocrRoot -Records $records
                        Write-Host ("Alliance logs CSV saved to {0} (rows: {1})" -f $csvPath, $records.Count) -ForegroundColor Green
                    }
                    else {
                        Write-Verbose "No attack records parsed from alliance logs OCR."
                    }
                }
            }
            catch {
                Write-Warning ("OCR of alliance logs failed: {0}" -f $_.Exception.Message)
            }
        }
        continue
    }

    $instancePath = Join-Path $InstanceRoot $instanceName
    if (-not (Test-Path -LiteralPath $instancePath)) {
        Write-Warning "Instance folder not found: $instancePath"
        continue
    }

    $memuFile = Join-Path $instancePath ($instanceName + '.memu')
    $adbPort = try { Get-AdbPort -MemuFile $memuFile } catch {
        Write-Warning $_.Exception.Message
        continue
    }
    $serial = "127.0.0.1:$adbPort"

    if ($ExistingLogsImagePath) {
        $finalShot = $ExistingLogsImagePath
        Write-Host ("Using existing logs image: {0}" -f $finalShot) -ForegroundColor DarkGray
        $logsOcrPath = $finalShot
        if ($UseLogsCrop -and -not (Split-Path $finalShot -Leaf).StartsWith('memu-logs-crop')) {
            $logsOcrPath = Join-Path $screenshotRoot ("memu-logs-crop-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
            try {
                Crop-ImageFraction -InputPath $finalShot -OutputPath $logsOcrPath -XRatio $LogsCropXRatio -YRatio $LogsCropYRatio -WidthRatio $LogsCropWidthRatio -HeightRatio $LogsCropHeightRatio | Out-Null
                Write-Host ("Cropped logs region saved to {0}" -f $logsOcrPath) -ForegroundColor DarkGray
            }
            catch {
                Write-Warning ("Logs crop failed; using full screenshot. {0}" -f $_.Exception.Message)
                $logsOcrPath = $finalShot
            }
        }

                    $ocrSources = New-Object System.Collections.Generic.List[string]
                    $ocrSources.Add($logsOcrPath) | Out-Null

                    # Add vertical slices to increase OCR hits
                    $slicePrefix = Join-Path $screenshotRoot ("memu-logs-slice-{0}-{1}" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                    try {
                        $slicePaths = Crop-ImageSlices -InputPath $logsOcrPath -OutputPrefix $slicePrefix -Slices $LogsCropSlices -OverlapRatio $LogsCropSliceOverlapRatio
                        foreach ($s in $slicePaths) { $ocrSources.Add($s) | Out-Null }
                    }
                    catch {
                        Write-Verbose ("Slice crop failed: {0}" -f $_.Exception.Message)
                    }

                    if ($UseLogsEnhance) {
                        $baseEnhance = Join-Path $screenshotRoot ("memu-logs-enhanced-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                        try {
                            Enhance-ImageForOcr -InputPath $logsOcrPath -OutputPath $baseEnhance -Threshold $LogsEnhanceThreshold -Gain $LogsEnhanceGain | Out-Null
                            $ocrSources.Add($baseEnhance) | Out-Null
                            Write-Host ("Enhanced logs image saved to {0}" -f $baseEnhance) -ForegroundColor DarkGray
                        }
                        catch {
                            Write-Warning ("Logs enhancement failed; using crop/full image. {0}" -f $_.Exception.Message)
                        }

                        $alt1 = Join-Path $screenshotRoot ("memu-logs-enhanced-alt1-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                        try {
                            Enhance-ImageForOcr -InputPath $logsOcrPath -OutputPath $alt1 -Threshold 140 -Gain 1.2 | Out-Null
                            $ocrSources.Add($alt1) | Out-Null
                        }
                        catch { }

                        $alt2 = Join-Path $screenshotRoot ("memu-logs-enhanced-alt2-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                        try {
                            Enhance-ImageForOcr -InputPath $logsOcrPath -OutputPath $alt2 -Threshold 110 -Gain 1.3 | Out-Null
                            $ocrSources.Add($alt2) | Out-Null
                        }
                        catch { }
                    }

        if (-not $DisableOcr) {
            try {
                $allTexts = New-Object System.Collections.Generic.List[string]
                $allRecords = New-Object System.Collections.Generic.List[pscustomobject]
                foreach ($src in ($ocrSources | Sort-Object -Unique)) {
                    try {
                        $txt = Invoke-Ocr -ImagePath $src
                        $allTexts.Add($txt) | Out-Null
                        if ($EmitLogsCsv) {
                            $recs = Parse-AttackLogs -Text $txt
                            foreach ($r in $recs) { $allRecords.Add($r) | Out-Null }
                        }
                    }
                    catch {
                        Write-Verbose ("OCR failed on {0}: {1}" -f $src, $_.Exception.Message)
                    }
                }
                $logsText = ($allTexts -join "`n")
                $logsReport = Join-Path $ocrRoot ("memu-logs-{0}-{1}.txt" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                $lines = @(
                    "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                    "Index: $index"
                    "ADB: (skipped)"
                    "Screenshot: $finalShot"
                    "OCRSource: $($ocrSources -join ', ')"
                    ""
                    "---- Alliance Logs OCR ----"
                    $logsText
                )
                Set-Content -Path $logsReport -Value $lines -Encoding UTF8
                Write-Host ("Alliance logs OCR saved to {0}" -f $logsReport) -ForegroundColor Green

                if ($EmitLogsCsv) {
                    $records = $allRecords
                    if ($records -and $records.Count -gt 0) {
                        $csvPath = Join-Path $ocrRoot ("memu-logs-{0}-{1}.csv" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                        $records | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                        Update-DailyAttackLog -OutputRoot $ocrRoot -Records $records
                        Write-Host ("Alliance logs CSV saved to {0} (rows: {1})" -f $csvPath, $records.Count) -ForegroundColor Green
                    }
                    else {
                        Write-Verbose "No attack records parsed from alliance logs OCR."
                    }
                }
            }
            catch {
                Write-Warning ("OCR of alliance logs failed: {0}" -f $_.Exception.Message)
            }
        }
        continue
    }

    Write-Host "=== Processing $instanceName ($serial) ===" -ForegroundColor Cyan
    try {
        Start-MEmuInstance -Index $index -InstanceName $instanceName -Serial $serial -TimeoutSeconds $DeviceBootTimeoutSeconds -RestartAttempts $RestartAttempts -PostStartWaitSeconds $PostStartWaitSeconds -ForceStopBeforeStart:$ForceStopBeforeStart
    }
    catch {
        Write-Warning ("Unable to start {0}: {1}" -f $instanceName, $_.Exception.Message)
        continue
    }

    Unlock-Device -Serial $serial

    $package = Get-ZAppPackage -Serial $serial
    Launch-App -Serial $serial -Package $package

    if ($package) {
        Write-Host "Waiting for $package to be in focus..." -ForegroundColor Cyan
        Wait-ForAppFocus -Serial $serial -Package $package -TimeoutSeconds 60
        
        Write-Host "Waiting 30 seconds for initial app load..." -ForegroundColor Cyan
        Start-Sleep -Seconds 30
    }
    else {
        Write-Warning "Skipping app load wait because no package was identified."
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $screenshotPath = Join-Path $screenshotRoot ("memu-{0}-{1}.png" -f $index, $stamp)
    $reportPath = Join-Path $ocrRoot ("memu-{0}-{1}.txt" -f $index, $stamp)
    $brightSpotNote = '(none)'

    if ($DismissWhiteX) {
        Write-Host "Starting ad dismissal loop (max 60s)..." -ForegroundColor Cyan
        $adTimeout = 60
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $dismissCount = 0
        
        while ($sw.Elapsed.TotalSeconds -lt $adTimeout) {
            try {
                Capture-Screenshot -Serial $serial -Path $screenshotPath | Out-Null
                
                # --- Check for Alliance Icon/Text to break loop ---
                $allianceFound = $false
                
                # 1. OCR Check
                if ($UseOcrForAllianceButton) {
                    $ocrAlli = Find-OcrPhraseCenter -ImagePath $screenshotPath -Phrase "Alliance"
                    if ($ocrAlli.Found) { $allianceFound = $true }
                }
                
                # 2. Template Check (if valid)
                if (-not $allianceFound -and $AllianceButtonTemplatePath -and (Test-Path -LiteralPath $AllianceButtonTemplatePath)) {
                    try {
                        $tplMatch = Find-TemplateMatch -ImagePath $screenshotPath -TemplatePath $AllianceButtonTemplatePath -RegionFraction $SearchRegionFraction -Tolerance $ColorTolerance -Threshold $MatchThreshold -SampleStep $TemplateSampleStep
                        if ($tplMatch.Found) { $allianceFound = $true }
                    } catch {}
                }
                
                if ($allianceFound) {
                    Write-Host "Alliance icon/text detected. Stopping ad dismissal." -ForegroundColor Green
                    break
                }
                # --------------------------------------------------

                $dismissed = $false
                # Try OCR for "X", "Close", "Skip"
                $ocrX = Find-OcrPhraseCenter -ImagePath $screenshotPath -Phrase "X"
                if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $screenshotPath -Phrase "x" }
                if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $screenshotPath -Phrase "Close" }
                if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $screenshotPath -Phrase "Skip" }

                if ($ocrX.Found) {
                    Write-Host ("OCR found close button at {0},{1}; tapping to dismiss." -f $ocrX.X, $ocrX.Y) -ForegroundColor DarkYellow
                    $brightSpotNote = "OCR_X=$($ocrX.X),Y=$($ocrX.Y)"
                    Invoke-AdbTap -Serial $serial -X $ocrX.X -Y $ocrX.Y
                    $dismissCount++
                    $dismissed = $true
                    Start-Sleep -Seconds 3
                }
                else {
                    Write-Verbose "No close button (X/Close/Skip) found via OCR."
                    if ($dismissCount -gt 0) {
                        Write-Host "No more ads detected (and previous ones dismissed). Proceeding." -ForegroundColor Green
                        break
                    }
                    # If we haven't dismissed anything yet, we just wait a bit and try again, 
                    # or eventually timeout if the Alliance button never appears.
                    Start-Sleep -Seconds 2
                }
            }
            catch {
                Write-Warning ("Error during ad check: {0}" -f $_.Exception.Message)
                Start-Sleep -Seconds 2
            }
        }
    }

    try {
        Capture-Screenshot -Serial $serial -Path $screenshotPath | Out-Null
    }
    catch {
        Write-Warning ("Screenshot failed for {0}: {1}" -f $instanceName, $_.Exception.Message)
        if (-not $LeaveRunning) {
            Stop-MEmuInstance -Index $index -Serial $serial -InstanceName $instanceName
        }
        continue
    }

    $ocrText = ''
    if ($DisableOcr) {
        $ocrText = '(OCR disabled)'
    }
    else {
        try {
            $ocrText = Invoke-AnyOcr -ImagePath $screenshotPath
        }
        catch {
            $ocrText = "(OCR failed: $($_.Exception.Message))"
            Write-Warning $ocrText
        }
    }

    try {
        $packageName = if ($package) { $package } else { '(none)' }
        Write-InstanceReport -Index $index -Serial $serial -Package $packageName -ScreenshotPath $screenshotPath -OcrText $ocrText -ReportPath $reportPath -BrightSpotNote $brightSpotNote
        Write-Host ("Report saved to {0}" -f $reportPath) -ForegroundColor Green
    }
    catch {
        Write-Warning ("Failed to write report for {0}: {1}" -f $instanceName, $_.Exception.Message)
    }

    if ($WatchAllianceFlow -and -not $ExistingLogsImagePath) {
        $alliTapX = if ($AllianceTapX) { $AllianceTapX } else { $TapX }
        $alliTapY = if ($AllianceTapY) { $AllianceTapY } else { $TapY }
        # Aggressively clear popups before looking for the Alliance button
        if ($alliTapX -and $alliTapY -and $PreAllianceTapCount -gt 0) {
            for ($i = 0; $i -lt $PreAllianceTapCount; $i++) {
                Invoke-AdbTap -Serial $serial -X $alliTapX -Y $alliTapY
                Start-Sleep -Milliseconds $PreAllianceTapDelayMs
            }
        }
        if ($PreAllianceCloseX -and $PreAllianceCloseY -and $PreAllianceCloseCount -gt 0) {
            for ($i = 0; $i -lt $PreAllianceCloseCount; $i++) {
                Invoke-AdbTap -Serial $serial -X $PreAllianceCloseX -Y $PreAllianceCloseY
                Start-Sleep -Milliseconds $PreAllianceTapDelayMs
            }
        }
        if ($PreAllianceBackCount -gt 0) {
            for ($i = 0; $i -lt $PreAllianceBackCount; $i++) {
                try { Invoke-Adb -Serial $serial -Arguments @('shell','input','keyevent','4') | Out-Null } catch {
                    Write-Verbose ("Back keyevent failed: {0}" -f $_.Exception.Message)
                }
                Start-Sleep -Milliseconds $PreAllianceBackDelayMs
            }
        }

        $alliResult = $null
        $alliTapped = $false
        $alliOcrShot = $null
        if ($UseOcrForAllianceButton) {
            $alliOcrShot = Join-Path $screenshotRoot ("memu-alliance-pre-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
            try {
                Capture-Screenshot -Serial $serial -Path $alliOcrShot | Out-Null
                $ocrAlli = Find-OcrPhraseCenter -ImagePath $alliOcrShot -Phrase 'Alliance'
                if ($ocrAlli.Found -and $ocrAlli.X -and $ocrAlli.Y) {
                    Write-Host ("OCR found Alliance at {0},{1}; tapping." -f $ocrAlli.X, $ocrAlli.Y) -ForegroundColor Cyan
                    Invoke-AdbTap -Serial $serial -X $ocrAlli.X -Y $ocrAlli.Y
                    Start-Sleep -Milliseconds 400
                    $alliResult = [pscustomobject]@{ Found = $true; X = $ocrAlli.X; Y = $ocrAlli.Y; Score = 1.0; Screenshot = $alliOcrShot }
                    $alliTapped = $true
                }
            }
            catch {
                Write-Warning ("OCR search for Alliance failed: {0}" -f $_.Exception.Message)
            }
        }

        if (-not $alliResult -and $AllianceButtonTemplatePath -and (Test-Path -LiteralPath $AllianceButtonTemplatePath)) {
            $alliResult = Wait-ForIcon -Serial $serial -TemplatePath $AllianceButtonTemplatePath -TapInterval $TapIntervalSeconds -MaxWait $MaxTapWaitSeconds -RegionFraction $SearchRegionFraction -TapXOverride $alliTapX -TapYOverride $alliTapY -Tolerance $ColorTolerance -Threshold $MatchThreshold -SampleStep $TemplateSampleStep -Label 'Alliance' -TapEachInterval:$TapWhileWaitingForAlliance -TapOnDetection:$false
        }
        if ((-not $alliResult -or -not $alliResult.Found) -and $alliTapX -and $alliTapY) {
            Write-Host ("Direct tap at provided alliance coords {0},{1}" -f $alliTapX, $alliTapY) -ForegroundColor DarkYellow
            Invoke-AdbTap -Serial $serial -X $alliTapX -Y $alliTapY
            Start-Sleep -Seconds 2
            if (-not $alliResult) {
                $alliResult = [pscustomobject]@{ Found = $true; X = $alliTapX; Y = $alliTapY; Screenshot = "DirectTap" }
            }
        }
        if ((-not $alliResult -or -not $alliResult.Found) -and $AllianceButtonTemplatePath -and (Test-Path -LiteralPath $AllianceButtonTemplatePath)) {
            # Try one more detection after direct tap
            $alliResult = Wait-ForIcon -Serial $serial -TemplatePath $AllianceButtonTemplatePath -TapInterval $TapIntervalSeconds -MaxWait $MaxTapWaitSeconds -RegionFraction $SearchRegionFraction -TapXOverride $alliTapX -TapYOverride $alliTapY -Tolerance $ColorTolerance -Threshold $MatchThreshold -SampleStep $TemplateSampleStep -Label 'Alliance' -TapEachInterval:$TapWhileWaitingForAlliance -TapOnDetection:$false
        }
        if ($alliResult) {
            Write-Host ("Alliance watch result: Found={0} Screenshot={1}" -f $alliResult.Found, $alliResult.Screenshot) -ForegroundColor Cyan
            if (-not $alliTapped) { $alliTapped = $false }
            if ($alliResult.Found -and $alliResult.X -and $alliResult.Y -and -not $alliTapped) {
                Invoke-AdbTap -Serial $serial -X $alliResult.X -Y $alliResult.Y
                Start-Sleep -Milliseconds 300
                $alliTapped = $true
            }
            if ($AllianceTapX -and $AllianceTapY -and -not $alliTapped) {
                Write-Host ("Direct tap for alliance at provided coords {0},{1}" -f $AllianceTapX, $AllianceTapY) -ForegroundColor DarkYellow
                Invoke-AdbTap -Serial $serial -X $AllianceTapX -Y $AllianceTapY
                Start-Sleep -Milliseconds 300
                $alliTapped = $true
            }
            if ($alliResult.Found) {
                if ($AllianceToLogsDelaySeconds -gt 0) {
                    Start-Sleep -Seconds $AllianceToLogsDelaySeconds
                }

                if ($DismissBottomBrightX) {
                    try {
                        for ($i = 1; $i -le 3; $i++) {
                            $preClose = Join-Path $screenshotRoot ("memu-logs-preclose-{0}-{1}-{2}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'), $i)
                            Capture-Screenshot -Serial $serial -Path $preClose | Out-Null
                            
                            $dismissed = $false
                            
                            # Try OCR for "X" or "Close" first
                            $ocrX = Find-OcrPhraseCenter -ImagePath $preClose -Phrase "X"
                            if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $preClose -Phrase "x" }
                            if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $preClose -Phrase "Close" }
                            if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $preClose -Phrase "Skip" }
                            
                            if ($ocrX.Found) {
                                Write-Host ("OCR found close button at {0},{1}; tapping to dismiss." -f $ocrX.X, $ocrX.Y) -ForegroundColor DarkYellow
                                Invoke-AdbTap -Serial $serial -X $ocrX.X -Y $ocrX.Y
                                $dismissed = $true
                            }
                            else {
                                Write-Verbose "No close button (X/Close/Skip) found via OCR; skipping dismissal to avoid random clicks."
                            }
                            
                            if ($dismissed) {
                                Start-Sleep -Seconds 2
                            } else {
                                break
                            }
                        }
                    }
                    catch {
                        Write-Warning ("Failed during pre-logs dismissal check: {0}" -f $_.Exception.Message)
                    }
                }

                $logsResult = $null
                $logsTapped = $false
                $logsShotPath = $null
                $useCoordsForLogs = ($AllianceLogsTapX -and $AllianceLogsTapY -and -not $ForceOcrForLogs)

                if ($UseOcrForAllianceLogs -and -not $useCoordsForLogs) {
                    $logsShotPath = Join-Path $screenshotRoot ("memu-logs-{0}-{1}.png" -f ($serial -replace '[:\.]', '_'), (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                    Capture-Screenshot -Serial $serial -Path $logsShotPath | Out-Null
                    try {
                        $ocrLogs = Find-OcrPhraseCenter -ImagePath $logsShotPath -Phrase 'Alliance Logs'
                        if ($ocrLogs.Found -and $ocrLogs.X -and $ocrLogs.Y) {
                            Write-Host ("OCR found Alliance Logs at {0},{1}; tapping." -f $ocrLogs.X, $ocrLogs.Y) -ForegroundColor Cyan
                            Invoke-AdbTap -Serial $serial -X $ocrLogs.X -Y $ocrLogs.Y
                            Start-Sleep -Milliseconds 300
                            $logsTapped = $true
                        }
                        else {
                            Write-Verbose "OCR did not find Alliance Logs."
                        }
                    }
                    catch {
                        Write-Warning ("OCR search for Alliance Logs failed: {0}" -f $_.Exception.Message)
                    }
                }

                if ($AllianceLogsTemplatePath -and -not $logsTapped -and -not $useCoordsForLogs) {
                    $logsResult = Wait-ForIcon -Serial $serial -TemplatePath $AllianceLogsTemplatePath -TapInterval $TapIntervalSeconds -MaxWait $MaxTapWaitSeconds -RegionFraction $SearchRegionFraction -TapXOverride $TapX -TapYOverride $TapY -Tolerance $ColorTolerance -Threshold $MatchThreshold -SampleStep $TemplateSampleStep -Label 'AllianceLogs' -TapEachInterval:$false -TapOnDetection
                    if ($logsResult) {
                        Write-Host ("Alliance logs result: Found={0} Screenshot={1}" -f $logsResult.Found, $logsResult.Screenshot) -ForegroundColor Cyan
                        if ($logsResult.Found -and $logsResult.X -and $logsResult.Y) {
                            Invoke-AdbTap -Serial $serial -X $logsResult.X -Y $logsResult.Y
                            Start-Sleep -Milliseconds 300
                            $logsTapped = $true
                        }
                    }
                }

                if ($AllianceLogsTapX -and $AllianceLogsTapY -and -not $logsTapped) {
                    Write-Host ("Direct tap for alliance logs at provided coords {0},{1}" -f $AllianceLogsTapX, $AllianceLogsTapY) -ForegroundColor DarkYellow
                    Invoke-AdbTap -Serial $serial -X $AllianceLogsTapX -Y $AllianceLogsTapY
                    Start-Sleep -Milliseconds 300
                    $logsTapped = $true
                }

                if (-not $logsTapped -and (-not $logsResult -or -not $logsResult.Found)) {
                    $size = Get-ScreenSize -Serial $serial
                    if ($size) {
                        $fx = [int]([double]$size.Width * $AllianceLogsXRatio)
                        $fy = [int]([double]$size.Height * $AllianceLogsYRatio)
                        Write-Host ("Fallback tap for alliance logs at {0},{1}" -f $fx, $fy) -ForegroundColor DarkYellow
                        Invoke-AdbTap -Serial $serial -X $fx -Y $fy
                        Start-Sleep -Milliseconds 300
                    }
                    else {
                        Write-Warning "Unable to determine screen size; skipped fallback alliance logs tap."
                    }
                }

                if ($logsTapped -or ($logsResult -and $logsResult.Found)) {
                    $finalShot = $null
                    if ($ExistingLogsImagePath) {
                        $finalShot = $ExistingLogsImagePath
                        Write-Host ("Using existing logs image: {0}" -f $finalShot) -ForegroundColor DarkGray
                    }
                    else {
                        if ($DismissBottomBrightX) {
                            try {
                                for ($i = 1; $i -le 3; $i++) {
                                    $preClose = Join-Path $screenshotRoot ("memu-logs-preclose-{0}-{1}-{2}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'), $i)
                                    Capture-Screenshot -Serial $serial -Path $preClose | Out-Null
                                    
                                    $dismissed = $false
                                    
                                    # Try OCR for "X" or "Close" first
                                    $ocrX = Find-OcrPhraseCenter -ImagePath $preClose -Phrase "X"
                                    if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $preClose -Phrase "x" }
                                    if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $preClose -Phrase "Close" }
                                    if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $preClose -Phrase "Skip" }
                                    
                                    if ($ocrX.Found) {
                                        Write-Host ("OCR found close button at {0},{1}; tapping to dismiss." -f $ocrX.X, $ocrX.Y) -ForegroundColor DarkYellow
                                        Invoke-AdbTap -Serial $serial -X $ocrX.X -Y $ocrX.Y
                                        $dismissed = $true
                                    }
                                    else {
                                        Write-Verbose "No close button (X/Close/Skip) found via OCR; skipping dismissal."
                                    }
                                    
                                    if ($dismissed) {
                                        Start-Sleep -Seconds 5
                                    } else {
                                        break
                                    }
                                }
                            }
                            catch {
                                Write-Verbose ("Bottom bright dismissal skipped: {0}" -f $_.Exception.Message)
                            }
                        }
                        if ($PostLogsDelaySeconds -gt 0) {
                            Start-Sleep -Seconds $PostLogsDelaySeconds
                        }
                        $finalShot = Join-Path $screenshotRoot ("memu-logs-final-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                        Capture-Screenshot -Serial $serial -Path $finalShot | Out-Null
                        Write-Host ("Captured alliance logs screenshot: {0}" -f $finalShot) -ForegroundColor Green
                    }

                    $logsOcrPath = $finalShot
                    if ($UseLogsCrop -and $finalShot -and -not (Split-Path $finalShot -Leaf).StartsWith('memu-logs-crop')) {
                        $logsOcrPath = Join-Path $screenshotRoot ("memu-logs-crop-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                        try {
                            Crop-ImageFraction -InputPath $finalShot -OutputPath $logsOcrPath -XRatio $LogsCropXRatio -YRatio $LogsCropYRatio -WidthRatio $LogsCropWidthRatio -HeightRatio $LogsCropHeightRatio | Out-Null
                            Write-Host ("Cropped logs region saved to {0}" -f $logsOcrPath) -ForegroundColor DarkGray
                        }
                        catch {
                            Write-Warning ("Logs crop failed; using full screenshot. {0}" -f $_.Exception.Message)
                            $logsOcrPath = $finalShot
                        }
                    }

                    $ocrSource = $logsOcrPath
                    $ocrSources = New-Object System.Collections.Generic.List[string]
                    $ocrSources.Add($logsOcrPath) | Out-Null

                    # Upscale to help OCR on small text
                    $scaledPath = Join-Path $screenshotRoot ("memu-logs-scale-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                    try {
                        Scale-Image -InputPath $logsOcrPath -OutputPath $scaledPath -ScaleFactor 2.0 | Out-Null
                        $ocrSources.Add($scaledPath) | Out-Null
                    }
                    catch {
                        Write-Verbose ("Scale failed: {0}" -f $_.Exception.Message)
                    }

                    # Add vertical slices to increase OCR hits
                    $slicePrefix = Join-Path $screenshotRoot ("memu-logs-slice-{0}-{1}" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                    try {
                        $slicePaths = Crop-ImageSlices -InputPath $logsOcrPath -OutputPrefix $slicePrefix -Slices $LogsCropSlices -OverlapRatio $LogsCropSliceOverlapRatio
                        foreach ($s in $slicePaths) { $ocrSources.Add($s) | Out-Null }
                    }
                    catch {
                        Write-Verbose ("Slice crop failed: {0}" -f $_.Exception.Message)
                    }

                    if ($UseLogsEnhance) {
                        $enhancedPath = Join-Path $screenshotRoot ("memu-logs-enhanced-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                        try {
                            Enhance-ImageForOcr -InputPath $ocrSource -OutputPath $enhancedPath | Out-Null
                            $ocrSources.Add($enhancedPath) | Out-Null
                            Write-Host ("Enhanced logs image saved to {0}" -f $enhancedPath) -ForegroundColor DarkGray
                        }
                        catch {
                            Write-Warning ("Logs enhancement failed; using crop/full image. {0}" -f $_.Exception.Message)
                        }
                    }

                    if (-not $DisableOcr) {
                        try {
                            $logsTextBlocks = New-Object System.Collections.Generic.List[string]
                            $allParsedRecords = New-Object System.Collections.Generic.List[pscustomobject]
                            $runningLastTime = ''

                            foreach ($src in ($ocrSources | Sort-Object -Unique)) {
                                try { 
                                    $ocrText = Invoke-AnyOcr -ImagePath $src
                                    $logsTextBlocks.Add($ocrText) | Out-Null
                                    
                                    $parseResult = Parse-AttackLogs -Text $ocrText -InitialTime $runningLastTime
                                    $runningLastTime = $parseResult.LastTime
                                    foreach ($r in $parseResult.Records) {
                                        $allParsedRecords.Add($r) | Out-Null
                                    }
                                } catch {
                                    Write-Verbose ("OCR/Parse failed on {0}: {1}" -f $src, $_.Exception.Message)
                                }
                            }

                            $logsText = ($logsTextBlocks -join "`n")
                            $logsReport = Join-Path $ocrRoot ("memu-logs-{0}-{1}.txt" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                            $lines = @(
                                "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                                "Index: $index"
                                "ADB: $serial"
                                "Screenshot: $finalShot"
                                "OCRSource: $ocrSource"
                                ""
                                "---- Alliance Logs OCR ----"
                                $logsText
                            )
                            Set-Content -Path $logsReport -Value $lines -Encoding UTF8
                            Write-Host ("Alliance logs OCR saved to {0}" -f $logsReport) -ForegroundColor Green

                            $looksLikeAttack = ($logsText -match 'attacked by' -or $logsText -match 'X:?\s*\d+.*Y:?\s*\d+')

                            if ($TapBackOnEmptyLogs -and (-not $looksLikeAttack)) {
                                $size = Get-ScreenSize -Serial $serial
                                if ($size) {
                                    $bx = [int]([double]$size.Width * $BackXRatio)
                                    $by = [int]([double]$size.Height * $BackYRatio)
                                    Write-Host ("Logs OCR empty; tapping in-game back at {0},{1}" -f $bx, $by) -ForegroundColor DarkYellow
                                    Invoke-AdbTap -Serial $serial -X $bx -Y $by
                                    Start-Sleep -Milliseconds 500
                                    # Retry one more logs capture after back tap
                                    try {
                                        $retryShot = Join-Path $screenshotRoot ("memu-logs-retry-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                                        Capture-Screenshot -Serial $serial -Path $retryShot | Out-Null
                                        $retryCrop = $retryShot
                                        if ($UseLogsCrop -and -not (Split-Path $retryShot -Leaf).StartsWith('memu-logs-crop')) {
                                            $retryCrop = Join-Path $screenshotRoot ("memu-logs-retry-crop-{0}-{1}.png" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
                                            Crop-ImageFraction -InputPath $retryShot -OutputPath $retryCrop -XRatio $LogsCropXRatio -YRatio $LogsCropYRatio -WidthRatio $LogsCropWidthRatio -HeightRatio $LogsCropHeightRatio | Out-Null
                                        }
                                        $retryOcr = Invoke-Ocr -ImagePath $retryCrop
                                        if ($retryOcr) {
                                            $logsText = $retryOcr
                                            $looksLikeAttack = ($logsText -match 'attacked by' -or $logsText -match 'X:?\s*\d+.*Y:?\s*\d+')
                                            $logsReport = Join-Path $ocrRoot ("memu-logs-{0}-{1}.txt" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                                            $lines = @(
                                                "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                                                "Index: $index"
                                                "ADB: $serial"
                                                "Screenshot: $retryShot"
                                                "OCRSource: $retryCrop"
                                                ""
                                                "---- Alliance Logs OCR (retry) ----"
                                                $logsText
                                            )
                                            Set-Content -Path $logsReport -Value $lines -Encoding UTF8
                                            Write-Host ("Alliance logs OCR (retry) saved to {0}" -f $logsReport) -ForegroundColor Green
                                            
                                            $retryParse = Parse-AttackLogs -Text $logsText
                                            foreach ($r in $retryParse.Records) { $allParsedRecords.Add($r) | Out-Null }
                                        }
                                    }
                                    catch {
                                        Write-Verbose ("Retry OCR after back failed: {0}" -f $_.Exception.Message)
                                    }
                                }
                            }

                            if ($EmitLogsCsv) {
                                $records = @(Normalize-AttackRecords -Records $allParsedRecords)
                                # Deduplicate records based on content (Time, Farm, Attacker, X, Y)
                                $records = @($records | Group-Object Time,Farm,Attacker,X,Y | ForEach-Object { $_.Group | Select-Object -First 1 })

                                if ($records -and @($records).Count -gt 0) {
                                    $csvPath = Join-Path $ocrRoot ("memu-logs-{0}-{1}.csv" -f $index, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                                    $records | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                                    Update-DailyAttackLog -OutputRoot $ocrRoot -Records $records
                                    Write-Host ("Alliance logs CSV saved to {0} (rows: {1})" -f $csvPath, @($records).Count) -ForegroundColor Green
                                }
                                else {
                                    Write-Verbose "No attack records parsed from alliance logs OCR."
                                }
                            }
                        }
                        catch {
                            Write-Warning ("OCR of alliance logs failed: {0}" -f $_.Exception.Message)
                        }
                    }
                }
            }
        }
    }

    if ($WatchForMailIcon) {
        $watchResult = Wait-ForIcon -Serial $serial -TemplatePath $MailIconTemplatePath -TapInterval $TapIntervalSeconds -MaxWait $MaxTapWaitSeconds -RegionFraction $SearchRegionFraction -TapXOverride $TapX -TapYOverride $TapY -Tolerance $ColorTolerance -Threshold $MatchThreshold -SampleStep $TemplateSampleStep -Label 'Mail' -TapEachInterval -TapOnDetection:$false
        if ($watchResult) {
            Write-Host ("Mail watch result: Found={0} Screenshot={1}" -f $watchResult.Found, $watchResult.Screenshot) -ForegroundColor Cyan
        }
    }

    if (-not $LeaveRunning) {
        Stop-MEmuInstance -Index $index -Serial $serial -InstanceName $instanceName
    }
}

Write-Host "All requested instances processed." -ForegroundColor Green
