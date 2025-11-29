param(
    [string]$MEmuRoot = 'Z:\Program Files\Microvirt\MEmu',
    [string]$TemplateFolder = "$PSScriptRoot\templates"
)

. "$PSScriptRoot\common-utils.ps1"
. "$PSScriptRoot\check-alliance.ps1"

$adbExe = Join-Path $MEmuRoot "adb.exe"
if (-not (Test-Path $adbExe)) {
    Write-Error "adb.exe not found at $adbExe"
    exit
}

Write-Host "Checking for connected devices..."
$devicesOut = & $adbExe devices
$devices = $devicesOut | Where-Object { $_ -match '\tdevice$' }
if (-not $devices) {
    Write-Warning "No devices connected."
    exit
}

$serial = $devices[0].Split("`t")[0]
Write-Host "Using device: $serial" -ForegroundColor Cyan

$screenshotPath = "$PSScriptRoot\debug_alliance_screenshot.png"

function Capture-Screenshot-Local {
    param($Serial, $Path)
    $cmd = '"{0}" -s "{1}" exec-out "screencap -p" > "{2}"' -f $adbExe, $Serial, $Path
    cmd.exe /c $cmd | Out-Null
}

Write-Host "Capturing screenshot to $screenshotPath..."
Capture-Screenshot-Local -Serial $serial -Path $screenshotPath

if (-not (Test-Path $screenshotPath)) {
    Write-Error "Failed to capture screenshot."
    exit
}

Add-Type -AssemblyName System.Drawing
$bmp = [System.Drawing.Bitmap]::FromFile($screenshotPath)
Write-Host "Screenshot Resolution: $($bmp.Width) x $($bmp.Height)" -ForegroundColor Cyan
$bmp.Dispose()

# 1. Test OCR
Write-Host "`n--- Testing OCR (Windows) ---" -ForegroundColor Yellow
$ocrResult = Find-OcrPhraseCenter -ImagePath $screenshotPath -Phrase "Alliance"
if ($ocrResult.Found) {
    Write-Host "OCR Found 'Alliance' at X=$($ocrResult.X), Y=$($ocrResult.Y)" -ForegroundColor Green
} else {
    Write-Host "OCR did NOT find 'Alliance'" -ForegroundColor Red
    # Dump all text found
    $allText = Invoke-Ocr -ImagePath $screenshotPath
    Write-Host "Full OCR Text found:" -ForegroundColor DarkGray
    Write-Host $allText -ForegroundColor DarkGray
}

# 2. Test Templates
Write-Host "`n--- Testing Templates ---" -ForegroundColor Yellow
$templates = Get-ChildItem $TemplateFolder -Filter "*alliance*.png"
if (-not $templates) {
    Write-Warning "No alliance templates found in $TemplateFolder"
}

function Find-TemplateMatch-Debug {
    param(
        [string]$ImagePath,
        [string]$TemplatePath,
        [double]$RegionFraction = 1.0,
        [int]$Tolerance = 20,
        [double]$Threshold = 0.90,
        [int]$SampleStep = 2
    )

    Add-Type -AssemblyName System.Drawing

    $img = [System.Drawing.Bitmap]::FromFile($ImagePath)
    $tpl = [System.Drawing.Bitmap]::FromFile($TemplatePath)
    try {
        $startX = 0
        $startY = 0
        $endX = $img.Width - $tpl.Width
        $endY = $img.Height - $tpl.Height
        $step = [Math]::Max(1, [int]$SampleStep)

        Write-Host "  Searching area: $endX x $endY with step $step"

        for ($y = $startY; $y -le $endY; $y += $step) {
            for ($x = $startX; $x -le $endX; $x += $step) {
                $matched = 0
                $total = 0
                for ($ty = 0; $ty -lt $tpl.Height; $ty += $step) {
                    for ($tx = 0; $tx -lt $tpl.Width; $tx += $step) {
                        $cTpl = $tpl.GetPixel($tx, $ty)
                        $cImg = $img.GetPixel($x + $tx, $y + $ty)
                        $total++
                        $dist = Get-ColorDistance -A $cTpl -B $cImg
                        if ($x -eq 0 -and $y -eq 0 -and $tx -eq 0 -and $ty -eq 0) {
                             Write-Host "DEBUG: 0,0 check. Tpl=$cTpl Img=$cImg Dist=$dist Tol=$Tolerance"
                        }
                        if ($dist -le $Tolerance) {
                            $matched++
                        }
                    }
                }
                $score = if ($total -gt 0) { $matched / $total } else { 0 }
                if ($score -ge $Threshold) {
                    Write-Host "  Match found at $x,$y with score $score" -ForegroundColor Green
                    return [pscustomobject]@{
                        Found = $true
                        X = $x + ($tpl.Width / 2)
                        Y = $y + ($tpl.Height / 2)
                    }
                }
            }
        }
    }
    finally {
        $img.Dispose()
        $tpl.Dispose()
    }
    return [pscustomobject]@{ Found = $false }
}

foreach ($tpl in $templates) {
    Write-Host "Testing template: $($tpl.Name)"
    try {
        $match = Find-TemplateMatch-Debug -ImagePath $screenshotPath -TemplatePath $tpl.FullName -RegionFraction 1.0 -Tolerance 20 -Threshold 0.90
        if ($match.Found) {
            Write-Host "  MATCH FOUND! X=$($match.X), Y=$($match.Y)" -ForegroundColor Green
        } else {
            Write-Host "  No match." -ForegroundColor Red
        }
    } catch {
        Write-Host "  Error: $_" -ForegroundColor Red
    }
}

# 3. Test Tesseract (if available)
$tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe" # Common path, adjust if needed
if (Test-Path $tesseractPath) {
    Write-Host "`n--- Testing Tesseract ---" -ForegroundColor Yellow
    try {
        $text = Invoke-TesseractOcr -ImagePath $screenshotPath -TesseractExe $tesseractPath
        if ($text -match "Alliance") {
            Write-Host "Tesseract Found 'Alliance'" -ForegroundColor Green
        } else {
            Write-Host "Tesseract did NOT find 'Alliance'" -ForegroundColor Red
            Write-Host "Full Tesseract Text:" -ForegroundColor DarkGray
            Write-Host $text -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "Tesseract Error: $_" -ForegroundColor Red
    }
} else {
    Write-Host "`n(Tesseract not found at default location, skipping)" -ForegroundColor DarkGray
}

Write-Host "`nDone."
