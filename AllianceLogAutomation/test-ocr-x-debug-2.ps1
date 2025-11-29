param([string]$ImagePath)

if (-not $ImagePath) {
    Write-Error "Please provide an image path."
    exit 1
}

$ImagePath = Resolve-Path $ImagePath
Write-Host "Testing OCR on: $ImagePath"

function Invoke-TesseractOcr {
    param(
        [string]$ImagePath,
        [string]$TesseractExe = 'C:\Program Files\Tesseract-OCR\tesseract.exe'
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

function Crop-Image {
    param(
        [string]$ImagePath,
        [string]$OutputPath,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height
    )
    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
    
    if ($X + $Width -gt $bmp.Width) { $Width = $bmp.Width - $X }
    if ($Y + $Height -gt $bmp.Height) { $Height = $bmp.Height - $Y }

    $rect = New-Object System.Drawing.Rectangle $X, $Y, $Width, $Height
    $cropped = $bmp.Clone($rect, $bmp.PixelFormat)
    $cropped.Save($OutputPath)
    $bmp.Dispose()
    $cropped.Dispose()
}

function Threshold-Image {
    param(
        [string]$ImagePath,
        [string]$OutputPath,
        [float]$Threshold
    )
    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
    
    for ($x = 0; $x -lt $bmp.Width; $x++) {
        for ($y = 0; $y -lt $bmp.Height; $y++) {
            $c = $bmp.GetPixel($x, $y)
            $brightness = $c.GetBrightness()
            if ($brightness -gt $Threshold) {
                $bmp.SetPixel($x, $y, [System.Drawing.Color]::White)
            } else {
                $bmp.SetPixel($x, $y, [System.Drawing.Color]::Black)
            }
        }
    }
    $bmp.Save($OutputPath)
    $bmp.Dispose()
}

Add-Type -AssemblyName System.Drawing
$bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
$w = $bmp.Width
$h = $bmp.Height
$bmp.Dispose()

# Crop top-right 30%
$cropW = [int]($w * 0.30)
$cropH = [int]($h * 0.30)
$cropX = $w - $cropW
$cropY = 0
$cropPath = [System.IO.Path]::ChangeExtension($ImagePath, ".debug_crop.png")
Crop-Image -ImagePath $ImagePath -OutputPath $cropPath -X $cropX -Y $cropY -Width $cropW -Height $cropH

$thresholds = @(0.3, 0.4, 0.5, 0.6, 0.7)
foreach ($t in $thresholds) {
    $threshPath = [System.IO.Path]::ChangeExtension($ImagePath, ".debug_thresh_$t.png")
    Threshold-Image -ImagePath $cropPath -OutputPath $threshPath -Threshold $t
    try {
        $text = Invoke-TesseractOcr -ImagePath $threshPath
        Write-Host "Threshold $t : '$text'"
        if ($text -match "X" -or $text -match "x" -or $text -match "\+") {
            Write-Host "  -> MATCH FOUND!" -ForegroundColor Green
        }
    } catch {
        Write-Host "Threshold $t failed: $_"
    }
}
