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
        $stream.Dispose()
        throw "Windows OCR engine not available on this machine."
    }
    try {
        return Get-AsyncResult -AsyncOperation ($engine.RecognizeAsync($bitmap)) -ResultType ([Windows.Media.Ocr.OcrResult])
    }
    finally {
        if ($stream) { $stream.Dispose() }
    }
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
    
    # Ensure crop is within bounds
    if ($X + $Width -gt $bmp.Width) { $Width = $bmp.Width - $X }
    if ($Y + $Height -gt $bmp.Height) { $Height = $bmp.Height - $Y }

    $rect = New-Object System.Drawing.Rectangle $X, $Y, $Width, $Height
    $cropped = $bmp.Clone($rect, $bmp.PixelFormat)
    $cropped.Save($OutputPath)
    $bmp.Dispose()
    $cropped.Dispose()
}

function Invert-Image {
    param(
        [string]$ImagePath,
        [string]$OutputPath
    )
    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
    
    for ($x = 0; $x -lt $bmp.Width; $x++) {
        for ($y = 0; $y -lt $bmp.Height; $y++) {
            $c = $bmp.GetPixel($x, $y)
            $newC = [System.Drawing.Color]::FromArgb($c.A, 255 - $c.R, 255 - $c.G, 255 - $c.B)
            $bmp.SetPixel($x, $y, $newC)
        }
    }
    $bmp.Save($OutputPath)
    $bmp.Dispose()
}

function Threshold-Image {
    param(
        [string]$ImagePath,
        [string]$OutputPath,
        [float]$Threshold = 0.5
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

function Scale-Image {
    param(
        [string]$ImagePath,
        [string]$OutputPath,
        [float]$ScaleFactor = 2.0
    )
    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
    $newWidth = [int]($bmp.Width * $ScaleFactor)
    $newHeight = [int]($bmp.Height * $ScaleFactor)
    $newBmp = New-Object System.Drawing.Bitmap $newWidth, $newHeight
    $g = [System.Drawing.Graphics]::FromImage($newBmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($bmp, 0, 0, $newWidth, $newHeight)
    $g.Dispose()
    $newBmp.Save($OutputPath)
    $bmp.Dispose()
    $newBmp.Dispose()
}

Write-Host "--- Full Windows OCR Text ---"
$text = Invoke-Ocr -ImagePath $ImagePath
Write-Host $text
Write-Host "---------------------"

$targets = @("X", "x", "Close", "Skip")
foreach ($t in $targets) {
    $res = Find-OcrPhraseCenter -ImagePath $ImagePath -Phrase $t
    if ($res.Found) {
        Write-Host "Found '$t' at $($res.X), $($res.Y)" -ForegroundColor Green
    } else {
        Write-Host "Did not find '$t'" -ForegroundColor Yellow
    }
}

Write-Host "`n--- Tesseract OCR Results ---"
try {
    $tessText = Invoke-TesseractOcr -ImagePath $ImagePath
    Write-Host "Raw Tesseract Output:"
    Write-Host $tessText
    if ($tessText -match "X" -or $tessText -match "x" -or $tessText -match "Close" -or $tessText -match "Skip") {
        Write-Host "  -> MATCH FOUND with Tesseract!" -ForegroundColor Green
    } else {
        Write-Host "Did not find match with Tesseract." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Tesseract failed: $_" -ForegroundColor Red
}

Write-Host "`n--- Cropped Image Test (Top Right 30%) ---"
Add-Type -AssemblyName System.Drawing
$bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
$w = $bmp.Width
$h = $bmp.Height
$bmp.Dispose()

# Crop top-right 30% width, top 30% height
$cropW = [int]($w * 0.30)
$cropH = [int]($h * 0.30)
$cropX = $w - $cropW
$cropY = 0

$croppedPath = [System.IO.Path]::ChangeExtension($ImagePath, ".cropped_30.png")
Crop-Image -ImagePath $ImagePath -OutputPath $croppedPath -X $cropX -Y $cropY -Width $cropW -Height $cropH
Write-Host "Created 30% cropped image: $croppedPath"

# Scale the crop for better OCR
$scaledPath = [System.IO.Path]::ChangeExtension($croppedPath, ".scaled.png")
Scale-Image -ImagePath $croppedPath -OutputPath $scaledPath -ScaleFactor 2.0
Write-Host "Scaled crop to: $scaledPath"
$croppedPath = $scaledPath # Use scaled image for subsequent tests

Write-Host "`n--- Windows OCR on 30% Crop (Scaled) ---"
$text = Invoke-Ocr -ImagePath $croppedPath
Write-Host "Text: '$text'"
if ($text -match "X" -or $text -match "x" -or $text -match "Close" -or $text -match "Skip" -or $text -match "\+") {
    Write-Host "  -> MATCH FOUND in 30% crop (including '+')!" -ForegroundColor Green
}

Write-Host "`n--- Tesseract OCR on 30% Crop ---"
try {
    $tessText = Invoke-TesseractOcr -ImagePath $croppedPath
    Write-Host "Raw Tesseract Output:"
    Write-Host $tessText
    if ($tessText -match "X" -or $tessText -match "x" -or $tessText -match "Close" -or $tessText -match "Skip" -or $tessText -match "\+") {
        Write-Host "  -> MATCH FOUND with Tesseract in 30% crop (including '+')!" -ForegroundColor Green
    }
} catch {
    Write-Host "Tesseract failed: $_" -ForegroundColor Red
}

Write-Host "`n--- Thresholded Tesseract OCR on 30% Crop (Multiple Thresholds) ---"
$thresholds = @(0.3, 0.4, 0.5, 0.6, 0.7)
foreach ($t in $thresholds) {
    $threshPath = [System.IO.Path]::ChangeExtension($croppedPath, ".thresh_$t.png")
    Threshold-Image -ImagePath $croppedPath -OutputPath $threshPath -Threshold $t
    try {
        $tessText = Invoke-TesseractOcr -ImagePath $threshPath
        Write-Host "Threshold $t Output: '$tessText'"
        if ($tessText -match "X" -or $tessText -match "x" -or $tessText -match "Close" -or $tessText -match "Skip" -or $tessText -match "\+") {
            Write-Host "  -> MATCH FOUND with Tesseract in 30% crop (Threshold $t)!" -ForegroundColor Green
        }
    } catch {
        Write-Host "Tesseract failed at threshold $t : $_" -ForegroundColor Red
    }
}

Write-Host "`n--- Cropped Image Test (Bottom Center 60% Width, 15% Height) ---"
# Crop bottom 15%, center 60% width
$cropW = [int]($w * 0.60)
$cropH = [int]($h * 0.15)
$cropX = [int](($w - $cropW) / 2)
$cropY = $h - $cropH

$croppedPathBottom = [System.IO.Path]::ChangeExtension($ImagePath, ".cropped_bottom.png")
Crop-Image -ImagePath $ImagePath -OutputPath $croppedPathBottom -X $cropX -Y $cropY -Width $cropW -Height $cropH
Write-Host "Created bottom cropped image: $croppedPathBottom"

Write-Host "`n--- Windows OCR on Bottom Crop ---"
$text = Invoke-Ocr -ImagePath $croppedPathBottom
Write-Host "Text: '$text'"
if ($text -match "X" -or $text -match "x" -or $text -match "Close" -or $text -match "Skip") {
    Write-Host "  -> MATCH FOUND in Bottom crop!" -ForegroundColor Green
}

Write-Host "`n--- Tesseract OCR on Bottom Crop ---"
try {
    $tessText = Invoke-TesseractOcr -ImagePath $croppedPathBottom
    Write-Host "Raw Tesseract Output:"
    Write-Host $tessText
    if ($tessText -match "X" -or $tessText -match "x" -or $tessText -match "Close" -or $tessText -match "Skip") {
        Write-Host "  -> MATCH FOUND with Tesseract in Bottom crop!" -ForegroundColor Green
    }
} catch {
    Write-Host "Tesseract failed: $_" -ForegroundColor Red
}

Write-Host "`n--- Thresholded Tesseract OCR on Bottom Crop (Multiple Thresholds) ---"
$thresholds = @(0.3, 0.4, 0.5, 0.6, 0.7)
foreach ($t in $thresholds) {
    $threshPathBottom = [System.IO.Path]::ChangeExtension($croppedPathBottom, ".thresh_$t.png")
    Threshold-Image -ImagePath $croppedPathBottom -OutputPath $threshPathBottom -Threshold $t
    try {
        $tessText = Invoke-TesseractOcr -ImagePath $threshPathBottom
        Write-Host "Threshold $t Output: '$tessText'"
        if ($tessText -match "X" -or $tessText -match "x" -or $tessText -match "Close" -or $tessText -match "Skip") {
            Write-Host "  -> MATCH FOUND with Tesseract in Bottom crop (Threshold $t)!" -ForegroundColor Green
        }
    } catch {
        Write-Host "Tesseract failed at threshold $t : $_" -ForegroundColor Red
    }
}
