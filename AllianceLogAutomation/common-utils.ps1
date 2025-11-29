
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
        # Force GC to release file handles held by WinRT/COM objects
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Invoke-Ocr {
    param([string]$ImagePath)
    $result = Get-OcrResult -ImagePath $ImagePath
    return $result.Text
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
    $bmp.Dispose() # Dispose source immediately
    
    $cropped.Save($OutputPath)
    $cropped.Dispose()
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

function Threshold-Image {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [float]$Threshold = 0.5
    )

    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($InputPath)
    try {
        # Lock bits for faster processing
        $rect = New-Object System.Drawing.Rectangle 0, 0, $bmp.Width, $bmp.Height
        $bmpData = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadWrite, $bmp.PixelFormat)
        $ptr = $bmpData.Scan0
        $bytes = [Math]::Abs($bmpData.Stride) * $bmp.Height
        $rgbValues = New-Object byte[] $bytes
        [System.Runtime.InteropServices.Marshal]::Copy($ptr, $rgbValues, 0, $bytes)

        $pixelSize = 4 # Assuming 32bpp
        if ($bmp.PixelFormat -eq [System.Drawing.Imaging.PixelFormat]::Format24bppRgb) { $pixelSize = 3 }

        for ($i = 0; $i -lt $rgbValues.Length; $i += $pixelSize) {
            $b = $rgbValues[$i]
            $g = $rgbValues[$i+1]
            $r = $rgbValues[$i+2]
            
            # Simple grayscale
            $gray = ($r * 0.3 + $g * 0.59 + $b * 0.11)
            
            # Threshold
            $val = if ($gray -lt ($Threshold * 255)) { 0 } else { 255 }
            
            $rgbValues[$i] = $val
            $rgbValues[$i+1] = $val
            $rgbValues[$i+2] = $val
        }

        [System.Runtime.InteropServices.Marshal]::Copy($rgbValues, 0, $ptr, $bytes)
        $bmp.UnlockBits($bmpData)
        
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
