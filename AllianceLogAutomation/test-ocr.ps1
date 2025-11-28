#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Helper Functions (Copied from memu-report-reader.ps1) ---

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

function Parse-AttackLogs {
    param(
        [string]$Text
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
    $lastTime = ''
    foreach ($m in $matches) {
        $timeVal = $m.Groups['time'].Value
        if ($timeVal) { 
            $timeVal = $timeVal -replace '\s+', '' 
            $lastTime = $timeVal
        } elseif ($lastTime) {
            $timeVal = $lastTime
        }

        if ($logDate) {
            if ($timeVal) { $timeVal = "$logDate $timeVal" } else { $timeVal = $logDate }
        }
        $entries.Add([pscustomobject]@{
            Time     = $timeVal
            Farm     = $m.Groups['farm'].Value.Trim()
            Attacker = $m.Groups['attacker'].Value.Trim()
            X        = $m.Groups['x'].Value.Trim()
            Y        = $m.Groups['y'].Value.Trim()
        }) | Out-Null
    }

    return $entries
}

# --- Main Test Logic ---

$imagePath = "Z:\git\AI_Creations\reporting\reports\screenshots\memu-logs-final-173-20251127_142750714.png"
$tempRoot = "Z:\git\AI_Creations\reporting\reports\temp_debug"
if (-not (Test-Path $tempRoot)) { New-Item -ItemType Directory -Path $tempRoot | Out-Null }

Write-Host "Processing $imagePath"

$ocrSources = New-Object System.Collections.Generic.List[string]
$ocrSources.Add($imagePath) | Out-Null

# 1. Scale
$scaledPath = Join-Path $tempRoot "debug-scale.png"
Scale-Image -InputPath $imagePath -OutputPath $scaledPath -ScaleFactor 2.0 | Out-Null
$ocrSources.Add($scaledPath) | Out-Null

# 2. Slice
$slicePrefix = Join-Path $tempRoot "debug-slice"
$slicePaths = Crop-ImageSlices -InputPath $imagePath -OutputPrefix $slicePrefix -Slices 4 -OverlapRatio 0.10
foreach ($s in $slicePaths) { $ocrSources.Add($s) | Out-Null }

# 3. OCR and Parse Separately
$allRecords = New-Object System.Collections.Generic.List[pscustomobject]

foreach ($src in ($ocrSources | Sort-Object -Unique)) {
    Write-Host "OCR on $src"
    try {
        $text = Invoke-Ocr -ImagePath $src
        Write-Host "--- Text from $src ---"
        Write-Host $text
        Write-Host "----------------------"
        
        $recs = Parse-AttackLogs -Text $text
        Write-Host "Parsed $(@($recs).Count) records from this source."
        foreach ($r in $recs) { $allRecords.Add($r) | Out-Null }
    }
    catch {
        Write-Warning "OCR failed on $src : $_"
    }
}

Write-Host "=== All Records (Before Dedup) ==="
$allRecords | Format-Table

# Simple Dedup
$deduped = $allRecords | Select-Object Time,Farm,Attacker,X,Y -Unique
Write-Host "=== Deduped Records ==="
$deduped | Format-Table
