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

$files = Get-ChildItem "z:\git\AI_Creations\reporting\reports\screenshots\*.png" | Sort-Object LastWriteTime -Descending | Select-Object -First 20

foreach ($file in $files) {
    Write-Host "Checking $($file.Name)..."
    try {
        $res = Get-OcrResult -ImagePath $file.FullName
        # Write-Host "  Text: $($res.Text.Substring(0, [Math]::Min($res.Text.Length, 50)))..."
        if ($res.Text -match "ApocaAid" -or $res.Text -match "Monthly Pass" -or $res.Text -match "Privilege") {
            Write-Host "FOUND MATCH: $($file.FullName)" -ForegroundColor Green
            Write-Host $res.Text
            break
        }
    } catch {
        Write-Host "Error processing $($file.Name): $_"
    }
}
