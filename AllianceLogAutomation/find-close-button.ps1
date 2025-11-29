. "$PSScriptRoot\common-utils.ps1"

function Find-CloseButton {
    param(
        [string]$ScreenshotPath,
        [string]$TesseractPath = $null,
        [bool]$UseTesseract = $false
    )

    $ocrX = [pscustomobject]@{ Found = $false }

    # Try OCR for "X", "Close", "Skip"
    $ocrX = Find-OcrPhraseCenter -ImagePath $ScreenshotPath -Phrase "X"
    if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $ScreenshotPath -Phrase "x" }
    if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $ScreenshotPath -Phrase "Close" }
    if (-not $ocrX.Found) { $ocrX = Find-OcrPhraseCenter -ImagePath $ScreenshotPath -Phrase "Skip" }

    # Fallback: Crop top-right 30% and check for "+" (which often misreads as X)
    if (-not $ocrX.Found) {
        $cropPath = [System.IO.Path]::ChangeExtension($ScreenshotPath, ".crop_ad.png")
        try {
            Add-Type -AssemblyName System.Drawing
            $bmp = [System.Drawing.Bitmap]::FromFile($ScreenshotPath)
            $w = $bmp.Width
            $h = $bmp.Height
            $bmp.Dispose()
            
            # Increased to 30% to catch larger/offset close buttons
            $cropW = [int]($w * 0.30)
            $cropH = [int]($h * 0.30)
            $cropX = $w - $cropW
            $cropY = 0
            
            Crop-Image -ImagePath $ScreenshotPath -OutputPath $cropPath -X $cropX -Y $cropY -Width $cropW -Height $cropH
            
            # Scale crop to improve OCR on small icons
            $scaledPath = [System.IO.Path]::ChangeExtension($cropPath, ".scaled.png")
            try {
                Scale-Image -InputPath $cropPath -OutputPath $scaledPath -ScaleFactor 2.0 | Out-Null
                # Use scaled image for OCR
                $cropPath = $scaledPath 
            } catch {
                Write-Verbose "Scaling failed: $_"
            }

            $cropText = Invoke-Ocr -ImagePath $cropPath
            
            # Try Tesseract if available
            if (-not ($cropText -match "\+" -or $cropText -match "X" -or $cropText -match "x") -and $UseTesseract -and $TesseractPath -and (Test-Path $TesseractPath)) {
                    # 1. Try raw crop
                    $cropText = Invoke-TesseractOcr -ImagePath $cropPath -TesseractExe $TesseractPath
                    
                    # 2. Try thresholding if raw failed
                    if (-not ($cropText -match "\+" -or $cropText -match "X" -or $cropText -match "x")) {
                        foreach ($th in @(0.3, 0.4, 0.5, 0.6, 0.7)) {
                            $threshPath = [System.IO.Path]::ChangeExtension($cropPath, ".thresh_$th.png")
                            try {
                                Threshold-Image -InputPath $cropPath -OutputPath $threshPath -Threshold $th | Out-Null
                                $thText = Invoke-TesseractOcr -ImagePath $threshPath -TesseractExe $TesseractPath
                                if ($thText -match "\+" -or $thText -match "X" -or $thText -match "x") {
                                    $cropText = $thText
                                    Write-Host "Found match with threshold $th" -ForegroundColor DarkGray
                                    break
                                }
                            } finally {
                                if (Test-Path -LiteralPath $threshPath) { Remove-Item -LiteralPath $threshPath -Force -ErrorAction SilentlyContinue }
                            }
                        }
                    }
            }

            if ($cropText -match "\+" -or $cropText -match "X" -or $cropText -match "x") {
                    # We found a match in the top right corner.
                    # We'll just click the center of the cropped area.
                    $clickX = $cropX + ($cropW / 2)
                    $clickY = $cropY + ($cropH / 2)
                    $ocrX = [pscustomobject]@{ Found = $true; X = [int]$clickX; Y = [int]$clickY }
                    Write-Host "Found '+' or 'X' in top-right crop. Assuming close button." -ForegroundColor Cyan
            }
        }
        catch {
            Write-Warning "Crop fallback failed: $_"
        }
        finally {
            if (Test-Path -LiteralPath $cropPath) { Remove-Item -LiteralPath $cropPath -Force -ErrorAction SilentlyContinue }
        }
    }

    # Fallback 2: Crop bottom-center 20% height and check for "Close" or "X"
    if (-not $ocrX.Found) {
        $cropPathBottom = [System.IO.Path]::ChangeExtension($ScreenshotPath, ".crop_ad_bottom.png")
        try {
            Add-Type -AssemblyName System.Drawing
            $bmp = [System.Drawing.Bitmap]::FromFile($ScreenshotPath)
            $w = $bmp.Width
            $h = $bmp.Height
            $bmp.Dispose()
            
            # Crop bottom 15%, center 60% width
            $cropW = [int]($w * 0.60)
            $cropH = [int]($h * 0.15)
            $cropX = [int](($w - $cropW) / 2)
            $cropY = $h - $cropH
            
            Crop-Image -ImagePath $ScreenshotPath -OutputPath $cropPathBottom -X $cropX -Y $cropY -Width $cropW -Height $cropH
            
            # Scale crop to improve OCR on small icons
            $scaledPathBottom = [System.IO.Path]::ChangeExtension($cropPathBottom, ".scaled.png")
            try {
                Scale-Image -InputPath $cropPathBottom -OutputPath $scaledPathBottom -ScaleFactor 2.0 | Out-Null
                # Use scaled image for OCR
                $cropPathBottom = $scaledPathBottom
            } catch {
                Write-Verbose "Scaling failed: $_"
            }

            $cropText = Invoke-Ocr -ImagePath $cropPathBottom
            
            # Try Tesseract if available
            if (-not ($cropText -match "Close" -or $cropText -match "X" -or $cropText -match "x") -and $UseTesseract -and $TesseractPath -and (Test-Path $TesseractPath)) {
                    # 1. Try raw crop
                    $cropText = Invoke-TesseractOcr -ImagePath $cropPathBottom -TesseractExe $TesseractPath
                    
                    # 2. Try thresholding if raw failed
                    if (-not ($cropText -match "Close" -or $cropText -match "X" -or $cropText -match "x")) {
                        foreach ($th in @(0.3, 0.4, 0.5, 0.6, 0.7)) {
                            $threshPath = [System.IO.Path]::ChangeExtension($cropPathBottom, ".thresh_$th.png")
                            try {
                                Threshold-Image -InputPath $cropPathBottom -OutputPath $threshPath -Threshold $th | Out-Null
                                $thText = Invoke-TesseractOcr -ImagePath $threshPath -TesseractExe $TesseractPath
                                if ($thText -match "Close" -or $thText -match "X" -or $thText -match "x") {
                                    $cropText = $thText
                                    Write-Host "Found match with threshold $th in bottom crop" -ForegroundColor DarkGray
                                    break
                                }
                            } finally {
                                if (Test-Path -LiteralPath $threshPath) { Remove-Item -LiteralPath $threshPath -Force -ErrorAction SilentlyContinue }
                            }
                        }
                    }
            }

            if ($cropText -match "Close" -or $cropText -match "X" -or $cropText -match "x") {
                $clickX = $cropX + ($cropW / 2)
                $clickY = $cropY + ($cropH / 2)
                $ocrX = [pscustomobject]@{ Found = $true; X = [int]$clickX; Y = [int]$clickY }
                Write-Host "Found 'Close' or 'X' in bottom-center crop. Assuming close button." -ForegroundColor Cyan
            }
        }
        catch {
            Write-Warning "Bottom crop fallback failed: $_"
        }
        finally {
            if (Test-Path -LiteralPath $cropPathBottom) { Remove-Item -LiteralPath $cropPathBottom -Force -ErrorAction SilentlyContinue }
        }
    }

    return $ocrX
}
