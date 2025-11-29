. "$PSScriptRoot\common-utils.ps1"

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
                if ($total -gt 0 -and ($matched / $total) -ge $Threshold) {
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

function Test-AllianceVisible {
    param(
        [string]$ScreenshotPath,
        [bool]$PreferTesseract = $false,
        [string]$TesseractPath = $null,
        [string]$AllianceButtonTemplatePath = $null,
        [double]$SearchRegionFraction = 0.6,
        [int]$ColorTolerance = 28,
        [double]$MatchThreshold = 0.78,
        [int]$TemplateSampleStep = 2
    )

    # 1. OCR Check (Windows)
    $ocrAlli = Find-OcrPhraseCenter -ImagePath $ScreenshotPath -Phrase "Alliance"
    if ($ocrAlli.Found) { return $true }

    # 2. Tesseract Check (if requested)
    if ($PreferTesseract -and $TesseractPath -and (Test-Path $TesseractPath)) {
         try {
             $text = Invoke-TesseractOcr -ImagePath $ScreenshotPath -TesseractExe $TesseractPath
             if ($text -match "Alliance") { return $true }
         } catch {}
    }

    # 3. Template Check (if valid)
    if ($AllianceButtonTemplatePath -and (Test-Path -LiteralPath $AllianceButtonTemplatePath)) {
        try {
            $tplMatch = Find-TemplateMatch -ImagePath $ScreenshotPath -TemplatePath $AllianceButtonTemplatePath -RegionFraction $SearchRegionFraction -Tolerance $ColorTolerance -Threshold $MatchThreshold -SampleStep $TemplateSampleStep
            if ($tplMatch.Found) { return $true }
        } catch {}
    }

    return $false
}
