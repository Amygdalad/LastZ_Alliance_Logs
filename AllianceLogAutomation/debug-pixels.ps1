param(
    [string]$ScreenshotPath = "$PSScriptRoot\debug_alliance_screenshot.png",
    [string]$TemplatePath = "$PSScriptRoot\templates\alliance_generated.png"
)

Add-Type -AssemblyName System.Drawing

$img = [System.Drawing.Bitmap]::FromFile($ScreenshotPath)
$tpl = [System.Drawing.Bitmap]::FromFile($TemplatePath)

Write-Host "Screenshot size: $($img.Width)x$($img.Height)"
Write-Host "Template size: $($tpl.Width)x$($tpl.Height)"

$x = 460
$y = 726

$pImg = $img.GetPixel($x, $y)
$pTpl = $tpl.GetPixel(0, 0)

Write-Host "Screenshot pixel at ${x},${y}: $pImg"
Write-Host "Template pixel at 0,0: $pTpl"

$pImg0 = $img.GetPixel(0, 0)
Write-Host "Screenshot pixel at 0,0: $pImg0"

$img.Dispose()
$tpl.Dispose()
