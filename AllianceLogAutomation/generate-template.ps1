param(
    [string]$ScreenshotPath = "$PSScriptRoot\debug_alliance_screenshot.png",
    [int]$CenterX = 500,
    [int]$CenterY = 746,
    [int]$Width = 80,
    [int]$Height = 40,
    [string]$OutputPath = "$PSScriptRoot\templates\alliance_generated.png"
)

. "$PSScriptRoot\common-utils.ps1"

$x = $CenterX - ($Width / 2)
$y = $CenterY - ($Height / 2)

Write-Host "Cropping template from $ScreenshotPath at $x,$y ($Width x $Height)..."
Crop-Image -ImagePath $ScreenshotPath -OutputPath $OutputPath -X $x -Y $y -Width $Width -Height $Height
Write-Host "Saved to $OutputPath"
