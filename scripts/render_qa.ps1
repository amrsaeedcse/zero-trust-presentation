param(
    [string]$InputPptx = "d:\zero_trust_presentation\dist\Zero_Trust_Campus_Networks.pptx",
    [string]$OutDir = "d:\zero_trust_presentation\dist\qa_render",
    [int]$Width = 1600,
    [int]$Height = 900
)

$ErrorActionPreference = "Stop"
$src = (Resolve-Path $InputPptx).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
Get-ChildItem $OutDir -Filter "*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($src, $true, $false, $false)
$n = $pres.Slides.Count
Write-Host "Exporting $n slides from $src ..."
for ($i = 1; $i -le $n; $i++) {
    $f = Join-Path $OutDir ("slide-{0:d2}.png" -f $i)
    $pres.Slides.Item($i).Export($f, "PNG", $Width, $Height)
}
$pres.Close()
if (-not $wasRunning) { $pp.Quit() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null

# Build combined 3-column QA contact sheet grid.png
Add-Type -AssemblyName System.Drawing
$files = Get-ChildItem $OutDir -Filter "slide-*.png" | Sort-Object Name
$cols = 3
$thumbW = 520; $thumbH = 293; $pad = 16
$rows = [math]::Ceiling($files.Count / $cols)
$gridW = $cols * $thumbW + ($cols + 1) * $pad
$gridH = $rows * $thumbH + ($rows + 1) * $pad
$bmp = New-Object System.Drawing.Bitmap($gridW, $gridH)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(18, 24, 26))
for ($i = 0; $i -lt $files.Count; $i++) {
    $img = [System.Drawing.Image]::FromFile($files[$i].FullName)
    $c = $i % $cols
    $r = [math]::Floor($i / $cols)
    $x = $pad + $c * ($thumbW + $pad)
    $y = $pad + $r * ($thumbH + $pad)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($img, $x, $y, $thumbW, $thumbH)
    $g.DrawRectangle([System.Drawing.Pens]::DimGray, $x, $y, $thumbW, $thumbH)
    $img.Dispose()
}
$gridPath = Join-Path $OutDir "grid.png"
$bmp.Save($gridPath, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Host "Exported $n PNGs + QA Grid: $gridPath"
