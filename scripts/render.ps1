<#
.SYNOPSIS
  PowerPoint COM으로 .pptx의 모든 슬라이드를 PNG로 내보내고, 한 장짜리 미리보기(contact sheet)를 만든다.
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File render.ps1 -Pptx C:\work\deck.pptx -OutDir C:\work\png
.NOTES
  - 실행할 때마다 OutDir의 기존 PNG를 지워, 예전 렌더링을 보고 착각하는 일을 막는다.
  - 결과: OutDir\slide01.png ... + OutDir\_contact.png
#>
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [string]$OutDir = "",
  [int]$Width = 1600,
  [int]$Height = 900,
  [int]$Columns = 3
)
$ErrorActionPreference = "Stop"
$Pptx = (Resolve-Path $Pptx).Path
if (-not $OutDir) { $OutDir = Join-Path ([IO.Path]::GetDirectoryName($Pptx)) (([IO.Path]::GetFileNameWithoutExtension($Pptx)) + "_png") }
New-Item -ItemType Directory -Force $OutDir | Out-Null
Get-ChildItem $OutDir -Filter *.png -ErrorAction SilentlyContinue | Remove-Item -Force

$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Open($Pptx, $true, $false, $false)   # ReadOnly, Untitled=false, WithWindow=false
  $n = $pres.Slides.Count
  for ($i = 1; $i -le $n; $i++) {
    $f = Join-Path $OutDir ("slide{0:D2}.png" -f $i)
    $pres.Slides.Item($i).Export($f, "PNG", $Width, $Height)
  }
  $pres.Close()
} finally {
  $ppt.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
}

# contact sheet
Add-Type -AssemblyName System.Drawing
$files = Get-ChildItem $OutDir -Filter "slide*.png" | Sort-Object Name
$tw = 480; $th = [int]($tw * $Height / $Width); $pad = 12; $label = 22
$rows = [Math]::Ceiling($files.Count / $Columns)
$bmp = New-Object System.Drawing.Bitmap ($Columns * ($tw + $pad) + $pad), ($rows * ($th + $pad + $label) + $pad)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(230, 230, 230))
$g.InterpolationMode = "HighQualityBicubic"
$font = New-Object System.Drawing.Font "Arial", 11, ([System.Drawing.FontStyle]::Bold)
for ($k = 0; $k -lt $files.Count; $k++) {
  $c = $k % $Columns; $r = [Math]::Floor($k / $Columns)
  $x = $pad + $c * ($tw + $pad); $y = $pad + $r * ($th + $pad + $label)
  $g.DrawString(("{0}" -f ($k + 1)), $font, [System.Drawing.Brushes]::Black, $x, $y)
  $img = [System.Drawing.Image]::FromFile($files[$k].FullName)
  $g.DrawImage($img, $x, $y + $label, $tw, $th)
  $img.Dispose()
}
$sheet = Join-Path $OutDir "_contact.png"
$bmp.Save($sheet, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "슬라이드 $($files.Count)장 → $OutDir"
Write-Output "미리보기: $sheet"
