# Quet anh PNG: vung goc duoi phai toi mau (thong bao noi cua Windows/ung dung khac) se bi bao cao
Add-Type -AssemblyName System.Drawing
$root = Join-Path (Split-Path (Split-Path $PSScriptRoot)) 'assets\img\practical'
foreach ($f in Get-ChildItem $root -Recurse -Filter *.png) {
  $bmp = [System.Drawing.Bitmap]::FromFile($f.FullName)
  if ($bmp.Width -lt 1500) { $bmp.Dispose(); continue }
  $dark = 0; $n = 0
  for ($x = 1500; $x -lt 1880; $x += 10) { for ($y = 930; $y -lt [Math]::Min(1000, $bmp.Height); $y += 10) {
    $c = $bmp.GetPixel($x, $y); $n++; if (($c.R + $c.G + $c.B) -lt 300) { $dark++ } } }
  $bmp.Dispose()
  if ($dark / [Math]::Max($n,1) -gt 0.15) { "TOAST? {0} ({1:P0} toi)" -f $f.FullName.Substring($root.Length + 1), ($dark / $n) }
}
"quet xong"
