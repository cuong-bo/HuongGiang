. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$root = [Windows.Automation.AutomationElement]::RootElement
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
# ---- anh minh hoa tu ve (khong dung anh ben ngoai) ----
$names = 'Bảo tàng Lịch sử Hà Nội', 'Bảo tàng Hà Nội', 'Bảo tàng Hồ Chí Minh'
$cols = @(@(180,60,40), @(30,110,170), @(40,130,80))
for ($i = 0; $i -lt 3; $i++) {
  $bmp = New-Object System.Drawing.Bitmap(840, 760); $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'
  $c = $cols[$i]
  $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush([System.Drawing.Point]::new(0,0), [System.Drawing.Point]::new(0,760), [System.Drawing.Color]::FromArgb(255, 235, 240, 250), [System.Drawing.Color]::FromArgb(255, $c[0], $c[1], $c[2]))
  $g.FillRectangle($br, 0, 0, 840, 760)
  $dark = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(230, 250, 250, 250))
  $g.FillRectangle($dark, 150, 300, 540, 330)                       # than toa nha
  $roof = [System.Drawing.Point[]]@([System.Drawing.Point]::new(120,300), [System.Drawing.Point]::new(420,160), [System.Drawing.Point]::new(720,300))
  $g.FillPolygon($dark, $roof)
  $pen = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, $c[0], $c[1], $c[2]))
  foreach ($k in 0..4) { $g.FillRectangle($pen, 190 + $k * 100, 350, 40, 230) }   # cot
  $font = New-Object System.Drawing.Font('Arial', 30, [System.Drawing.FontStyle]::Bold)
  $fmt = New-Object System.Drawing.StringFormat; $fmt.Alignment = 'Center'
  $g.DrawString($names[$i], $font, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]::new(20, 660, 800, 90), $fmt)
  $bmp.Save("$env:TEMP\bt$($i + 1).png", [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose()
}
"anh da ve: " + ((Get-ChildItem "$env:TEMP\bt*.png").Count)
# ---- mo thu vien Layout tren tab Home ----
$null = Select-RibbonTab $win 'Home'
$lay = Find-Ui $win 'Layout' 'MenuItem'
if (-not $lay) { $lay = Find-Like $win '^Layout$' }
"Layout ctl: " + $lay.Current.ControlType.ProgrammaticName + " " + $lay.Current.BoundingRectangle
$lay.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand(); Start-Sleep 2
Save-Shot (Rc 0 60 1100 800) @() (Join-Path $env:TEMP 'pp_layoutmenu.png')
