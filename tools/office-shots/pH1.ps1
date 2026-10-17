. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$p = New-Object -ComObject PowerPoint.Application
$p.Visible = -1
$pres = $p.Presentations.Add(-1)
Start-Sleep 6
$win = Get-OfficeWin 'POWERPNT'
Close-Activation $win
for ($i = 0; $i -lt 10; $i++) { try { $p.WindowState = 3; break } catch { Start-Sleep 2; Close-Activation (Get-OfficeWin 'POWERPNT') } }
Start-Sleep 2
[void]$pres.Slides.AddSlide(1, $pres.SlideMaster.CustomLayouts.Item(1))
$fs = $pres.SlideMaster.Theme.ThemeFontScheme
$fs.MajorFont.Item(1).Name = 'Arial'; $fs.MinorFont.Item(1).Name = 'Arial'
foreach ($s in $pres.SlideMaster.Shapes) { if ($s.Type -eq 14) { if ($s.PlaceholderFormat.Type -eq 2) { $s.TextFrame.TextRange.ParagraphFormat.Alignment = 4 } } }
$null = Select-RibbonTab $win 'View'
Click-Ui (Find-Ui $win 'Slide Master' 'Button'); Start-Sleep 3
Wheel-At 170 700 -25; Start-Sleep 1
Click-At 150 305; Start-Sleep 2           # thumbnail Slide Master (lon nhat)
Click-At 1100 560; Start-Sleep 1.5         # nhap vao khung noi dung
$null = Select-RibbonTab $win 'Home'
$j = Find-Ui $win 'Justify' 'Button'
"Justify: " + $j.Current.BoundingRectangle
Shot 'b2-4' @($j)
