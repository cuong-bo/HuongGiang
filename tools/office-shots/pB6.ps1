. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
$lay = $pres.SlideMaster.CustomLayouts
function Place($shape, $l, $t, $w, $h) { $shape.Left = $l; $shape.Top = $t; $shape.Width = $w; $shape.Height = $h }
for ($i = 1; $i -le $lay.Count; $i++) {
  $l = $lay.Item($i)
  if ($l.Name -eq 'Gioi thieu') {
    foreach ($s in $l.Shapes) { if ($s.PlaceholderFormat.Type -eq 1) { Place $s 60 150 840 90 } elseif ($s.PlaceholderFormat.Type -eq 2) { Place $s 60 250 840 120 } }
  }
  if ($l.Name -eq 'Bao tang va anh') {
    foreach ($s in $l.Shapes) {
      $t = $s.PlaceholderFormat.Type
      if ($t -eq 1) { Place $s 40 20 880 85 } elseif ($t -eq 2) { Place $s 40 115 440 380 } elseif ($t -eq 18) { Place $s 500 115 420 380 }
    }
  }
}
# mo lai Slide Master bang giao dien
$null = Select-RibbonTab $win 'View'
Click-Ui (Find-Ui $win 'Slide Master' 'Button'); Start-Sleep 3
Wheel-At 170 700 25; Start-Sleep 1
Click-At 170 970; Start-Sleep 1.5
$null = Select-RibbonTab $win 'Slide Master'
Shot 'b3-5' @((Rc 86 916 168 102))
$cl = Find-Ui $win 'Close Master View' 'Button'
Shot 'b3-6' @($cl)
Click-Ui $cl; Start-Sleep 3
"view type: " + $p.ActiveWindow.ViewType
