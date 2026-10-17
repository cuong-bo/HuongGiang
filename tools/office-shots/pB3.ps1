. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
$lay = $pres.SlideMaster.CustomLayouts
# go khung chu them nham o layout 'Vertical Title and Text'
for ($i = 1; $i -le $lay.Count; $i++) {
  $l = $lay.Item($i)
  if ($l.Name -eq 'Vertical Title and Text') {
    foreach ($s in @($l.Shapes)) { if ($s.Name -eq 'Text Placeholder 6') { $s.Delete(); "da go khung nham" } }
  }
}
for ($i = 1; $i -le $lay.Count; $i++) {
  $l = $lay.Item($i)
  if ($l.Name -eq 'Gioi thieu') { $null = $l.Shapes.AddPlaceholder(2, 60, 250, 840, 120); "da them khung phu de vao Gioi thieu (vi tri $i)" }
}
"danh sach layout:"; for ($i = 1; $i -le $lay.Count; $i++) { "  $i " + $lay.Item($i).Name }
$null = Select-RibbonTab $win 'Slide Master'
Shot 'b3-4' @((Rc 86 368 168 102))
