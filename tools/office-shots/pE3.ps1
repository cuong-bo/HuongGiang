. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
$null = Select-RibbonTab $win 'Animations'
$p.ActiveWindow.View.GotoSlide(5); Start-Sleep 1
$pres.Slides.Item(5).Shapes.Item(3).Select(); Start-Sleep 1
Click-At 1048 175; Start-Sleep 2           # mo rong thu vien hieu ung
Shot 'b6-4' @((Rc 222 345 72 66))
Click-At 107 80; Start-Sleep 1             # dong thu vien
"done"
