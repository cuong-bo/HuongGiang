. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
$null = Select-RibbonTab $win 'Animations'
$p.ActiveWindow.View.GotoSlide(3); Start-Sleep 1
$pres.Slides.Item(3).Shapes.Item(1).Select(); Start-Sleep 1
Shot 'b6-1' @((Rc 311 114 103 70))
Click-At 1325 117; Start-Sleep 2           # Animation Pane
Shot 'b6-2' @((Rc 1252 102 146 30))
$p.ActiveWindow.View.GotoSlide(4); Start-Sleep 1
$pres.Slides.Item(4).Shapes.Item(3).Select(); Start-Sleep 1
Shot 'b6-3' @((Rc 414 114 103 70))
$p.ActiveWindow.View.GotoSlide(5); Start-Sleep 1
$pres.Slides.Item(5).Shapes.Item(3).Select(); Start-Sleep 1
Click-At 1048 175; Start-Sleep 2           # mo rong thu vien hieu ung
Shot 'b6-4' @((Rc 222 345 72 66))
Click-At 107 80; Start-Sleep 1             # dong thu vien
"done"
