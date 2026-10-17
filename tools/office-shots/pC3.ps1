. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
$null = Select-RibbonTab $win 'Home'
function ShowSlide($n) { $p.ActiveWindow.View.GotoSlide($n); Start-Sleep 1.5 }
ShowSlide 1; $s = $pres.Slides.Item(1)
Shot 'b4-2' @((Get-PRect $p $s.Shapes.Item(1)), (Get-PRect $p $s.Shapes.Item(2)))
ShowSlide 2; $s = $pres.Slides.Item(2)
Shot 'b4-3' @((Get-PRect $p $s.Shapes.Item(1)), (Get-PRect $p $s.Shapes.Item(2)))
ShowSlide 3; $s = $pres.Slides.Item(3)
Shot 'b4-4' @((Get-PRect $p $s.Shapes.Item(1)), (Get-PRect $p $s.Shapes.Item(2)), (Get-PRect $p $s.Shapes.Item(3)))
ShowSlide 5
Shot 'b4-5' @((Rc 10 180 290 700))
