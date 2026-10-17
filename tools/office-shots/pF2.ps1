. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
$null = Select-RibbonTab $win 'Transitions'
# dong Animation Pane neu dang mo
$plan = @(@(1,'Split','b7-1',620), @(2,'Wipe','b7-2',517), @(3,'Push','b7-3',414), @(4,'Fade','b7-4',311))
foreach ($it in $plan) {
  $p.ActiveWindow.View.GotoSlide($it[0]); Start-Sleep 1.2
  Click-At ($it[3] + 51) 149; Start-Sleep 2           # bam hieu ung trong thu vien
  Shot $it[2] @((Rc $it[3] 114 103 70))
  Start-Sleep 1
}
foreach ($i in 1..5) { "slide $i entry effect = " + $pres.Slides.Item($i).SlideShowTransition.EntryEffect }
$p.ActiveWindow.View.GotoSlide(2); Start-Sleep 1
Shot 'b7-5' @((Rc 10 180 290 760))
