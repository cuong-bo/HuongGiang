. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$p.ActiveWindow.View.GotoSlide(1); Start-Sleep 1
$null = Select-RibbonTab $win 'Slide Show'
Shot 'b8-1' @((Rc 25 102 72 96))
$pres = $p.ActivePresentation
while ($p.Presentations.Count -gt 0) { $p.Presentations.Item(1).Close() }
"da dong bai trinh chieu"
