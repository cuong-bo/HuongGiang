. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
Click-At 170 970; Start-Sleep 1.5
$null = Select-RibbonTab $win 'Slide Master'
Shot 'b3-5' @((Rc 86 916 168 102))
$cl = Find-Ui $win 'Close Master View' 'Button'
Shot 'b3-6' @($cl)
Click-Ui $cl; Start-Sleep 3
"view type: " + $p.ActiveWindow.ViewType
