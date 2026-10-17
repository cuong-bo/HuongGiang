. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$null = Select-RibbonTab $win 'View'
Click-Ui (Find-Ui $win 'Slide Master' 'Button'); Start-Sleep 3
Wheel-At 170 700 -25; Start-Sleep 1
Click-At 170 559; Start-Sleep 1.5
$null = Select-RibbonTab $win 'Slide Master'
Shot 'b3-4' @((Rc 86 511 168 98))
Click-Ui (Find-Ui $win 'Close Master View' 'Button'); Start-Sleep 3
"view type: " + $p.ActiveWindow.ViewType
