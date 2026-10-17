. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
Click-At 1892 249; Start-Sleep 1.5          # dong Animation Pane
$p.ActiveWindow.View.GotoSlide(1); Start-Sleep 1
$null = Select-RibbonTab $win 'Slide Show'
$fb = Find-Ui $win 'From Beginning' 'Button'
if (-not $fb) { $fb = Find-Like $win 'From Beginning' }
"From Beginning: " + ($fb -ne $null) + " " + $fb.Current.BoundingRectangle
Shot 'b8-1' @($fb)
Click-At 1548 1066; Start-Sleep 2.5         # Slide Sorter
Shot 'b8-2' @()
Click-At 1498 1066; Start-Sleep 2           # ve Normal
"view type: " + $p.ActiveWindow.ViewType
