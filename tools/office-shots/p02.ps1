. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
[void]$pres.Slides.AddSlide(1, $pres.SlideMaster.CustomLayouts.Item(1))
Start-Sleep 1
Shot 'b1-1' @((Rc 20 64 40 32))
Click-At 38 80; Start-Sleep 2.5
Save-Shot (Rc 0 60 200 600) @((Rc 4 440 195 44)) (Join-Path $script:ShotDir 'b1-2.png')
Click-At 31 87; Start-Sleep 2
# Buoc 2: View > Slide Master
$null = Select-RibbonTab $win 'View'
$sm = Find-Ui $win 'Slide Master' 'Button'
if (-not $sm) { $sm = Find-Like $win '^Slide Master$' }
"Slide Master ctl: " + ($sm -ne $null) + " " + $sm.Current.ControlType.ProgrammaticName + " " + $sm.Current.BoundingRectangle
