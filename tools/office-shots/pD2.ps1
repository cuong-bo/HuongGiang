. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$root = [Windows.Automation.AutomationElement]::RootElement
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
$dlg = Find-Ui $root 'Header and Footer' 'Window'
$cb1 = Find-Ui $dlg 'Slide number' 'CheckBox'; $cb2 = Find-Ui $dlg "Don't show on title slide" 'CheckBox'
foreach ($c in $cb1, $cb2) { $t = $c.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern); if ($t.Current.ToggleState -ne 1) { $t.Toggle() } }
Start-Sleep 1
$apply = Find-Ui $dlg 'Apply to All' 'Button'
Shot 'b5-2' @($cb1, $cb2, $apply)
Invoke-Ui $apply; Start-Sleep 2.5
$pres = $p.ActivePresentation
$pres.Slides.Item(1).HeadersFooters.SlideNumber.Visible = 0
$p.ActiveWindow.View.GotoSlide(2); Start-Sleep 1.5
$num = $null
foreach ($sh in $pres.Slides.Item(2).Shapes) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 13) { $num = $sh } }
"slide number shape on slide 2: " + ($num -ne $null)
$null = Select-RibbonTab $win 'Home'
$marks = @(); if ($num) { $marks += (Get-PRect $p $num) }
Shot 'b5-3' $marks
foreach ($i in 1..5) { "slide $i so trang hien: " + $pres.Slides.Item($i).HeadersFooters.SlideNumber.Visible }
