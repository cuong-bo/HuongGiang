. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$root = [Windows.Automation.AutomationElement]::RootElement
$p = New-Object -ComObject PowerPoint.Application
$p.Visible = -1
$pres = $p.Presentations.Add(-1)
Start-Sleep 4
Start-Sleep 3
$win = Get-OfficeWin 'POWERPNT'
Close-Activation $win
for ($i = 0; $i -lt 10; $i++) { try { $p.WindowState = 3; break } catch { Start-Sleep 2; Close-Activation (Get-OfficeWin 'POWERPNT') } }
Start-Sleep 2
$null = Select-RibbonTab $win 'Home'
[void]$pres.Slides.AddSlide(1, $pres.SlideMaster.CustomLayouts.Item(1))
Start-Sleep 1
Shot 'b1-1' @((Rc 20 64 40 32))
Click-At 38 80; Start-Sleep 2.5
Save-Shot (Rc 0 60 200 600) @((Rc 4 440 195 44)) (Join-Path $script:ShotDir 'b1-2.png')
Click-At 31 87; Start-Sleep 2
# --- Buoc 2: View > Slide Master ---
$null = Select-RibbonTab $win 'View'
$sm = Find-Ui $win 'Slide Master' 'Button'
Shot 'b2-1' @($sm)
Click-Ui $sm; Start-Sleep 3
# font chu cho theme (COM) roi mo hop thoai Customize Fonts
$fs = $pres.SlideMaster.Theme.ThemeFontScheme
$fs.MajorFont.Item(1).Name = 'Arial'; $fs.MinorFont.Item(1).Name = 'Arial'
$null = Select-RibbonTab $win 'Slide Master'
$fonts = Find-Ui $win 'Fonts' 'MenuItem'
$fonts.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand(); Start-Sleep 1.5
$cf = Wait-Ui $root 'Customize Fonts...' '' 5
Shot 'b2-2' @($fonts, $cf)
Invoke-Async $cf
$dlg = Wait-Ui $root 'Create New Theme Fonts' 'Window' 10
Start-Sleep 1
Set-Value (Find-Ui $dlg 'Name' 'Edit') 'Arial All'
Start-Sleep 1
Shot 'b2-3' @((Rc 682 493 266 28), (Rc 682 550 266 28), (Rc 744 590 494 24), (Rc 1051 618 88 30))
Invoke-Ui (Find-Ui $dlg 'Save' 'Button'); Start-Sleep 2
"da luu bo font; fonts=" + $fs.MajorFont.Item(1).Name + "/" + $fs.MinorFont.Item(1).Name
