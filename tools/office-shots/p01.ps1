. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$p = New-Object -ComObject PowerPoint.Application
$p.Visible = -1
$pres = $p.Presentations.Add(-1)
Start-Sleep 4
$p.WindowState = 3     # maximized
Start-Sleep 2
$win = Get-OfficeWin 'POWERPNT'
Close-Activation $win
Start-Sleep 1
$tabs = $win.FindAll('Descendants', (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)))
"Tabs: " + (($tabs | ForEach-Object { $_.Current.Name }) -join ' | ')
"Slides: " + $pres.Slides.Count + " layout: " + $pres.SlideMaster.CustomLayouts.Count + " size: " + $pres.PageSetup.SlideWidth + "x" + $pres.PageSetup.SlideHeight
Shot '_start' @()
