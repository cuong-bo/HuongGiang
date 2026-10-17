. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
$root = [Windows.Automation.AutomationElement]::RootElement
Close-Activation $win
$pres = $p.ActivePresentation
$fs = $pres.SlideMaster.Theme.ThemeFontScheme
$fs.MajorFont.Item(1).Name = 'Arial'; $fs.MinorFont.Item(1).Name = 'Arial'
$null = Select-RibbonTab $win 'Slide Master'
$fonts = Find-Ui $win 'Fonts' 'MenuItem'
$fonts.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand(); Start-Sleep 1.5
$cf = Wait-Ui $root 'Customize Fonts...' '' 5
"Customize Fonts: " + ($cf -ne $null)
Shot 'b2-2' @($fonts, $cf)
Invoke-Async $cf
$dlg = Wait-Ui $root 'Create New Theme Fonts' 'Window' 10
"dialog: " + ($dlg -ne $null)
Start-Sleep 1
Dump-Ui $dlg 60 | Where-Object { $_ -notmatch '^Text' }
Shot '_dlg' @()
