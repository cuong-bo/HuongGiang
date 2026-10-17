. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$root = [Windows.Automation.AutomationElement]::RootElement
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
$p.ActiveWindow.View.GotoSlide(2); Start-Sleep 1
$null = Select-RibbonTab $win 'Insert'
$sn = Find-Ui $win 'Slide Number' 'Button'
"Slide Number ctl: " + ($sn -ne $null) + " " + $sn.Current.BoundingRectangle
Shot 'b5-1' @($sn)
Click-Ui $sn; Start-Sleep 2.5
$dlg = Find-Ui $root 'Header and Footer' 'Window'
"dialog: " + ($dlg -ne $null)
Dump-Ui $dlg 80 | Where-Object { $_ -notmatch '^Text' }
Save-Shot (Rc 400 100 1100 900) @() (Join-Path $env:TEMP 'pp_hf.png')
