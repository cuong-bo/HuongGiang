. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$root = [Windows.Automation.AutomationElement]::RootElement
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
# --- justify cho Slide Master ---
Click-At 170 296                         # thumbnail dau tien (Slide Master)
Start-Sleep 1.5
foreach ($s in $pres.SlideMaster.Shapes) { if ($s.Type -eq 14) { if ($s.PlaceholderFormat.Type -eq 2) { $s.TextFrame.TextRange.ParagraphFormat.Alignment = 4 } } }
Click-At 1100 600; Start-Sleep 1         # chon khung noi dung
$null = Select-RibbonTab $win 'Home'
$j = Find-Ui $win 'Justify' 'Button'
Shot 'b2-4' @($j)
# --- Buoc 3: Insert Layout ---
$null = Select-RibbonTab $win 'Slide Master'
$il = Find-Ui $win 'Insert Layout' 'Button'
Shot 'b3-1' @($il)
Click-Ui $il; Start-Sleep 2
$rn = Find-Ui $win 'Rename' 'Button'
Click-Ui $rn; Start-Sleep 2
$dlg = Find-Like $root 'Rename' 'Window'
"rename dialog: " + ($dlg -ne $null) + " " + $dlg.Current.Name
if ($dlg) { Dump-Ui $dlg 30 | Where-Object { $_ -notmatch '^Text' } }
Save-Shot (Rc 500 150 900 800) @() (Join-Path $env:TEMP 'pp_rename.png')
