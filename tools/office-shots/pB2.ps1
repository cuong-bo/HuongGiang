. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$root = [Windows.Automation.AutomationElement]::RootElement
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
$pres = $p.ActivePresentation
$dlg = Find-Ui $root 'Rename Layout' 'Window'
Set-Value (Find-Ui $dlg 'Layout name' 'Edit') 'Gioi thieu'
Start-Sleep 1
Shot 'b3-2' @((Rc 861 536 198 24), (Rc 872 564 88 30))
Invoke-Ui (Find-Ui $dlg 'Rename' 'Button'); Start-Sleep 2
# menu Insert Placeholder
$null = Select-RibbonTab $win 'Slide Master'
Click-At 380 175; Start-Sleep 1.5
Shot 'b3-3' @((Rc 331 102 98 97))
Click-At 130 80; Start-Sleep 1           # dong menu
# Noi dung layout 1 va layout 2 bang COM
$lay = $pres.SlideMaster.CustomLayouts
"so layout: " + $lay.Count
$l1 = $lay.Item($lay.Count)
"ten layout vua doi: " + $l1.Name
$null = $l1.Shapes.AddPlaceholder(2, 60, 250, 840, 120)    # ppPlaceholderBody (Text) lam phu de
$l2 = $lay.Add($lay.Count + 1)
$l2.Name = 'Bao tang va anh'
$null = $l2.Shapes.AddPlaceholder(2, 40, 120, 440, 380)
$null = $l2.Shapes.AddPlaceholder(18, 500, 120, 420, 380)  # ppPlaceholderPicture
foreach ($s in $l2.Shapes) { "L2: " + $s.Name + " type=" + $s.PlaceholderFormat.Type + " " + [int]$s.Left + "," + [int]$s.Top }
foreach ($s in $l1.Shapes) { "L1: " + $s.Name + " type=" + $s.PlaceholderFormat.Type + " " + [int]$s.Left + "," + [int]$s.Top }
Save-Shot (Rc 0 60 400 1000) @() (Join-Path $env:TEMP 'pp_panel.png')
