. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
$null = Select-RibbonTab $win 'View'
$sm = Find-Ui $win 'Slide Master' 'Button'
Shot 'b2-1' @($sm)
Click-Ui $sm
Start-Sleep 3
$null = Select-RibbonTab $win 'Slide Master'
$all = $win.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition)
foreach ($e in $all) { $r = $e.Current.BoundingRectangle
  if ($r.Y -gt 95 -and $r.Y -lt 200 -and $e.Current.Name -and $e.Current.ControlType.ProgrammaticName -match 'Button|MenuItem|SplitButton|CheckBox') { "{0}|{1}|{2}" -f $e.Current.ControlType.ProgrammaticName.Replace('ControlType.',''), $e.Current.Name, $r } }
