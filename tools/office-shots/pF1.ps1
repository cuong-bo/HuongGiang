. "$PSScriptRoot\lib.ps1"
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$p.ActiveWindow.View.GotoSlide(1); Start-Sleep 1
$null = Select-RibbonTab $win 'Transitions'
$all = $win.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition)
foreach ($e in $all) { $r = $e.Current.BoundingRectangle
  if ($r.Y -gt 95 -and $r.Y -lt 200 -and $e.Current.Name -and $e.Current.ControlType.ProgrammaticName -match 'Button|MenuItem|SplitButton|ListItem|DataGrid|ComboBox|CheckBox') { "{0}|{1}|{2}" -f $e.Current.ControlType.ProgrammaticName.Replace('ControlType.',''), $e.Current.Name, $r } }
