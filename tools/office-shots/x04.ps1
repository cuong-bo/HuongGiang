. "$PSScriptRoot\lib.ps1"
$p = Get-Process EXCEL | Select-Object -First 1
$root = [Windows.Automation.AutomationElement]::RootElement
$c = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $p.Id)
foreach ($w in $root.FindAll('Children', $c)) { "{0} | {1} | {2} | off={3}" -f $w.Current.ClassName, $w.Current.Name, $w.Current.BoundingRectangle, $w.Current.IsOffscreen
  $wc = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Window)
  foreach ($d in $w.FindAll('Descendants', $wc)) { "   dlg: {0} | {1} | off={2}" -f $d.Current.ClassName, $d.Current.Name, $d.Current.IsOffscreen } }
"Cursor: " + [System.Windows.Forms.Cursor]::Position
