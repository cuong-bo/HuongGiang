. "$PSScriptRoot\lib.ps1"
$x = New-Object -ComObject Excel.Application
$x.Visible = $true
[void]$x.Workbooks.Add()
Start-Sleep 4
$x.WindowState = -4137   # xlMaximized
Start-Sleep 2
$win = Get-OfficeWin 'EXCEL'
Close-Activation $win
Start-Sleep 1
$tabs = $win.FindAll('Descendants', (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)))
"Tabs: " + (($tabs | ForEach-Object { $_.Current.Name }) -join ' | ')
"Decimal sep: " + $x.International(3) + " list sep: " + $x.International(5)
Set-ShotDir 'de1-c2'
Shot 'start' @()
