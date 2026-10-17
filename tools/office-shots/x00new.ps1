. "$PSScriptRoot\lib.ps1"
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
Click-At 1500 700
while ($x.Workbooks.Count -gt 0) { $x.Workbooks.Item(1).Close($false) }
[void]$x.Workbooks.Add()
Start-Sleep 2
$x.WindowState = -4137
$x.ActiveWindow.Zoom = 100
Start-Sleep 1
"workbook moi: " + $x.Workbooks.Count
$null = Select-RibbonTab $win 'Home'
