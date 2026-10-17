. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
$null = Select-RibbonTab $win 'Home'
$ws = $x.ActiveSheet
$ws.Cells.Font.Name = 'Times New Roman'; $ws.Cells.Font.Size = 13
$ws.Cells.Select()
Shot 'b7-1' @((Rc 158 112 210 30))
$ws.Range('A1:H11').Select()
Click-At 315 166                       # mui ten cua nut Borders
Start-Sleep 1.5
Save-Shot (Rc 0 60 700 700) @() (Join-Path $env:TEMP 'xl_borders.png')
