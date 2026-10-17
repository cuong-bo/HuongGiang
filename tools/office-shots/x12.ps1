. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
$null = Select-RibbonTab $win 'Home'
$ws = $x.ActiveSheet
Shot 'b7-5' @((Rc 1379 102 53 96), (Rc 1382 350 218 32))
Click-At 1491 366                      # AutoFit Column Width
Start-Sleep 1
$ws.Range('A1').Select()
Shot 'b7-6' @((Get-XRect $x 'A1:H11'))
# Buoc 8: tab Page Layout
$ws.PageSetup.PaperSize = 9; $ws.PageSetup.Orientation = 2
$ws.PageSetup.TopMargin = $x.CentimetersToPoints(2.5); $ws.PageSetup.BottomMargin = $x.CentimetersToPoints(2.5)
$ws.PageSetup.LeftMargin = $x.CentimetersToPoints(2.5); $ws.PageSetup.RightMargin = $x.CentimetersToPoints(1.5)
$ws.PageSetup.CenterHorizontally = $true; $ws.PageSetup.CenterVertically = $false
$null = Select-RibbonTab $win 'Page Layout'
Save-Shot (Rc 0 60 1100 200) @() (Join-Path $env:TEMP 'xl_pl.png')
