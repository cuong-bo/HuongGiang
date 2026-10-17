. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
$null = Select-RibbonTab $win 'Home'
$ws = $x.ActiveSheet
# b7-2: menu Borders dang mo
Shot 'b7-2' @((Rc 277 150 50 32), (Rc 281 407 282 31))
Click-At 360 422                       # All Borders
Start-Sleep 1
# b7-3: can chinh tieu de
$ws.Range('A1:H1').HorizontalAlignment = -4108
$ws.Range('A1:H1').VerticalAlignment = -4108
$ws.Range('A1:H1').WrapText = $true
$ws.Range('A1:H1').Font.Bold = $true
$ws.Range('A2:B11').HorizontalAlignment = -4108
$ws.Range('D2:D11').HorizontalAlignment = -4108
$ws.Range('A1:H1').Select()
Shot 'b7-3' @((Rc 160 150 30 30), (Rc 483 109 30 30), (Rc 483 149 30 30), (Rc 630 111 105 30))
# b7-4: dinh dang so
$ws.Range('C2:C11,E2:H11').NumberFormat = '#,##0'
$ws.Range('C2:C11,E2:H11').Select()
Shot 'b7-4' @((Rc 909 150 30 30), (Rc 976 150 30 30))
# mo menu Format de xem vi tri AutoFit
$ws.Range('A1:H11').Select()
Click-At 1405 160
Start-Sleep 1.5
Save-Shot (Rc 1100 60 820 700) @() (Join-Path $env:TEMP 'xl_format.png')
