. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
$null = Select-RibbonTab $win 'Home'
$ws = $x.ActiveSheet
# --- font ---
$ws.Cells.Font.Name = 'Times New Roman'; $ws.Cells.Font.Size = 13
$ws.Cells.Select()
Shot 'b7-1' @((Rc 158 112 210 30))
# --- vien: menu Borders ---
$ws.Range('A1:H11').Select()
Click-At 315 166; Start-Sleep 1.5
Shot 'b7-2' @((Rc 277 150 50 32), (Rc 281 407 282 31))
Click-At 360 422; Start-Sleep 1          # All Borders
# --- can chinh tieu de ---
$ws.Range('A1:H1').HorizontalAlignment = -4108
$ws.Range('A1:H1').VerticalAlignment = -4108
$ws.Range('A1:H1').WrapText = $true
$ws.Range('A1:H1').Font.Bold = $true
$ws.Range('A2:B11').HorizontalAlignment = -4108
$ws.Range('D2:D11').HorizontalAlignment = -4108
$ws.Range('A1:H1').Select()
Shot 'b7-3' @((Rc 160 150 30 30), (Rc 483 109 30 30), (Rc 483 149 30 30), (Rc 630 111 105 30))
# --- dinh dang so ---
$ws.Range('C2:C11,E2:H11').NumberFormat = '#,##0'
$ws.Range('C2:C11,E2:H11').Select()
Shot 'b7-4' @((Rc 909 150 30 30), (Rc 976 150 30 30))
# --- tu gian cot: menu Format ---
$ws.Range('A1:H11').Select()
Click-At 1405 160; Start-Sleep 1.5
Shot 'b7-5' @((Rc 1379 102 53 96), (Rc 1382 350 218 32))
Click-At 1491 366; Start-Sleep 1         # AutoFit Column Width
$ws.Range('A1').Select()
Shot 'b7-6' @((Get-XRect $x 'A1:H11'))
# --- Buoc 8: trang in ---
$ws.PageSetup.PaperSize = 9; $ws.PageSetup.Orientation = 2
$ws.PageSetup.TopMargin = $x.CentimetersToPoints(2.5); $ws.PageSetup.BottomMargin = $x.CentimetersToPoints(2.5)
$ws.PageSetup.LeftMargin = $x.CentimetersToPoints(2.5); $ws.PageSetup.RightMargin = $x.CentimetersToPoints(1.5)
$ws.PageSetup.CenterHorizontally = $true; $ws.PageSetup.CenterVertically = $false
$null = Select-RibbonTab $win 'Page Layout'
Shot 'b8-1' @((Rc 638 198 24 24))
Click-At 649 210
$root = [Windows.Automation.AutomationElement]::RootElement
$dlg = Wait-Ui $root 'Page Setup' 'Window' 10
Start-Sleep 1
Shot 'b8-2' @((Rc 479 300 92 26), (Rc 289 485 508 32))
Click-At 311 250; Start-Sleep 1.5       # tab Margins
Shot 'b8-3' @((Rc 467 296 66 26), (Rc 467 546 66 26), (Rc 319 421 66 26), (Rc 616 421 66 26), (Rc 216 603 100 22))
Click-At 627 750; Start-Sleep 1.5       # OK
$ps = $ws.PageSetup
"Landscape=" + $ps.Orientation + " Paper=" + $ps.PaperSize + " CenterH=" + $ps.CenterHorizontally
