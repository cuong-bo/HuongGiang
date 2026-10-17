. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
$null = Select-RibbonTab $win 'Home'
$ws = $x.ActiveSheet
[void]$ws.Cells.Clear()
$h = 'TT','Mã hàng','Đơn giá nhập','Số lượng','Phí vận chuyển','Thành tiền','Hao phí','Đơn giá bán'
$cols = 'A','B','C','D','E','F','G','H'
for ($i = 0; $i -lt 8; $i++) { $ws.Range($cols[$i] + '1').Value2 = $h[$i] }
$rows = @(@('Y001',150000,20), @('N002',70000,35), @('Y003',80000,35), @('Y004',90000,10), @('Y005',125000,13), @('N006',135000,15), @('N007',145000,45), @('N008',275000,16), @('N009',352000,17), @('N010',236000,21))
for ($r = 0; $r -lt 10; $r++) {
  $row = $r + 2
  $ws.Range("A$row").Value2 = [double]($r + 1)
  $ws.Range("B$row").Value2 = $rows[$r][0]
  $ws.Range("C$row").Value2 = [double]$rows[$r][1]
  $ws.Range("D$row").Value2 = [double]$rows[$r][2]
}
$ws.Range('A1').Select()
Shot 'b2-1' @((Get-XRect $x 'A1:H1'), (Get-XRect $x 'A2:D11'))
