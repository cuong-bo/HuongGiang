. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
Start-Sleep 1
Click-At 38 80; Start-Sleep 2.5
Save-Shot (Rc 0 60 200 600) @((Rc 4 440 195 44)) (Join-Path $script:ShotDir 'b1-2.png')
Click-At 31 87; Start-Sleep 2
# ---- Buoc 2: nhap du lieu ----
$ws = $x.ActiveSheet
$h = 'TT','Mã hàng','Đơn giá nhập','Số lượng','Phí vận chuyển','Thành tiền','Hao phí','Đơn giá bán'
for ($i = 0; $i -lt 8; $i++) { $ws.Cells.Item(1, $i + 1).Value2 = $h[$i] }
$rows = @(@('Y001',150000,20), @('N002',70000,35), @('Y003',80000,35), @('Y004',90000,10), @('Y005',125000,13), @('N006',135000,15), @('N007',145000,45), @('N008',275000,16), @('N009',352000,17), @('N010',236000,21))
for ($r = 0; $r -lt 10; $r++) {
  $ws.Cells.Item($r + 2, 1).Value2 = $r + 1
  $ws.Cells.Item($r + 2, 2).Value2 = $rows[$r][0]
  $ws.Cells.Item($r + 2, 3).Value2 = $rows[$r][1]
  $ws.Cells.Item($r + 2, 4).Value2 = $rows[$r][2]
}
$ws.Range('A1').Select()
Shot 'b2-1' @((Get-XRect $x 'A1:H1'), (Get-XRect $x 'A2:D11'))
