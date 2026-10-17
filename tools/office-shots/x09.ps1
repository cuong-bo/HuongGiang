. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
$null = Select-RibbonTab $win 'Home'
$ws = $x.ActiveSheet
$bar = Rc 276 236 640 44; $nameBox = Rc 10 232 126 32
function StepShots($step, $col, $formula, $fillRange) {
  $ws.Range("${col}2").Formula = $formula
  $ws.Range("${col}2").Select()
  Shot "b$step-1" @((Get-XRect $x "${col}2"), $nameBox, $bar)
  $ws.Range($fillRange).FillDown()
  $ws.Range("${col}2").Select()
  Shot "b$step-2" @((Get-XRect $x "${col}2:${col}11"))
}
StepShots 3 'E' '=ROUND(5%*C2*D2,-3)' 'E2:E11'
StepShots 4 'F' '=C2*D2+E2' 'F2:F11'
StepShots 5 'G' '=IF(LEFT(B2,1)="N",1%*F2,2%*F2)' 'G2:G11'
StepShots 6 'H' '=ROUND((F2+G2)*110%,-3)' 'H2:H11'
$ws.Range('E2:H11').Value2 | Out-Null
"E2=" + $ws.Range('E2').Value2 + " F2=" + $ws.Range('F2').Value2 + " G2=" + $ws.Range('G2').Value2 + " H2=" + $ws.Range('H2').Value2
"H11=" + $ws.Range('H11').Value2
