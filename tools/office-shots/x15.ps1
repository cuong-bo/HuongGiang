. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'
Shot 'b8-2' @((Rc 479 300 92 26), (Rc 289 485 508 32))
Click-At 311 250                       # tab Margins
Start-Sleep 1.5
Save-Shot (Rc 150 60 700 700) @() (Join-Path $env:TEMP 'xl_margins.png')
