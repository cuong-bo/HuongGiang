. "$PSScriptRoot\lib.ps1"
$win = Get-OfficeWin 'POWERPNT'
Wheel-At 170 700 25
Start-Sleep 1
Save-Shot (Rc 0 60 400 1000) @() (Join-Path $env:TEMP 'pp_panel2.png')
