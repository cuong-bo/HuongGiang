. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
Start-Sleep 1
# dam bao da thoat man hinh File
$back = Find-Like $win '^Back$' 'Button'
"Back button: " + ($back -ne $null)
Click-At 31 87; Start-Sleep 2
Save-Shot (Rc 0 60 400 300) @() (Join-Path $env:TEMP 'xl_check.png')
