. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'; $x = Get-OfficeApp 'Excel.Application'
Close-Activation $win
Shot 'b8-1' @((Rc 638 198 24 24))
Click-At 649 210
$dlg = Wait-Ui $win 'Page Setup' 'Window' 10
"dialog: " + ($dlg -ne $null)
Start-Sleep 1
Dump-Ui $dlg 80 | Where-Object { $_ -notmatch '^Text' }
Save-Shot (Rc 500 150 900 800) @() (Join-Path $env:TEMP 'xl_ps.png')
