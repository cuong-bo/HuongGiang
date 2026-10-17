. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'
Close-Activation $win
Start-Sleep 1
$dlg = Find-Ui $win 'Microsoft Office Activation Wizard' 'Window'
"wizard con lai: " + ($dlg -ne $null)
Click-At 38 80
Start-Sleep 2.5
Save-Shot (Rc 0 60 200 700) @() (Join-Path $script:ShotDir '_nav.png')
