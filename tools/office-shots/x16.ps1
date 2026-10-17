. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
$win = Get-OfficeWin 'EXCEL'
Shot 'b8-3' @((Rc 467 296 66 26), (Rc 467 546 66 26), (Rc 319 421 66 26), (Rc 616 421 66 26), (Rc 216 603 100 22))
Click-At 627 750                       # OK
Start-Sleep 1.5
"xong"
