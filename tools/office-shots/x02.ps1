. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c2'
Remove-Item (Join-Path $script:ShotDir 'start.png') -ErrorAction SilentlyContinue
$x = Get-OfficeApp 'Excel.Application'
Shot 'b1-1' @((Rc 20 64 40 32))
Click-At 38 80
Start-Sleep 2
Save-Shot (Rc 0 60 200 700) @() (Join-Path $script:ShotDir '_nav.png')
