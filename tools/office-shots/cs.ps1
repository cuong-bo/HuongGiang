param([string]$Name, [string]$Match = '.', [int]$Cols = 3)
. "$PSScriptRoot\lib.ps1"
$dir = Join-Path $script:Proj "assets\img\practical\$Name"
New-ContactSheet $dir (Join-Path $env:TEMP "cs_$Name.png") $Match $Cols 620
"ok"
