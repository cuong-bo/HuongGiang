. "$PSScriptRoot\lib.ps1"
$root = [Windows.Automation.AutomationElement]::RootElement
$dlg = Find-Ui $root 'Page Setup' 'Window'
"dialog: " + ($dlg -ne $null) + " " + $dlg.Current.BoundingRectangle
Dump-Ui $dlg 120 | Where-Object { $_ -notmatch '^Text' }
Save-Shot (Rc 0 60 1920 1020) @() (Join-Path $env:TEMP 'xl_ps_full.png')
