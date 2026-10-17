. "$PSScriptRoot\lib.ps1"
$win = Get-OfficeWin 'EXCEL'
Close-Activation $win
Start-Sleep 1
$root = [Windows.Automation.AutomationElement]::RootElement
$dlg = Find-Ui $root 'Page Setup' 'Window'
"Page Setup dialog con mo: " + ($dlg -ne $null)
$x = Get-OfficeApp 'Excel.Application'
$ps = $x.ActiveSheet.PageSetup
"Landscape=" + $ps.Orientation + " Paper=" + $ps.PaperSize + " L=" + [math]::Round($ps.LeftMargin/28.3465,2) + " R=" + [math]::Round($ps.RightMargin/28.3465,2) + " T=" + [math]::Round($ps.TopMargin/28.3465,2) + " B=" + [math]::Round($ps.BottomMargin/28.3465,2) + " CenterH=" + $ps.CenterHorizontally
