. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
$pres = $p.ActivePresentation
"fonts: " + $pres.SlideMaster.Theme.ThemeFontScheme.MajorFont.Item(1).Name + "/" + $pres.SlideMaster.Theme.ThemeFontScheme.MinorFont.Item(1).Name
$p.ActiveWindow.ViewType = 3      # ppViewSlideMaster
Start-Sleep 1
$p.ActiveWindow.View.Slide = $pres.SlideMaster
Start-Sleep 1
$ph = $null
foreach ($s in $pres.SlideMaster.Shapes) { if ($s.Type -eq 14) { if ($s.PlaceholderFormat.Type -eq 2) { $ph = $s } } }
"body placeholder: " + ($ph -ne $null)
$ph.TextFrame.TextRange.ParagraphFormat.Alignment = 4     # ppAlignJustify
$ph.Select()
$null = Select-RibbonTab $win 'Home'
$j = Find-Ui $win 'Justify' 'Button'
"Justify: " + ($j -ne $null) + " " + $j.Current.BoundingRectangle
Shot 'b2-4' @($j)
