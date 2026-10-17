. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
# hieu ung khac nhau: Fade (10) o tieu de slide 3, Fly In (2) o hinh slide 4, Zoom (53) o hinh slide 5
$e1 = $pres.Slides.Item(3).TimeLine.MainSequence.AddEffect($pres.Slides.Item(3).Shapes.Item(1), 10, 0, 1)
$e2 = $pres.Slides.Item(4).TimeLine.MainSequence.AddEffect($pres.Slides.Item(4).Shapes.Item(3), 2, 0, 1)
$e3 = $pres.Slides.Item(5).TimeLine.MainSequence.AddEffect($pres.Slides.Item(5).Shapes.Item(3), 53, 0, 1)
"effects: {0} {1} {2}" -f $e1.EffectType, $e2.EffectType, $e3.EffectType
$p.ActiveWindow.View.GotoSlide(3); Start-Sleep 1
$pres.Slides.Item(3).Shapes.Item(1).Select()
Start-Sleep 1
$null = Select-RibbonTab $win 'Animations'
$all = $win.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition)
foreach ($e in $all) { $r = $e.Current.BoundingRectangle
  if ($r.Y -gt 95 -and $r.Y -lt 200 -and $e.Current.Name -and $e.Current.ControlType.ProgrammaticName -match 'Button|MenuItem|SplitButton|ListItem|DataGrid|ComboBox') { "{0}|{1}|{2}" -f $e.Current.ControlType.ProgrammaticName.Replace('ControlType.',''), $e.Current.Name, $r } }
Save-Shot (Rc 0 60 1920 200) @() (Join-Path $env:TEMP 'pp_anim.png')
