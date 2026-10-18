# Cac khoi chup dung lai cho Word, Excel, PowerPoint (dung chung voi lib.ps1).
# Moi ham nhan ten anh (vd 'b2-1') va chup + khoanh do. Toa do co dinh dung cho man hinh 1920x1080 (do phan giai cua may nay).
. "$PSScriptRoot\lib.ps1"

$script:Procs = @{ Word = 'WINWORD'; Excel = 'EXCEL'; PowerPoint = 'POWERPNT' }
$script:Root = [Windows.Automation.AutomationElement]::RootElement

function Get-Win([string]$Kind) { Get-OfficeWin $script:Procs[$Kind] }
function Calm([string]$Kind) { $w = Get-Win $Kind; Close-Activation $w; $w }

# Mo ung dung moi (COM), phong to, dong hop thoai kich hoat
function Start-Office([string]$Kind) {
  Get-Process $script:Procs[$Kind] -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.Id -Force }
  Start-Sleep 2
  $app = New-Object -ComObject "$Kind.Application"
  $app.Visible = $(if ($Kind -eq 'PowerPoint') { -1 } else { $true })
  switch ($Kind) { 'Word' { [void]$app.Documents.Add() } 'Excel' { [void]$app.Workbooks.Add() } 'PowerPoint' { [void]$app.Presentations.Add(-1) } }
  Start-Sleep 6
  $win = Get-Win $Kind; Close-Activation $win
  $state = switch ($Kind) { 'Word' { 1 } 'Excel' { -4137 } 'PowerPoint' { 3 } }
  for ($i = 0; $i -lt 10; $i++) { try { $app.WindowState = $state; break } catch { Start-Sleep 2; Close-Activation (Get-Win $Kind) } }
  Start-Sleep 2
  $win = Get-Win $Kind; Close-Activation $win
  if ($Kind -eq 'Word') { Click-At 410 249; Start-Sleep 1 }          # dong khung Navigation neu co
  $null = Select-RibbonTab $win 'Home'
  $app
}
function Stop-Office([string]$Kind) {
  try {
    $app = Get-OfficeApp "$Kind.Application"
    switch ($Kind) {
      'Word' { while ($app.Documents.Count -gt 0) { $app.Documents.Item(1).Close(0) } }
      'Excel' { while ($app.Workbooks.Count -gt 0) { $app.Workbooks.Item(1).Close($false) } }
      'PowerPoint' { while ($app.Presentations.Count -gt 0) { $app.Presentations.Item(1).Close() } }
    }
  } catch {}
  Start-Sleep 2
  Get-Process $script:Procs[$Kind] -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.Id -Force }
  Start-Sleep 2
}

# ---------- dung chung ----------
function B-FileOpen([string]$Kind, [string]$N1, [string]$N2, [string]$Which = 'Save As') {
  $win = Calm $Kind
  $null = Select-RibbonTab $win 'Home'
  Shot $N1 @((Rc 20 64 40 32))
  Click-At 38 80; Start-Sleep 2.5
  $y = if ($Which -eq 'Save') { 389 } else { 440 }
  Save-Shot (Rc 0 60 200 600) @((Rc 4 $y 195 44)) (Join-Path $script:ShotDir "$N2.png")
  Click-At 31 87; Start-Sleep 2
}
function B-SaveMenu([string]$Kind, [string]$N) {
  $win = Calm $Kind
  Click-At 38 80; Start-Sleep 2.5
  Save-Shot (Rc 0 60 200 600) @((Rc 4 389 195 44)) (Join-Path $script:ShotDir "$N.png")
  Click-At 31 87; Start-Sleep 2
}

# ---------- Word ----------
function BW-PageSetup($W, [string]$NLaunch, [string]$NMargins, [string]$NPaper) {
  $win = Calm 'Word'; $doc = $W.ActiveDocument; $ps = $doc.PageSetup
  $ps.PaperSize = 7; $ps.Orientation = 0
  $ps.TopMargin = $W.CentimetersToPoints(2.5); $ps.BottomMargin = $W.CentimetersToPoints(2)
  $ps.LeftMargin = $W.CentimetersToPoints(3.5); $ps.RightMargin = $W.CentimetersToPoints(1.5)
  $null = Select-RibbonTab $win 'Layout'
  Shot $NLaunch @((Rc 424 200 22 22))
  Click-At 434 210; Start-Sleep 2.5
  Click-At 748 230; Start-Sleep 1.5
  Shot $NMargins @((Rc 806 280 129 26), (Rc 1068 280 129 26), (Rc 806 310 129 26), (Rc 1068 310 129 26), (Rc 736 404 71 58))
  Click-At 829 230; Start-Sleep 1.5
  Shot $NPaper @((Rc 736 280 208 28))
  Click-At 1027 851; Start-Sleep 1.5
  $null = Select-RibbonTab $win 'Home'
}
function BW-NormalStyle($W, [string]$N1, [string]$N2) {
  $win = Calm 'Word'; $doc = $W.ActiveDocument
  $n = $doc.Styles.Item('Normal'); $n.Font.Name = 'Times New Roman'; $n.Font.Size = 13
  $n.ParagraphFormat.SpaceBefore = 0; $n.ParagraphFormat.SpaceAfter = 0; $n.ParagraphFormat.LineSpacingRule = 0
  $null = Select-RibbonTab $win 'Home'
  $normal = Find-Like $win '^\W*Normal$' 'ListItem'
  Click-Ui $normal -Right; Start-Sleep 1
  $mod = Wait-Ui $script:Root 'Modify...' '' 5
  Shot $N1 @($normal, $mod)
  Invoke-Async $mod
  [void](Wait-Ui $win 'Modify Style' 'Window' 10); Start-Sleep 1
  Shot $N2 @((Rc 657 405 141 24), (Rc 803 405 48 24), (Rc 643 829 113 30))
  Click-At 1092 844; Start-Sleep 1.5
}
# Chen bang bang hop thoai Insert Table (chua dien noi dung)
function BW-InsertTable($W, [int]$Cols, [int]$Rows, [string]$NMenu, [string]$NDlg) {
  $win = Calm 'Word'
  $null = Select-RibbonTab $win 'Insert'
  $tbl = Find-Ui $win 'Table' 'MenuItem'
  $tbl.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand(); Start-Sleep 1.5
  $it = Wait-Ui $script:Root 'Insert Table...' '' 5
  if ($NMenu) { Shot $NMenu @($tbl, $it) }
  Invoke-Async $it
  $d = Wait-Ui $script:Root 'Insert Table' 'Window' 10; Start-Sleep 1
  $null = Set-EditById $d '19' "$Cols"; $null = Set-EditById $d '21' "$Rows"
  if ($NDlg) { Shot $NDlg @((Rc 988 420 111 26), (Rc 988 450 111 26)) }
  Click-At 915 670; Start-Sleep 2
}
# Menu Borders tab Home: 'No Border' hoac 'All Borders'
function BW-Borders([string]$N, [string]$Item = 'No Border') {
  $win = Calm 'Word'
  $null = Select-RibbonTab $win 'Home'
  Click-At 790 167; Start-Sleep 1.5
  $y = if ($Item -eq 'No Border') { 323 } else { 358 }
  Shot $N @((Rc 752 150 50 34), (Rc 755 $y 222 34))
  Click-At 860 ($y + 17); Start-Sleep 1.5
}
# Chen mot doan vao cuoi tai lieu va tra ve doan do
function BW-AddPara($W, [string]$Text, [int]$Align = 0, [bool]$Bold = $false) {
  $doc = $W.ActiveDocument
  $doc.Content.InsertParagraphAfter()
  $p = $doc.Paragraphs.Last
  $p.Range.Font.Reset(); $p.Range.ParagraphFormat.Reset()
  $p.Range.InsertBefore($Text)
  $p.Alignment = $Align; $p.Range.Font.Bold = [int]$Bold * -1
  $p
}
function BW-FullPage($W, [string]$N) {
  $win = Calm 'Word'
  [void]$W.Selection.EndKey(6)
  $W.ActiveWindow.View.Zoom.PageFit = 1; Start-Sleep 1
  Shot $N @((Rc 8 1058 82 22))
  $W.ActiveWindow.View.Zoom.Percentage = 100
}

# ---------- Excel ----------
function BX-Formula($X, [string]$N1, [string]$N2, [string]$Cell, [string]$Formula, [string]$FillRange, [string]$ResultRange) {
  $win = Calm 'Excel'; $ws = $X.ActiveSheet
  $null = Select-RibbonTab $win 'Home'
  $bar = Rc 276 236 640 44; $nameBox = Rc 10 232 126 32
  $ws.Range($Cell).Formula = $Formula
  $ws.Range($Cell).Select()
  Shot $N1 @((Get-XRect $X $Cell), $nameBox, $bar)
  if ($FillRange) { $ws.Range($FillRange).FillDown() }
  $ws.Range($Cell).Select()
  Shot $N2 @((Get-XRect $X $(if ($ResultRange) { $ResultRange } else { $Cell })))
}
# Phan dinh dang: font, vien (All Borders), tu gian cot, anh ket qua
function BX-FormatBlock($X, [string]$Table, [string]$NFont, [string]$NBorders, [string]$NAutofit, [string]$NFinal) {
  $win = Calm 'Excel'; $ws = $X.ActiveSheet
  $null = Select-RibbonTab $win 'Home'
  $ws.Cells.Font.Name = 'Times New Roman'; $ws.Cells.Font.Size = 13
  $ws.Cells.Select()
  Shot $NFont @((Rc 158 112 210 30))
  $ws.Range($Table).Select()
  Click-At 315 166; Start-Sleep 1.5
  Shot $NBorders @((Rc 277 150 50 32), (Rc 281 407 282 31))
  Click-At 360 422; Start-Sleep 1
  $ws.Range($Table).Select()
  Click-At 1405 160; Start-Sleep 1.5
  Shot $NAutofit @((Rc 1379 102 53 96), (Rc 1382 350 218 32))
  Click-At 1491 366; Start-Sleep 1
  $ws.Range('A1').Select()
  Shot $NFinal @((Get-XRect $X $Table))
}
function BX-PageSetup($X, [string]$NLaunch, [string]$NPage, [string]$NMargins) {
  $win = Calm 'Excel'; $ws = $X.ActiveSheet
  $ws.PageSetup.PaperSize = 9; $ws.PageSetup.Orientation = 2
  $ws.PageSetup.TopMargin = $X.CentimetersToPoints(2.5); $ws.PageSetup.BottomMargin = $X.CentimetersToPoints(2.5)
  $ws.PageSetup.LeftMargin = $X.CentimetersToPoints(2.5); $ws.PageSetup.RightMargin = $X.CentimetersToPoints(1.5)
  $ws.PageSetup.CenterHorizontally = $true; $ws.PageSetup.CenterVertically = $false
  $null = Select-RibbonTab $win 'Page Layout'
  Shot $NLaunch @((Rc 638 198 24 24))
  Click-At 649 210; Start-Sleep 2.5
  Click-At 230 250; Start-Sleep 1.5
  Shot $NPage @((Rc 479 300 92 26), (Rc 289 485 508 32))
  Click-At 311 250; Start-Sleep 1.5
  Shot $NMargins @((Rc 467 296 66 26), (Rc 467 546 66 26), (Rc 319 421 66 26), (Rc 616 421 66 26), (Rc 216 603 100 22))
  Click-At 627 750; Start-Sleep 1.5
  $null = Select-RibbonTab $win 'Home'
}
# Du lieu: mang 2 chieu (hang) -> ghi tu o $Top ('A1')
function BX-Fill($X, [string]$Top, $Rows) {
  $ws = $X.ActiveSheet; $r0 = $ws.Range($Top).Row; $c0 = $ws.Range($Top).Column
  for ($r = 0; $r -lt $Rows.Count; $r++) { for ($c = 0; $c -lt $Rows[$r].Count; $c++) {
    $v = $Rows[$r][$c]
    if ($null -eq $v -or "$v" -eq '') { continue }
    $cell = $ws.Cells.Item(($r0 + $r), ($c0 + $c))
    if ($v -is [int] -or $v -is [double]) { $cell.Value2 = [double]$v } else { $cell.Value2 = "$v" }
  } }
}

# ---------- PowerPoint ----------
function BP-MasterFonts($P, [string]$NView, [string]$NMenu, [string]$NDlg) {
  $win = Calm 'PowerPoint'; $pres = $P.ActivePresentation
  $null = Select-RibbonTab $win 'View'
  $sm = Find-Ui $win 'Slide Master' 'Button'
  Shot $NView @($sm)
  Click-Ui $sm; Start-Sleep 3
  $fs = $pres.SlideMaster.Theme.ThemeFontScheme
  $fs.MajorFont.Item(1).Name = 'Arial'; $fs.MinorFont.Item(1).Name = 'Arial'
  $null = Select-RibbonTab $win 'Slide Master'
  $fonts = Find-Ui $win 'Fonts' 'MenuItem'
  $fonts.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand(); Start-Sleep 1.5
  $cf = Wait-Ui $script:Root 'Customize Fonts...' '' 5
  Shot $NMenu @($fonts, $cf)
  Invoke-Async $cf
  $dlg = Wait-Ui $script:Root 'Create New Theme Fonts' 'Window' 10; Start-Sleep 1
  Set-Value (Find-Ui $dlg 'Name' 'Edit') 'Arial'
  Start-Sleep 1
  Shot $NDlg @((Rc 682 493 266 28), (Rc 682 550 266 28), (Rc 744 590 494 24), (Rc 1051 618 88 30))
  Invoke-Ui (Find-Ui $dlg 'Save' 'Button'); Start-Sleep 2
}
function BP-Justify($P, [string]$N) {
  $win = Calm 'PowerPoint'; $pres = $P.ActivePresentation
  Wheel-At 170 700 -25; Start-Sleep 1
  Click-At 150 305; Start-Sleep 2
  foreach ($s in $pres.SlideMaster.Shapes) { if ($s.Type -eq 14) { if ($s.PlaceholderFormat.Type -eq 2) { $s.TextFrame.TextRange.ParagraphFormat.Alignment = 4 } } }
  Click-At 1100 560; Start-Sleep 1.5
  $null = Select-RibbonTab $win 'Home'
  Shot $N @((Rc 803 152 30 30))
}
# Tao layout moi bang Insert Layout + Rename
function BP-Layout($P, [string]$Name, [string]$NInsert, [string]$NRename, [string]$NPlaceholder) {
  $win = Calm 'PowerPoint'; $pres = $P.ActivePresentation
  $null = Select-RibbonTab $win 'Slide Master'
  Wheel-At 170 700 -25; Start-Sleep 1
  Click-At 170 440; Start-Sleep 1.5
  $il = Find-Ui $win 'Insert Layout' 'Button'
  if ($NInsert) { Shot $NInsert @($il) }
  Click-Ui $il; Start-Sleep 2
  Click-Ui (Find-Ui $win 'Rename' 'Button'); Start-Sleep 2
  $dlg = Find-Ui $script:Root 'Rename Layout' 'Window'
  Set-Value (Find-Ui $dlg 'Layout name' 'Edit') $Name
  Start-Sleep 1
  if ($NRename) { Shot $NRename @((Rc 861 536 198 24), (Rc 872 564 88 30)) }
  Invoke-Ui (Find-Ui $dlg 'Rename' 'Button'); Start-Sleep 2
  if ($NPlaceholder) {
    $null = Select-RibbonTab $win 'Slide Master'
    Click-At 380 175; Start-Sleep 1.5
    Shot $NPlaceholder @((Rc 331 102 98 97))
    Click-At 130 80; Start-Sleep 1
  }
}
function BP-PlaceBox($S, $l, $t, $w, $h) { $S.LockAspectRatio = 0; $S.Left = $l; $S.Top = $t; $S.Width = $w; $S.Height = $h }
function BP-LayoutByName($P, [string]$N) { $l = $P.ActivePresentation.SlideMaster.CustomLayouts; for ($i = 1; $i -le $l.Count; $i++) { if ($l.Item($i).Name -eq $N) { return $l.Item($i) } } }
function BP-ShowSlide($P, [int]$N) { $P.ActiveWindow.View.GotoSlide($N); Start-Sleep 1.5 }
function BP-LayoutGallery($P, [string]$N, $Marks) {
  $win = Calm 'PowerPoint'
  BP-ShowSlide $P 1
  $null = Select-RibbonTab $win 'Home'
  $lay = Find-Ui $win 'Layout' 'MenuItem'
  $lay.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand(); Start-Sleep 2
  Shot $N (@((Rc 214 102 100 30)) + $Marks)
  Click-At 107 80; Start-Sleep 1.5
}
function BP-SlideNumbers($P, [string]$N1, [string]$N2, [string]$N3) {
  $win = Calm 'PowerPoint'; $pres = $P.ActivePresentation
  BP-ShowSlide $P 2
  $null = Select-RibbonTab $win 'Insert'
  $sn = Find-Ui $win 'Slide Number' 'Button'
  Shot $N1 @($sn)
  Click-Ui $sn; Start-Sleep 2.5
  $dlg = Find-Ui $script:Root 'Header and Footer' 'Window'
  $cb1 = Find-Ui $dlg 'Slide number' 'CheckBox'; $cb2 = Find-Ui $dlg "Don't show on title slide" 'CheckBox'
  foreach ($c in $cb1, $cb2) { $t = $c.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern); if ($t.Current.ToggleState -ne 1) { $t.Toggle() } }
  Start-Sleep 1
  $apply = Find-Ui $dlg 'Apply to All' 'Button'
  Shot $N2 @($cb1, $cb2, $apply)
  Invoke-Ui $apply; Start-Sleep 2.5
  $num = $null
  foreach ($sh in $pres.Slides.Item(2).Shapes) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 13) { $num = $sh } }
  $null = Select-RibbonTab $win 'Home'
  $m = @(); if ($num) { $m += (Get-PRect $P $num) }
  Shot $N3 $m
}
# $Plan: mang @(so-slide, ten-anh-hoac-'', x-o-thu-vien)  (Fade 311, Push 414, Wipe 517, Split 620)
function BP-Transitions($P, $Plan) {
  $win = Calm 'PowerPoint'
  $null = Select-RibbonTab $win 'Transitions'
  foreach ($it in $Plan) {
    BP-ShowSlide $P $it[0]
    Click-At ($it[2] + 51) 149; Start-Sleep 2
    if ($it[1]) { Shot $it[1] @((Rc $it[2] 114 103 70)); Start-Sleep 1 }
  }
}
function BP-Final($P, [string]$NShow, [string]$NSorter) {
  $win = Calm 'PowerPoint'
  BP-ShowSlide $P 1
  $null = Select-RibbonTab $win 'Slide Show'
  Shot $NShow @((Rc 25 102 72 96))
  Click-At 1548 1066; Start-Sleep 2.5
  Shot $NSorter @()
  Click-At 1498 1066; Start-Sleep 2
}
# Ve anh minh hoa don gian (khong dung anh ben ngoai). $Kind: 0 bien/dao, 1 hang/toi, 2 pho/den long, 3 toa nha
function New-Illustration([string]$Path, [string]$Label, [int]$Kind = 3, $Rgb = @(180, 60, 40)) {
  $bmp = New-Object System.Drawing.Bitmap(840, 760); $g = [System.Drawing.Graphics]::FromImage($bmp); $g.SmoothingMode = 'AntiAlias'
  $c = $Rgb
  $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush([System.Drawing.Point]::new(0,0), [System.Drawing.Point]::new(0,760), [System.Drawing.Color]::FromArgb(255,235,240,250), [System.Drawing.Color]::FromArgb(255,$c[0],$c[1],$c[2]))
  $g.FillRectangle($br, 0, 0, 840, 760)
  $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(230,250,250,250))
  $acc = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255,$c[0],$c[1],$c[2]))
  switch ($Kind) {
    0 { $g.FillRectangle($acc, 0, 480, 840, 280); foreach ($r in @(@(100,330,200,230), @(330,270,180,290), @(540,350,230,210))) { $g.FillEllipse($white, $r[0], $r[1], $r[2], $r[3]) } }
    1 { $g.Clear([System.Drawing.Color]::FromArgb(255,25,30,28)); $g.FillEllipse($acc, 170, 60, 500, 560) }
    2 { $g.FillRectangle($white, 0, 380, 840, 380); foreach ($k in 0..5) { $g.FillEllipse($acc, 60 + $k * 130, 120 + (($k % 2) * 60), 70, 90) } }
    default { $g.FillRectangle($white, 150, 300, 540, 330); $g.FillPolygon($white, [System.Drawing.Point[]]@([System.Drawing.Point]::new(120,300), [System.Drawing.Point]::new(420,160), [System.Drawing.Point]::new(720,300))); foreach ($k in 0..4) { $g.FillRectangle($acc, 190 + $k * 100, 350, 40, 230) } }
  }
  $font = New-Object System.Drawing.Font('Arial', 30, [System.Drawing.FontStyle]::Bold)
  $fmt = New-Object System.Drawing.StringFormat; $fmt.Alignment = 'Center'
  $g.DrawString($Label, $font, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]::new(20, 660, 800, 90), $fmt)
  $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose()
}
