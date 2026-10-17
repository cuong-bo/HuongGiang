# Thu vien chup man hinh Office + khoanh do (Windows PowerShell 5.1)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms
Add-Type @"
using System; using System.Runtime.InteropServices;
public class Win32 {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
}
"@
[void][Win32]::SetProcessDPIAware()

function Get-AppWindow([int]$ProcId) {
  $cond = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $ProcId)
  [Windows.Automation.AutomationElement]::RootElement.FindFirst('Children', $cond)
}
function Find-Ui($Root, [string]$Name, [string]$Type = '') {
  $c = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name)
  if ($Type) {
    $ct = [Windows.Automation.ControlType]::($Type)
    $c2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, $ct)
    $c = New-Object Windows.Automation.AndCondition($c, $c2)
  }
  $Root.FindFirst('Descendants', $c)
}
function Get-Rect($El) { $r = $El.Current.BoundingRectangle; [System.Drawing.Rectangle]::new([int]$r.X, [int]$r.Y, [int]$r.Width, [int]$r.Height) }
function Invoke-Ui($El) {
  $p = $null
  if ($El.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$p)) { $p.Invoke(); return }
  if ($El.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern, [ref]$p)) { $p.Select(); return }
  if ($El.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$p)) { $p.Expand(); return }
  throw "Khong bam duoc: $($El.Current.Name)"
}
# Chup vung hinh chu nhat (toa do man hinh), ve khoanh do quanh cac vung $Marks, luu PNG
function Save-Shot([System.Drawing.Rectangle]$Area, [System.Drawing.Rectangle[]]$Marks, [string]$Path) {
  $bmp = New-Object System.Drawing.Bitmap($Area.Width, $Area.Height)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($Area.X, $Area.Y, 0, 0, $bmp.Size)
  $g.SmoothingMode = 'AntiAlias'
  $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 230, 20, 20), 4)
  foreach ($m in $Marks) {
    $pad = $(if ($m.Height -lt 40) { 3 } else { 6 })
    $x = $m.X - $Area.X - $pad; $y = $m.Y - $Area.Y - $pad; $w = $m.Width + 2 * $pad; $h = $m.Height + 2 * $pad
    $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
    $rad = [Math]::Min(14, [Math]::Min($w, $h) / 2)
    $d = 2 * $rad
    $gp.AddArc($x, $y, $d, $d, 180, 90); $gp.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $gp.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90); $gp.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $gp.CloseFigure(); $g.DrawPath($pen, $gp)
  }
  $dir = Split-Path $Path; if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null }
  $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
}

# ---------- cac ham cap cao ----------
$script:Proj = Split-Path (Split-Path $PSScriptRoot)
$script:ShotDir = $null
$vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
$script:Area = [System.Drawing.Rectangle]::new(0, 60, $vs.Width, $vs.Height - 60)   # bo thanh tieu de (co dong chu Product Activation)

function Set-ShotDir([string]$Name) {
  $script:ShotDir = Join-Path $script:Proj "assets\img\practical\$Name"
  New-Item -ItemType Directory -Force $script:ShotDir | Out-Null
}
function Get-Word { [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
function Get-WordWin {
  $p = Get-Process WINWORD | Select-Object -First 1
  [void][Win32]::SetForegroundWindow($p.MainWindowHandle)
  Get-AppWindow $p.Id
}
function Wait-Ui($Root, [string]$Name, [string]$Type = '', [int]$Sec = 8) {
  $end = (Get-Date).AddSeconds($Sec)
  do { $e = Find-Ui $Root $Name $Type; if ($e) { return $e }; Start-Sleep -Milliseconds 300 } while ((Get-Date) -lt $end)
  $null
}
# Bam phan tu o luong khac (hop thoai modal khong chan script)
function Invoke-Async($El) {
  $ps = [powershell]::Create()
  [void]$ps.AddScript({
    param($e)
    Add-Type -AssemblyName UIAutomationClient
    $p = $null
    if ($e.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$p)) { $p.Invoke() }
    elseif ($e.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern, [ref]$p)) { $p.Toggle() }
    elseif ($e.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern, [ref]$p)) { $p.Select() }
    elseif ($e.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$p)) { $p.Expand() }
  }).AddArgument($El)
  [void]$ps.BeginInvoke()
}
function Dump-Ui($Root, [int]$Max = 120) {
  $all = $Root.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition); $i = 0
  foreach ($e in $all) { if ($i++ -ge $Max) { break }
    "{0}|{1}|{2}|{3}" -f $e.Current.ControlType.ProgrammaticName.Replace('ControlType.', ''), $e.Current.Name, $e.Current.AutomationId, $e.Current.BoundingRectangle }
}
# Thong bao noi (toast) o goc duoi phai man hinh la vung toi mau: cho no tat truoc khi chup
function Test-Toast {
  $b = New-Object System.Drawing.Bitmap(380, 70); $g = [System.Drawing.Graphics]::FromImage($b)
  $g.CopyFromScreen(1500, 990, 0, 0, $b.Size); $g.Dispose()
  $dark = 0; $n = 0
  for ($x = 0; $x -lt 380; $x += 10) { for ($y = 0; $y -lt 70; $y += 10) { $c = $b.GetPixel($x, $y); $n++; if (($c.R + $c.G + $c.B) -lt 300) { $dark++ } } }
  $b.Dispose(); ($dark / $n) -gt 0.15
}
function Wait-NoToast { for ($i = 0; $i -lt 15; $i++) { if (-not (Test-Toast)) { return }; Start-Sleep -Seconds 2 }; Write-Warning 'thong bao noi van con' }
function Shot([string]$Name, $Marks = @()) {
  $rects = @()
  foreach ($m in $Marks) {
    if ($m -is [System.Drawing.Rectangle]) { $rects += $m } elseif ($m) { $rects += (Get-Rect $m) }
  }
  try { (Get-Word).ScreenRefresh() } catch {}
  [void][Mouse]::SetCursorPos(1300, 640)   # dua chuot ra khoi nut de khong hien tooltip
  Start-Sleep -Milliseconds 1200
  Wait-NoToast
  Save-Shot $script:Area $rects (Join-Path $script:ShotDir "$Name.png")
  "shot: $Name"
}
function Select-RibbonTab($Win, [string]$Name) {
  $c1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name)
  $c2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)
  $tabs = $Win.FindAll('Descendants', (New-Object Windows.Automation.AndCondition($c1, $c2)))
  $t = $tabs[$tabs.Count - 1]      # tab ngu canh (Table) nam sau tab thuong
  Invoke-Ui $t; Start-Sleep -Milliseconds 800
  $t
}

Add-Type @"
using System; using System.Runtime.InteropServices;
public class Mouse {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
}
"@
function Click-At([int]$X, [int]$Y, [switch]$Right) {
  [void][Mouse]::SetCursorPos($X, $Y); Start-Sleep -Milliseconds 150
  if ($Right) { [Mouse]::mouse_event(0x08, 0, 0, 0, [UIntPtr]::Zero); [Mouse]::mouse_event(0x10, 0, 0, 0, [UIntPtr]::Zero) }
  else { [Mouse]::mouse_event(0x02, 0, 0, 0, [UIntPtr]::Zero); [Mouse]::mouse_event(0x04, 0, 0, 0, [UIntPtr]::Zero) }
  Start-Sleep -Milliseconds 500
}
function Click-Ui($El, [switch]$Right) { $r = $El.Current.BoundingRectangle; Click-At ([int]($r.X + $r.Width / 2)) ([int]($r.Y + $r.Height / 2)) -Right:$Right }
function Close-Activation($Win) {
  foreach ($i in 1..3) {
    $dlg = Find-Ui $Win 'Microsoft Office Activation Wizard' 'Window'
    if (-not $dlg) { return }
    $b = Find-Ui $dlg 'Close' 'Button'
    if ($b) { Invoke-Ui $b; Start-Sleep 1 } else { return }
  }
}

function Set-Value($El, [string]$Text) {
  $p = $null
  if ($El.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$p)) { $p.SetValue($Text); return }
  $ed = $El.FindFirst('Descendants', (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Edit)))
  if ($ed -and $ed.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$p)) { $p.SetValue($Text); return }
  throw "Khong dat duoc gia tri cho $($El.Current.Name)"
}

# Tim phan tu theo mau ten (regex) - dung khi ten co ky tu trang dac biet
function Find-Like($Root, [string]$Pattern, [string]$Type = '') {
  $cond = [Windows.Automation.Condition]::TrueCondition
  if ($Type) { $cond = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::($Type)) }
  foreach ($e in $Root.FindAll('Descendants', $cond)) { if ($e.Current.Name -match $Pattern) { return $e } }
  $null
}

function Rc($x, $y, $w, $h) { [System.Drawing.Rectangle]::new($x, $y, $w, $h) }

# Them mot doan van vao cuoi tai lieu. $Tabs: mang @(cm, canh 0=L 1=C 2=R, leader 0=none 1=dots)
function Add-Para($Word, [string]$Text, $Tabs = @(), [int]$Align = 0, [switch]$First) {
  $doc = $Word.ActiveDocument
  if (-not $First) { $doc.Content.InsertParagraphAfter() }
  $p = $doc.Paragraphs.Last
  $p.Range.Font.Reset(); $p.Range.ParagraphFormat.Reset()
  $p.TabStops.ClearAll()
  foreach ($t in $Tabs) { [void]$p.TabStops.Add($Word.CentimetersToPoints($t[0]), $t[1], $t[2]) }
  $p.Alignment = $Align
  $p.Range.InsertBefore($Text)
  $p
}
# In dam / in nghieng cho doan con $Sub trong doan $Para
function Format-Sub($Word, $Para, [string]$Sub, [string]$Kind = 'b') {
  $doc = $Word.ActiveDocument
  $txt = $Para.Range.Text; $i = $txt.IndexOf($Sub)
  if ($i -lt 0) { return }
  $s = $Para.Range.Start + $i
  $r = $doc.Range($s, $s + $Sub.Length)
  if ($Kind -eq 'b') { $r.Font.Bold = 1 } else { $r.Font.Italic = 1 }
}
# Toa do man hinh cua mot Range (pixel vat ly)
function Get-RangeRect($Word, $Range) {
  $l = 0; $t = 0; $wd = 0; $h = 0
  $Word.ActiveWindow.GetPoint([ref]$l, [ref]$t, [ref]$wd, [ref]$h, $Range)
  [System.Drawing.Rectangle]::new($l, $t, $wd, $h)
}

# Hop chu nhat bao tron cac doan tu so $From den $To (1-based)
function Get-ParaRect($Word, [int]$From, [int]$To) {
  $doc = $Word.ActiveDocument
  $a = Get-RangeRect $Word $doc.Paragraphs.Item($From).Range
  $b = Get-RangeRect $Word $doc.Paragraphs.Item($To).Range
  $x = [Math]::Min($a.X, $b.X); $y = [Math]::Min($a.Y, $b.Y)
  $r = [Math]::Max($a.Right, $b.Right); $bt = [Math]::Max($a.Bottom, $b.Bottom)
  [System.Drawing.Rectangle]::new($x, $y, $r - $x, $bt - $y)
}
function Go-End($Word) { [void]$Word.Selection.EndKey(6) }

# An/hien cac doan tu $From den het tai lieu (de anh chup khop voi buoc dang huong dan)
function Set-HiddenFrom($Word, [int]$From, [bool]$Hidden) {
  $doc = $Word.ActiveDocument
  if ($From -gt $doc.Paragraphs.Count) { return }
  $r = $doc.Range($doc.Paragraphs.Item($From).Range.Start, $doc.Content.End - 1)
  $r.Font.Hidden = [int]$Hidden * -1
}

# Dat gia tri vao o Edit (theo AutomationId) cua hop thoai, roi chuyen focus de Word ghi nhan
function Set-EditById($Dlg, [string]$Id, [string]$Text) {
  $c = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
  $e = $Dlg.FindFirst('Descendants', $c)
  if (-not $e) { throw "Khong thay o Edit id=$Id" }
  $e.SetFocus(); Start-Sleep -Milliseconds 200
  $e.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Text)
  Start-Sleep -Milliseconds 300
  $e
}

# ---------- Excel / PowerPoint ----------
function Get-OfficeApp([string]$ProgId) { [Runtime.InteropServices.Marshal]::GetActiveObject($ProgId) }
function Get-OfficeWin([string]$ProcName) {
  $p = Get-Process $ProcName | Select-Object -First 1
  [void][Win32]::SetForegroundWindow($p.MainWindowHandle)
  Get-AppWindow $p.Id
}
function Stop-OfficeProc([string]$ProcName) { Get-Process $ProcName -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.Id -Force } }

# Toa do man hinh cua mot vung o Excel (pixel vat ly). Excel tra ve diem theo don vi logic nen phai nhan he so DPI.
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetDpiForSystem();' -Name Dpi -Namespace Win32x
function Get-XRect($Xl, [string]$Addr) {
  $f = ([Win32x.Dpi]::GetDpiForSystem() / 96.0) * (4.0 / 3.0)
  $r = $Xl.ActiveSheet.Range($Addr); $w = $Xl.ActiveWindow
  $ox = [double]$w.PointsToScreenPixelsX(0); $oy = [double]$w.PointsToScreenPixelsY(0)
  $l = $ox + ([double]$w.PointsToScreenPixelsX($r.Left) - $ox) * $f
  $t = $oy + ([double]$w.PointsToScreenPixelsY($r.Top) - $oy) * $f
  $rr = $ox + ([double]$w.PointsToScreenPixelsX($r.Left + $r.Width) - $ox) * $f
  $b = $oy + ([double]$w.PointsToScreenPixelsY($r.Top + $r.Height) - $oy) * $f
  [System.Drawing.Rectangle]::new([int]$l, [int]$t, [int]($rr - $l), [int]($b - $t))
}

function Wheel-At([int]$X, [int]$Y, [int]$Ticks) {
  Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, int d, UIntPtr e);' -Name MW -Namespace Win32w -ErrorAction SilentlyContinue
  [void][Mouse]::SetCursorPos($X, $Y); Start-Sleep -Milliseconds 200
  for ($i = 0; $i -lt [Math]::Abs($Ticks); $i++) { [Win32w.MW]::mouse_event(0x0800, 0, 0, [int](-120 * [Math]::Sign($Ticks)), [UIntPtr]::Zero); Start-Sleep -Milliseconds 120 }
}

# Toa do man hinh cua mot shape tren slide dang hien o PowerPoint (pixel vat ly)
function Get-PRect($App, $Shape) {
  $f = ([Win32x.Dpi]::GetDpiForSystem() / 96.0) * (4.0 / 3.0)
  $w = $App.ActiveWindow
  $x0 = [double]$w.PointsToScreenPixelsX($Shape.Left); $y0 = [double]$w.PointsToScreenPixelsY($Shape.Top)
  $x1 = [double]$w.PointsToScreenPixelsX($Shape.Left + $Shape.Width); $y1 = [double]$w.PointsToScreenPixelsY($Shape.Top + $Shape.Height)
  [System.Drawing.Rectangle]::new([int]$x0, [int]$y0, [int]($x1 - $x0), [int]($y1 - $y0))
}
