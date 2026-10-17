. "$PSScriptRoot\lib.ps1"
Set-ShotDir 'de1-c3'
$win = Get-OfficeWin 'POWERPNT'; $p = Get-OfficeApp 'PowerPoint.Application'
Close-Activation $win
$pres = $p.ActivePresentation
# b4-1: menu Layout dang mo
Shot 'b4-1' @((Rc 214 102 100 30), (Rc 356 176 125 100), (Rc 491 545 125 98))
Click-At 107 80; Start-Sleep 1.5          # dong menu
function LayoutByName($n) { $l = $pres.SlideMaster.CustomLayouts; for ($i = 1; $i -le $l.Count; $i++) { if ($l.Item($i).Name -eq $n) { return $l.Item($i) } } }
$lGt = LayoutByName 'Gioi thieu'; $lTc = LayoutByName 'Title and Content'; $lBt = LayoutByName 'Bao tang va anh'
# slide 1
$s1 = $pres.Slides.Item(1); $s1.CustomLayout = $lGt
$s1.Shapes.Title.TextFrame.TextRange.Text = 'BẢO TÀNG Ở HÀ NỘI'
foreach ($sh in $s1.Shapes) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 2) { $sh.TextFrame.TextRange.Text = 'Giới thiệu' } }
# slide 2
$s2 = $pres.Slides.AddSlide(2, $lTc)
$s2.Shapes.Title.TextFrame.TextRange.Text = 'Nội dung trình bày'
foreach ($sh in $s2.Shapes) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 2 -or $sh.PlaceholderFormat.Type -eq 7) { $sh.TextFrame.TextRange.Text = "Bảo tàng Lịch sử Hà Nội`rBảo tàng Hà Nội`rBảo tàng Hồ Chí Minh" } }
# slide 3-5
$titles = 'Bảo tàng Lịch sử Hà Nội', 'Bảo tàng Hà Nội', 'Bảo tàng Hồ Chí Minh'
$bodies = @(
  "Tọa lạc ở quận Hoàn Kiếm`rLưu giữ khoảng 200.000 hiện vật, tư liệu`rTrưng bày lịch sử văn hóa Việt Nam",
  "Thủ đô có địa điểm trưng bày riêng`rGiới thiệu nét độc đáo của nền văn hóa`rMảnh đất hơn một ngàn năm tuổi",
  "Lưu trữ hiện vật gắn liền với Chủ tịch Hồ Chí Minh`rCuộc đời người lãnh tụ vĩ đại của dân tộc`rĐiểm đến của nhiều du khách")
for ($i = 0; $i -lt 3; $i++) {
  $s = $pres.Slides.AddSlide(3 + $i, $lBt)
  $s.Shapes.Title.TextFrame.TextRange.Text = $titles[$i]
  $pic = $null
  foreach ($sh in @($s.Shapes)) {
    if ($sh.Type -eq 14) { $t = $sh.PlaceholderFormat.Type
      if ($t -eq 2 -or $t -eq 7) { $sh.TextFrame.TextRange.Text = $bodies[$i] }
      if ($t -eq 18) { $pic = $sh } }
  }
  if ($pic) { $l = $pic.Left; $tp = $pic.Top; $w = $pic.Width; $h = $pic.Height; $pic.Delete()
    $null = $s.Shapes.AddPicture("$env:TEMP\bt$($i + 1).png", 0, -1, $l, $tp, $w, $h) }
}
"so slide: " + $pres.Slides.Count
foreach ($s in $pres.Slides) { "S{0}: layout={1} shapes={2}" -f $s.SlideIndex, $s.CustomLayout.Name, $s.Shapes.Count }
