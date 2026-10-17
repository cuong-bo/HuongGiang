. "$PSScriptRoot\lib.ps1"
$x = Get-OfficeApp 'Excel.Application'
$ws = $x.ActiveSheet; $w = $x.ActiveWindow
$r = $ws.Range('A1:H1')
"Left={0} Top={1} Width={2} Height={3}" -f $r.Left, $r.Top, $r.Width, $r.Height
"PxX(left)={0} PxY(top)={1} PxX(left+w)={2} PxY(top+h)={3}" -f $w.PointsToScreenPixelsX($r.Left), $w.PointsToScreenPixelsY($r.Top), $w.PointsToScreenPixelsX($r.Left + $r.Width), $w.PointsToScreenPixelsY($r.Top + $r.Height)
"Zoom=" + $w.Zoom + " ColA width=" + $ws.Columns.Item(1).ColumnWidth + " DPI?"
$c = $ws.Range('A2'); $cell = $c.Value2; "A2 Value2 type test:"
try { $ws.Range('A2').Value2 = [double]1; "ok double" } catch { "double loi: " + $_.Exception.Message }
try { $ws.Range('A3').Value2 = 2; "ok int" } catch { "int loi: " + $_.Exception.Message }
try { $ws.Range('A4').Formula = '3'; "ok formula" } catch { "formula loi" }
