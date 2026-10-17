. "$PSScriptRoot\lib.ps1"
$p = Get-OfficeApp 'PowerPoint.Application'
$pres = $p.ActivePresentation
$lay = $pres.SlideMaster.CustomLayouts
for ($i = 1; $i -le $lay.Count; $i++) {
  $l = $lay.Item($i)
  if ($l.Name -eq 'Bao tang va anh') {
    foreach ($s in $l.Shapes) { "{0} type={1} L={2} T={3} W={4} H={5} lock={6}" -f $s.Name, $s.PlaceholderFormat.Type, [int]$s.Left, [int]$s.Top, [int]$s.Width, [int]$s.Height, $s.LockAspectRatio
      if ($s.PlaceholderFormat.Type -eq 18) { $s.LockAspectRatio = 0; $s.Left = 500; $s.Top = 115; $s.Width = 420; $s.Height = 380; "  -> sau: W={0} H={1}" -f [int]$s.Width, [int]$s.Height } }
  }
}
