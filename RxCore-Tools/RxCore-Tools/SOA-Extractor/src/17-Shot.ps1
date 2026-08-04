# Shot : save a PNG screenshot of the Accounting window to $ShotDir\<name>.png.
# Best-effort (never throws) - used for troubleshooting each step.
function Shot([string]$name) {
  try {
    $p = Get-Acct; $r = New-Object SOA.Win+RECT
    [SOA.Win]::GetWindowRect($p.MainWindowHandle, [ref]$r) | Out-Null
    $w = $r.Right-$r.Left; $h = $r.Bottom-$r.Top; if ($w -le 0 -or $h -le 0) { return }
    $bmp = New-Object System.Drawing.Bitmap $w, $h; $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, (New-Object System.Drawing.Size $w, $h))
    $bmp.Save((Join-Path $ShotDir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
  } catch {}
}
