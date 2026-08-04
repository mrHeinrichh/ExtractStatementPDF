# WinRect : return the Contact window rectangle (Left/Top/Right/Bottom), used as
# the anchor for the grid click offsets.
function WinRect { $p=Get-Contact; $r=New-Object CF.Win+RECT; [CF.Win]::GetWindowRect($p.MainWindowHandle,[ref]$r) | Out-Null; return $r }
