# Click : move the mouse to a screen coordinate and left-click (optionally
# double-click). Used for controls that UI Automation cannot invoke directly.
function Click([int]$x, [int]$y, [switch]$Double) {
  [SOA.Win]::SetCursorPos($x, $y) | Out-Null
  Start-Sleep -Milliseconds 120
  [SOA.Win]::mouse_event(0x02, 0, 0, 0, 0); [SOA.Win]::mouse_event(0x04, 0, 0, 0, 0)
  if ($Double) { Start-Sleep -Milliseconds 60; [SOA.Win]::mouse_event(0x02,0,0,0,0); [SOA.Win]::mouse_event(0x04,0,0,0,0) }
}
