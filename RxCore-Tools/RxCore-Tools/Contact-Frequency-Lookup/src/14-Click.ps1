# Click : move the mouse to a screen coordinate and left-click.
function Click([int]$x,[int]$y){ [CF.Win]::SetCursorPos($x,$y)|Out-Null; Start-Sleep -Milliseconds 120; [CF.Win]::mouse_event(0x02,0,0,0,0); [CF.Win]::mouse_event(0x04,0,0,0,0) }
