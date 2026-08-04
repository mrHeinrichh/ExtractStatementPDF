# Normalize-Window : move Contact to the primary monitor and MAXIMIZE it, so the
# grid offsets in 01-Settings are correct and consistent every run.
function Normalize-Window {
  $p = Get-Contact
  [CF.Win]::ShowWindow($p.MainWindowHandle,9) | Out-Null                                  # restore
  [CF.Win]::SetWindowPos($p.MainWindowHandle,[IntPtr]::Zero,0,0,1400,900,0x0040) | Out-Null # move to 0,0
  Start-Sleep -Milliseconds 300
  [CF.Win]::ShowWindow($p.MainWindowHandle,3) | Out-Null                                  # maximize
  Start-Sleep -Milliseconds 400
  [CF.Win]::SetForegroundWindow($p.MainWindowHandle) | Out-Null
  Start-Sleep -Milliseconds 400
  return $p
}
