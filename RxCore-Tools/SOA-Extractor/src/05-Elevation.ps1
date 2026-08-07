# ---- Elevation / UAC mismatch detection -----------------------------------
# UI Automation cannot drive an app running at a HIGHER integrity level than
# itself (Windows UIPI). If Accounting is "Run as administrator" but this tool is
# not, the window's title is readable (login check passes) yet its controls are
# invisible and it won't respond to maximize -> the confusing "pane not found".
# These helpers let us detect that and tell the user exactly what to do.

# Is THIS process elevated (running as administrator)?
function Test-SelfElevated {
  try {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object System.Security.Principal.WindowsPrincipal($id)).IsInRole(
      [System.Security.Principal.WindowsBuiltInRole]::Administrator)
  } catch { return $false }
}

# Can we read the chosen Accounting process's module path? If NOT (access denied),
# it's almost certainly running at a higher integrity level than us.
function Test-AcctAccessible {
  try { $p = Get-Acct; $null = $p.Path; return $true } catch { return $false }
}

# Returns $true if there is an elevation mismatch that will block automation.
function Test-ElevationMismatch {
  return ((-not (Test-AcctAccessible)) -and (-not (Test-SelfElevated)))
}

$ElevationHelp =
  "Accounting appears to be running as Administrator, but this tool is not." + [Environment]::NewLine +
  "Windows blocks a normal program from automating an elevated one, so its" + [Environment]::NewLine +
  "buttons/panes are invisible to this tool." + [Environment]::NewLine + [Environment]::NewLine +
  "Fix (either one):" + [Environment]::NewLine +
  "  - Restart Accounting WITHOUT 'Run as administrator' (matches this tool), OR" + [Environment]::NewLine +
  "  - Right-click SOA-Extractor.exe and 'Run as administrator'." + [Environment]::NewLine + [Environment]::NewLine +
  "Note: if you run this as administrator, a Google Drive 'G:' reference may not be" + [Environment]::NewLine +
  "visible - copy RxOffice.csv to a local folder (e.g. Documents) and browse to that."
