# ===================== src\00-Header.ps1 =====================
<#
  SOA Extractor  -  assembled source
  ==================================
  This tool is written as one function per file under this src\ folder.
  build.ps1 concatenates every src\*.ps1 (in filename order) into a single
  script and compiles it to SOA-Extractor.exe with ps2exe.

  Load order is controlled by the numeric filename prefix:
     00  header + assemblies         (this file)
     01  settings
     02  Win32 interop + shortcuts
     10-52  one function per file
     99  main program body (runs last)

  What it does: for every .xls in a chosen folder, reads the customer + dates,
  drives Accounting.exe's "Accounting Print > Statement" screen, sets the
  customer's known Invoice Frequency, generates the report and exports it to a
  .csv in a SEPARATE "<folder> - SOA CSV" output folder.

  Requires: Accounting.exe already OPEN and LOGGED IN (user hfabros); Excel installed.
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes


# ===================== src\01-Settings.ps1 =====================
# ---- Settings (safe to tweak) ---------------------------------------------
# Fallback order used only when a customer's frequency is NOT in the lookup table.
# These strings must match the dropdown items exactly: Monthly, Daily, Weekly, Bi-Weekly.
$FrequencyOrder = @('Monthly', 'Daily', 'Weekly', 'Bi-Weekly')

# $true  = re-generate and replace CSVs that already exist in the output folder.
# $false = skip any .xls that already has a .csv in the output folder.
$Overwrite = $false

# Where per-step screenshots and the run log are written.
$ShotDir = Join-Path $env:TEMP 'soa_shots'


# ===================== src\02-Interop.ps1 =====================
# ---- Win32 interop + UI Automation shortcuts ------------------------------
# Native window/mouse functions used to focus the app, click at screen
# coordinates, and capture screenshots.
$sig = @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool BringWindowToTop(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
[DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr hWnd, out RECT lpRect);
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
'@
if (-not ([System.Management.Automation.PSTypeName]'SOA.Win').Type) {
  Add-Type -MemberDefinition $sig -Name Win -Namespace SOA -PassThru | Out-Null
}

# Shorthand for the UI Automation entry type used throughout.
$AE = [System.Windows.Automation.AutomationElement]


# ===================== src\03-Ui-Dialogs.ps1 =====================
# ---- Simple GUI dialogs (message box + file browser) ----------------------

# Show-Message : pop a message box (used for the login warning).
function Show-Message([string]$text, [string]$title = 'SOA Extractor') {
  [System.Windows.Forms.MessageBox]::Show($text, $title,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}

# Browse-File : open-file dialog for the reference CSV. Starts in Documents.
# Returns the chosen path, or $null if cancelled.
function Browse-File([string]$title) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $title
  $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
  $docs = [Environment]::GetFolderPath('MyDocuments')
  if ($docs -and (Test-Path $docs)) { $dlg.InitialDirectory = $docs }
  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  return $dlg.FileName
}


# ===================== src\04-Choose-Acct.ps1 =====================
# ---- Choose which Accounting instance to drive ----------------------------
# Lets you pick from the running Accounting windows (PID + title + exe path) when
# more than one is open, and remembers the choice ($script:AcctPid) so every other
# function targets that exact window. This is what lets it work on a PC that has a
# differently-located window or more than one Accounting open.

$script:AcctPid = $null

# Return the running Accounting windows as @{ Pid; Path; Title; Handle }.
function Get-AcctCandidates {
  $list = New-Object System.Collections.Generic.List[object]
  $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object {
    $_.MainWindowHandle -ne 0 -and (
      $_.ProcessName -like 'Accounting*' -or ($_.MainWindowTitle -like '*Accounting( Version*')
    )
  }
  foreach ($p in $procs) {
    $path = ''
    try { $path = $p.Path } catch {}
    $list.Add([pscustomobject]@{ Pid=$p.Id; Path=$path; Title=$p.MainWindowTitle; Handle=$p.MainWindowHandle })
  }
  return $list
}

# Show the picker (only when >1). Sets $script:AcctPid. Returns the chosen
# candidate, or $null if none running / cancelled.
function Choose-Acct {
  $cands = @(Get-AcctCandidates)
  if ($cands.Count -eq 0) { return $null }
  if ($cands.Count -eq 1) { $script:AcctPid = $cands[0].Pid; return $cands[0] }

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Choose the Accounting window to use"
  $form.Size = New-Object System.Drawing.Size(760, 300)
  $form.StartPosition = "CenterScreen"; $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.Text = "More than one Accounting is open - pick the one to extract from:"
  $label.Location = '12,10'; $label.AutoSize = $true; $form.Controls.Add($label)

  $lb = New-Object System.Windows.Forms.ListBox
  $lb.Location = '12,35'; $lb.Size = New-Object System.Drawing.Size(720, 180)
  $lb.Anchor = "Top,Left,Right,Bottom"; $lb.HorizontalScrollbar = $true
  foreach ($c in $cands) { [void]$lb.Items.Add(("PID {0}   |   {1}   |   {2}" -f $c.Pid, $c.Title, $c.Path)) }
  $lb.SelectedIndex = 0
  $form.Controls.Add($lb)

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "Use this one"; $ok.Location = '556,222'; $ok.Size = '95,30'
  $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK; $ok.Anchor = "Bottom,Right"
  $form.Controls.Add($ok); $form.AcceptButton = $ok

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"; $cancel.Location = '657,222'; $cancel.Size = '75,30'
  $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $cancel.Anchor = "Bottom,Right"
  $form.Controls.Add($cancel); $form.CancelButton = $cancel

  if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  $chosen = $cands[$lb.SelectedIndex]
  $script:AcctPid = $chosen.Pid
  return $chosen
}


# ===================== src\05-Elevation.ps1 =====================
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


# ===================== src\10-Pick-Folder.ps1 =====================
# Pick-Folder : show the paste-a-path dialog (with a Browse... button).
# Returns the chosen folder path, or $null if cancelled.
function Pick-Folder {
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "SOA Extractor"
  $form.Size = New-Object System.Drawing.Size(620, 190)
  $form.StartPosition = "CenterScreen"
  $form.FormBorderStyle = "FixedDialog"
  $form.MaximizeBox = $false; $form.MinimizeBox = $false
  $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.Text = "Paste the folder path that contains the .xls files:"
  $label.Location = New-Object System.Drawing.Point(15, 15)
  $label.AutoSize = $true
  $form.Controls.Add($label)

  $box = New-Object System.Windows.Forms.TextBox
  $box.Location = New-Object System.Drawing.Point(15, 45)
  $box.Size = New-Object System.Drawing.Size(490, 25)
  $box.Anchor = "Top,Left,Right"
  $box.Add_KeyDown({ if ($_.Control -and $_.KeyCode -eq 'A') { $box.SelectAll(); $_.SuppressKeyPress = $true } })
  $form.Controls.Add($box)

  $browse = New-Object System.Windows.Forms.Button
  $browse.Text = "Browse..."
  $browse.Location = New-Object System.Drawing.Point(510, 44)
  $browse.Size = New-Object System.Drawing.Size(75, 25)
  $browse.Anchor = "Top,Right"
  $browse.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "Select the folder that contains the .xls files"
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $box.Text = $fb.SelectedPath }
  })
  $form.Controls.Add($browse)

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "OK"; $ok.Location = New-Object System.Drawing.Point(410, 100)
  $ok.Size = New-Object System.Drawing.Size(80, 30)
  $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
  $ok.Anchor = "Bottom,Right"
  $form.Controls.Add($ok); $form.AcceptButton = $ok

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"; $cancel.Location = New-Object System.Drawing.Point(500, 100)
  $cancel.Size = New-Object System.Drawing.Size(80, 30)
  $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $cancel.Anchor = "Bottom,Right"
  $form.Controls.Add($cancel); $form.CancelButton = $cancel

  $form.Add_Shown({ $form.Activate(); $box.Focus() })
  $res = $form.ShowDialog()
  if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  return $box.Text.Trim().Trim('"')
}


# ===================== src\11-Get-Acct.ps1 =====================
# Get-Acct : return the Accounting process to drive. Prefers the one chosen via
# Choose-Acct ($script:AcctPid); otherwise the first running "Accounting" process.
function Get-Acct {
  if ($script:AcctPid) {
    $p = Get-Process -Id $script:AcctPid -ErrorAction SilentlyContinue
    if ($p) { return $p }
  }
  $p = Get-Process Accounting -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $p) { throw "Accounting.exe is not running. Open it and log in first." }
  return $p
}


# ===================== src\12-Focus-Acct.ps1 =====================
# Focus-Acct : bring Accounting to the front and MAXIMIZE it, so all screen
# coordinates used later (toolbar/export offsets) are deterministic.
function Focus-Acct {
  $p = Get-Acct
  [SOA.Win]::ShowWindow($p.MainWindowHandle, 3)       | Out-Null   # SW_MAXIMIZE
  [SOA.Win]::SetForegroundWindow($p.MainWindowHandle) | Out-Null
  [SOA.Win]::BringWindowToTop($p.MainWindowHandle)    | Out-Null
  Start-Sleep -Milliseconds 400
  return $p
}


# ===================== src\13-Reset-State.ps1 =====================
# Reset-State : close any leftover Report window and Statement dialog so the
# next customer starts from a clean screen.
function Reset-State {
  for ($k = 0; $k -lt 3; $k++) {
    $rep = Get-ReportWindow
    if (-not $rep) { break }
    try { $rep.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch {}
    Start-Sleep -Milliseconds 600
  }
  $d = Get-Dialog
  if ($d) {
    try { $d.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch {}
    Start-Sleep -Milliseconds 400
  }
}


# ===================== src\14-Click.ps1 =====================
# Click : move the mouse to a screen coordinate and left-click (optionally
# double-click). Used for controls that UI Automation cannot invoke directly.
function Click([int]$x, [int]$y, [switch]$Double) {
  [SOA.Win]::SetCursorPos($x, $y) | Out-Null
  Start-Sleep -Milliseconds 120
  [SOA.Win]::mouse_event(0x02, 0, 0, 0, 0); [SOA.Win]::mouse_event(0x04, 0, 0, 0, 0)
  if ($Double) { Start-Sleep -Milliseconds 60; [SOA.Win]::mouse_event(0x02,0,0,0,0); [SOA.Win]::mouse_event(0x04,0,0,0,0) }
}


# ===================== src\15-Send.ps1 =====================
# Send : send raw keystrokes to the focused control (SendKeys syntax, e.g. '{ENTER}').
function Send([string]$k) { [System.Windows.Forms.SendKeys]::SendWait($k) }


# ===================== src\16-Send-Literal.ps1 =====================
# Send-Literal : type text literally, escaping SendKeys' special characters
# ( + ^ % ~ ( ) { } [ ] ) so names/dates are entered exactly as given.
function Send-Literal([string]$t) { Send ($t -replace '([+^%~(){}\[\]])', '{$1}') }


# ===================== src\17-Shot.ps1 =====================
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


# ===================== src\18-Assert-Login.ps1 =====================
# Assert-AcctLoggedIn : verify Accounting.exe is running AND logged in.
#
# When logged in, one of the app's top-level windows is titled
# "Accounting( Version ... ) Company: <co> User: <name>". At the login screen no
# window contains "User:". We scan ALL of the process's top-level windows rather
# than just MainWindowTitle, because MainWindowTitle can read empty transiently
# (e.g. while a child window/dialog has focus) and give a false "not logged in".
#
# Shows a message box and returns $false if not ready; otherwise $true.

# Titles of every top-level window of the CHOSEN Accounting process (via Get-Acct).
function Get-AcctWindowTitles {
  $proc = $null
  try { $proc = Get-Acct } catch { return @() }
  $titles = New-Object System.Collections.Generic.List[string]
  $desktop = $AE::RootElement
  try {
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, [int]$proc.Id)
    $wins = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
    foreach ($w in $wins) { $titles.Add([string]$w.Current.Name) }
  } catch {}
  if ($proc.MainWindowTitle) { $titles.Add([string]$proc.MainWindowTitle) }
  return $titles
}

function Assert-AcctLoggedIn {
  $running = Get-Process Accounting -ErrorAction SilentlyContinue
  if (-not $running -and -not $script:AcctPid) {
    Show-Message "Accounting is not open.`n`nPlease open Accounting.exe and log in (user hfabros), then run this again."
    return $false
  }
  $titles = Get-AcctWindowTitles
  if (@($titles | Where-Object { $_ -match 'User:' }).Count -gt 0) { return $true }
  Show-Message "You are not logged in to Accounting.`n`nPlease log in on the Accounting window (user hfabros), then run this again."
  return $false
}


# ===================== src\20-ConvertTo-MDY.ps1 =====================
# ConvertTo-MDY : parse a statement date like "Mar. 02, ' 26" into "3/2/2026",
# the format Accounting's date fields expect.
function ConvertTo-MDY([string]$txt) {
  $t = ($txt -replace "'","" -replace "\.","" -replace ",","" -replace "\s+"," ").Trim()
  $parts = $t.Split(' '); if ($parts.Count -lt 3) { throw "Cannot parse date '$txt'" }
  $map = @{Jan=1;Feb=2;Mar=3;Apr=4;May=5;Jun=6;Jul=7;Aug=8;Sep=9;Oct=10;Nov=11;Dec=12}
  $m = $map[$parts[0].Substring(0,3)]; $d = [int]$parts[1]; $y = [int]$parts[2]
  if ($y -lt 100) { $y += 2000 }; return "$m/$d/$y"
}


# ===================== src\21-Read-XlsMeta.ps1 =====================
# Read-XlsMeta : open an .xls (read-only, via Excel COM) and return the
# customer name (cell E9) and the AS-OF From/To dates (E18 / H18) as M/D/YYYY.
function Read-XlsMeta([string]$path) {
  $x = New-Object -ComObject Excel.Application; $x.Visible=$false; $x.DisplayAlerts=$false
  try {
    $wb = $x.Workbooks.Open($path, [Type]::Missing, $true); $ws = $wb.Sheets.Item(1)
    $m = [ordered]@{
      Customer = ([string]$ws.Cells.Item(9,5).Text).Trim()
      From = ConvertTo-MDY ([string]$ws.Cells.Item(18,6).Text)
      To   = ConvertTo-MDY ([string]$ws.Cells.Item(18,8).Text)
    }
    $wb.Close($false) | Out-Null; return $m
  } finally { $x.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
}


# ===================== src\22-Get-Candidates.ps1 =====================
# Get-Candidates : progressively shorter forms of a customer name, used for fuzzy
# matching against the reference table when the exact name isn't found. Full name
# first, then drop trailing words, then trim letters off the first word.
function Get-Candidates([string]$name) {
  $cands = New-Object System.Collections.Generic.List[string]
  $n = ($name -replace '\s+',' ').Trim()
  if ($n) { $cands.Add($n) }
  $words = $n.Split(' ')
  for ($i = $words.Count - 1; $i -ge 1; $i--) {
    $sub = (($words[0..($i-1)] -join ' ').TrimEnd(' ','-',',','.','&')).Trim()
    if ($sub.Length -ge 3 -and -not $cands.Contains($sub)) { $cands.Add($sub) }
  }
  $first = $words[0]
  for ($L = $first.Length - 1; $L -ge 4; $L--) {
    $sub = $first.Substring(0,$L)
    if (-not $cands.Contains($sub)) { $cands.Add($sub) }
  }
  return $cands
}


# ===================== src\23-WordWrap.ps1 =====================
# WordWrap : normalize a KeyOf() value to a word-bounded form for safe fuzzy
# matching. Non-alphanumerics become spaces and the whole thing is wrapped in
# spaces, so .Contains(" ESCA ") matches the WORD "ESCA" only - never the "ESCA"
# buried inside "LESCANO". Used by Resolve-Frequency.
function WordWrap([string]$k) {
  if (-not $k) { return ' ' }
  return ' ' + ((($k -replace '[^A-Z0-9]',' ') -replace '\s+',' ').Trim()) + ' '
}

# Tokenize : split a name into UPPERCASE alphanumeric words (punctuation dropped),
# used for word-level prefix matching between AR names and RxOffice names.
#   "ABALOS GUILLERMO OPTICAL" -> @('ABALOS','GUILLERMO','OPTICAL')
function Tokenize([string]$s) {
  if (-not $s) { return @() }
  return @((($s.ToUpperInvariant() -replace '[^A-Z0-9]',' ') -split '\s+') | Where-Object { $_ -ne '' })
}

# Test-Prefix : is $short the leading run of words of $long? (both string[])
function Test-Prefix($short, $long) {
  if ($short.Count -eq 0 -or $short.Count -gt $long.Count) { return $false }
  for ($i = 0; $i -lt $short.Count; $i++) { if ($short[$i] -ne $long[$i]) { return $false } }
  return $true
}

# TightKey : strip EVERYTHING except letters/digits and upper-case, so different
# spacing/punctuation/casing collapse to one key. This is how an AR name matches
# the reference's ARName column:
#   "ABALOS GUILLERMO OPTICAL"  ->  "ABALOSGUILLERMOOPTICAL"
#   "AbalosGuillermoOptical"    ->  "ABALOSGUILLERMOOPTICAL"
function TightKey([string]$s) {
  if (-not $s) { return '' }
  return ($s -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
}


# ===================== src\30-Get-Dialog.ps1 =====================
# Get-Dialog : return the "Statement options" window element, or $null if closed.
function Get-Dialog {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Statement options")
  return $AE::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
}


# ===================== src\31-Open-StatementDialog.ps1 =====================
# Open-StatementDialog : make sure the Accounting Print pane is showing, then
# double-click the Statement tile's ICON (~25px above its label) and wait for the
# "Statement options" dialog. Returns the dialog element.
#
# Robust across machines/monitors:
#   * clicks the LEFTMOST "Accounting Print" sidebar item relative to THIS window
#     (not a hardcoded screen-left assumption), and retries;
#   * if the pane still can't be found, throws a diagnostic listing what it DID see
#     (window rect, left-column labels, custom panes) so a failing PC is easy to debug.
function Open-StatementDialog {
  $p = Focus-Acct
  $root = $AE::FromHandle($p.MainWindowHandle)
  $wr = New-Object SOA.Win+RECT
  [SOA.Win]::GetWindowRect($p.MainWindowHandle, [ref]$wr) | Out-Null

  # Click the "Accounting Print" sidebar item (leftmost one within this window),
  # then wait for the PrintAccountingReports pane. Retry a few times.
  $par = $null
  for ($attempt = 0; $attempt -lt 4 -and -not $par; $attempt++) {
    $apCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Accounting Print")
    $aps = @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $apCond) |
             Where-Object { $_.Current.ClassName -eq 'TextBlock' -and $_.Current.BoundingRectangle.Width -gt 0 } |
             Sort-Object { $_.Current.BoundingRectangle.X })
    if ($aps.Count -gt 0) {
      $b = $aps[0].Current.BoundingRectangle
      Click ([int]($b.X + $b.Width/2)) ([int]($b.Y + $b.Height/2))
    }
    for ($t = 0; $t -lt 10; $t++) {
      $parCond = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, "PrintAccountingReports")
      $par = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $parCond)
      if ($par) { break }
      Start-Sleep -Milliseconds 300
    }
  }

  if (-not $par) {
    # Build a diagnostic of what this window actually exposes.
    $labels = @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                 (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text))) |
               Where-Object { $_.Current.BoundingRectangle.X -lt ($wr.Left + 320) -and $_.Current.Name } |
               ForEach-Object { $_.Current.Name } | Select-Object -Unique -First 25)
    $customs = @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                 (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Custom))) |
               ForEach-Object { $_.Current.ClassName } | Where-Object { $_ } | Select-Object -Unique)
    $elevNote = if (Test-ElevationMismatch) {
      "`n  *** ELEVATION MISMATCH: Accounting is elevated but this tool is not - run this tool as administrator, or restart Accounting without admin. ***"
    } else {
      "  (this tool elevated=$(Test-SelfElevated); Accounting readable=$(Test-AcctAccessible))"
    }
    throw ("Accounting Print pane not found.`n" +
           "  Window rect: $($wr.Left),$($wr.Top)-$($wr.Right),$($wr.Bottom)`n" +
           "  Left-column labels seen: " + ($labels -join ', ') + "`n" +
           "  Custom panes seen: " + ($customs -join ', ') + "`n" +
           $elevNote + "`n" +
           "  -> Make sure the chosen Accounting window is logged in and on the General > Accounting Print tab.")
  }

  $stCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Statement")
  $st = $par.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $stCond)
  if (-not $st) { throw "'Statement' tile not found inside the Accounting Print pane." }
  $r = $st.Current.BoundingRectangle; $cx = [int]($r.X + $r.Width/2); $iy = [int]($r.Y - 25)
  for ($i = 0; $i -lt 4; $i++) {
    Click $cx $iy -Double
    for ($t = 0; $t -lt 15; $t++) { $d = Get-Dialog; if ($d) { return $d }; Start-Sleep -Milliseconds 200 }
    Start-Sleep -Milliseconds 400
  }
  throw "Statement options dialog did not open."
}


# ===================== src\32-Field-XY.ps1 =====================
# Field-XY : the dialog has no automatable child controls, so each field is
# located by a fixed offset from the dialog window's top-left corner.
# Returns @(screenX, screenY) for the named field/button.
function Field-XY([string]$name) {
  $d = Get-Dialog; if (-not $d) { throw "dialog not open" }
  $r = $d.Current.BoundingRectangle; $dx = [int]$r.X; $dy = [int]$r.Y
  switch ($name) {
    'Customers'        { return @(($dx+240),($dy+43)) }
    'InvoiceFrequency' { return @(($dx+200),($dy+104)) }
    'From'             { return @(($dx+200),($dy+122)) }
    'To'               { return @(($dx+200),($dy+147)) }
    'OK'               { return @(($dx+247),($dy+407)) }
    'No'               { return @(($dx+328),($dy+407)) }
    default            { throw "unknown field $name" }
  }
}


# ===================== src\33-Set-Customer.ps1 =====================
# Set-Customer : type the customer name into the Customers box, wait for the
# autocomplete suggestion, then press TAB to commit it (binds "SOA-#|NAME").
function Set-Customer([string]$name) {
  $xy = Field-XY 'Customers'; Click $xy[0] $xy[1]; Start-Sleep -Milliseconds 250
  Send '^a'; Start-Sleep -Milliseconds 80; Send '{DEL}'; Start-Sleep -Milliseconds 150
  Send-Literal $name; Start-Sleep -Milliseconds 1500; Send '{TAB}'; Start-Sleep -Milliseconds 500
}


# ===================== src\34-Set-DateField.ps1 =====================
# Set-DateField : clear a date field ('From' or 'To') and type an M/D/YYYY value.
function Set-DateField([string]$name, [string]$value) {
  $xy = Field-XY $name; Click $xy[0] $xy[1]; Start-Sleep -Milliseconds 200
  Send '^a'; Start-Sleep -Milliseconds 80; Send '{DEL}'; Start-Sleep -Milliseconds 80
  Send-Literal $value; Start-Sleep -Milliseconds 200; Send '{TAB}'; Start-Sleep -Milliseconds 200
}


# ===================== src\35-Freq-Index.ps1 =====================
# Freq-Index : map a frequency name to its position in the Invoice Frequency
# dropdown (0=Monthly, 1=Daily, 2=Weekly, 3=Bi-Weekly). Returns -1 if unknown.
# Accepts spelling variants from the Contact table (e.g. "Bi-Weekly"/"BiWeekly").
function Freq-Index([string]$value) {
  switch (($value -replace '\s','').ToUpperInvariant()) {
    'MONTHLY'     { return 0 }
    'DAILY'       { return 1 }
    'WEEKLY'      { return 2 }
    'BI-WEEKLY'   { return 3 }
    'BIWEEKLY'    { return 3 }
    'FORTNIGHTLY' { return 3 }
    default       { return -1 }
  }
}


# ===================== src\36-Set-Frequency.ps1 =====================
# Set-Frequency : choose an Invoice Frequency deterministically by keyboard.
# Open the list, go to the top (Monthly), then arrow DOWN to the target index.
# (Typing the value as text left the field blank for "Bi-Weekly" - hence index nav.)
function Set-Frequency([string]$value) {
  $idx = Freq-Index $value
  if ($idx -lt 0) { return }   # unknown -> leave the dialog default (Monthly)
  $xy = Field-XY 'InvoiceFrequency'; Click $xy[0] $xy[1]; Start-Sleep -Milliseconds 300
  Send '%{DOWN}'; Start-Sleep -Milliseconds 500          # Alt+Down opens the dropdown
  Send ('{UP}' * 6); Start-Sleep -Milliseconds 200        # force selection to the top (Monthly)
  if ($idx -gt 0) { Send ('{DOWN}' * $idx); Start-Sleep -Milliseconds 200 }
  Send '{ENTER}'; Start-Sleep -Milliseconds 300
}


# ===================== src\37-Click-OK.ps1 =====================
# Click-OK : press the dialog's OK button to generate the report.
function Click-OK { $xy = Field-XY 'OK'; Click $xy[0] $xy[1] }


# ===================== src\38-Close-Dialog.ps1 =====================
# Close-Dialog : close the Statement dialog (via the window Close pattern, or
# the "No" button as a fallback).
function Close-Dialog {
  $d = Get-Dialog
  if ($d) { try { $d.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch { $xy=Field-XY 'No'; Click $xy[0] $xy[1] }; Start-Sleep -Milliseconds 500 }
}


# ===================== src\40-Get-ReportWindow.ps1 =====================
# Get-ReportWindow : return the "Report" viewer window element, or $null.
function Get-ReportWindow {
  $root = $AE::FromHandle((Get-Acct).MainWindowHandle)
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Report")
  return $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
}


# ===================== src\41-Close-Report.ps1 =====================
# Close-Report : close the Report viewer window if open.
function Close-Report {
  $rep = Get-ReportWindow
  if ($rep) { try { $rep.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch {}; Start-Sleep -Milliseconds 600 }
}


# ===================== src\42-Export-Csv-FromReport.ps1 =====================
# Export-Csv-FromReport : from the open Report viewer, click the export-format
# dropdown, choose "CSV (comma delimited)", then drive the Save As dialog to save
# to $targetCsv. Returns a status string: saved / overwritten / skipped-existing /
# no-report / no-saveas / save-failed.
#
# The report toolbar isn't automatable, so the two toolbar clicks are fixed
# offsets from the MAXIMIZED window's top-left:
#   export-format dropdown  = window-topleft + (440, 86)
#   "CSV (comma delimited)" = window-topleft + (475, 138)   (2nd item in the menu)
# The Save As dialog itself IS a standard dialog, so its fields are found via UIA.
function Export-Csv-FromReport([string]$targetCsv) {
  $p = Get-Acct
  $rep = Get-ReportWindow; if (-not $rep) { return 'no-report' }
  $wr = New-Object SOA.Win+RECT
  [SOA.Win]::GetWindowRect($p.MainWindowHandle, [ref]$wr) | Out-Null
  $wl = $wr.Left; $wt = $wr.Top
  Click ($wl + 440) ($wt + 86);  Start-Sleep -Milliseconds 800; Shot 'export_menu'
  Click ($wl + 475) ($wt + 138); Start-Sleep -Milliseconds 1200

  $root = $AE::RootElement
  $saCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Save As")
  $sa = $null
  for ($t=0; $t -lt 20; $t++) { $sa = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $saCond); if ($sa) { break }; Start-Sleep -Milliseconds 200 }
  if (-not $sa) { return 'no-saveas' }

  # File name combo (standard dialog AutomationId 1001): clear, type the full path.
  $fnCond = New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty, "1001")
  $fn = $sa.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $fnCond)
  if ($fn) { $fr = $fn.Current.BoundingRectangle; Click ([int]($fr.X+$fr.Width/2)) ([int]($fr.Y+$fr.Height/2)) }
  Start-Sleep -Milliseconds 250; Send '^a'; Start-Sleep -Milliseconds 80; Send '{DEL}'; Start-Sleep -Milliseconds 80
  Send-Literal $targetCsv; Start-Sleep -Milliseconds 300

  $svCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Save")
  $sv = $sa.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $svCond)
  if ($sv) { $svr = $sv.Current.BoundingRectangle; Click ([int]($svr.X+$svr.Width/2)) ([int]($svr.Y+$svr.Height/2)) } else { Send '{ENTER}' }
  Start-Sleep -Milliseconds 1000

  # "already exists - replace?" prompt: Yes if -Overwrite, else No + Cancel (skip).
  $confirm = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Confirm Save As")))
  if ($confirm) {
    if ($Overwrite) {
      $yes = $confirm.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Yes")))
      if ($yes) { $yr=$yes.Current.BoundingRectangle; Click ([int]($yr.X+$yr.Width/2)) ([int]($yr.Y+$yr.Height/2)) }; Start-Sleep -Milliseconds 800; return 'overwritten'
    } else {
      $no = $confirm.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "No")))
      if ($no) { $nr=$no.Current.BoundingRectangle; Click ([int]($nr.X+$nr.Width/2)) ([int]($nr.Y+$nr.Height/2)) }; Start-Sleep -Milliseconds 500
      $cancel = $sa.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Cancel")))
      if ($cancel) { $cr=$cancel.Current.BoundingRectangle; Click ([int]($cr.X+$cr.Width/2)) ([int]($cr.Y+$cr.Height/2)) }; Start-Sleep -Milliseconds 500
      return 'skipped-existing'
    }
  }
  Start-Sleep -Milliseconds 600
  if (Test-Path -LiteralPath $targetCsv) { return 'saved' }
  return 'save-failed'
}


# ===================== src\50-KeyOf.ps1 =====================
# KeyOf : normalize a customer name for matching (collapse spaces, upper-case)
# so the .xls name and the frequency-table name compare reliably.
function KeyOf([string]$s) { return ($s -replace '\s+',' ').Trim().ToUpperInvariant() }


# ===================== src\51-Load-Reference.ps1 =====================
# Load-Reference : read the reference customer list into two lookups:
#   ByKey   : KeyOf(RxOfficeName) -> @{ Name=<RxOffice name>; Freq; Tokens }  (exact + word-prefix)
#   ByTight : TightKey(x)         -> @{ Name=<RxOffice name>; Freq }          (punctuation/spacing-
#             insensitive; includes the RxOfficeName AND its ARName alias)
# A match gives us BOTH the exact RxOffice name to type in Accounting and the frequency.
#
# Accepts:
#   * RxOffice + ARName - comma CSV: Id,Code,RxOfficeName,StatementFrequency,ARName
#                         (ARName is the AR file/customer alias, e.g. "AbalosGuillermoOptical")
#   * RxOffice          - comma CSV: Id,Code,Name,StatementFrequency
#   * Contact export    - UTF-16, TAB-delimited: Name / Alias / StatementFrequency
#   * Simple            - comma CSV: Customer,Frequency
# Returns @{ ByKey=<hashtable>; ByTight=<hashtable>; Source=<path or '(none)'> }.
function Load-Reference([string]$path) {
  $byKey = @{}; $byTight = @{}
  $result = @{ ByKey=$byKey; ByTight=$byTight; Source='(none)' }
  if (-not $path -or -not (Test-Path -LiteralPath $path)) { return $result }

  # shared read (the file may be open in Excel)
  $fs = New-Object System.IO.FileStream($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
  $rd = New-Object System.IO.BinaryReader($fs); $bytes = $rd.ReadBytes([int]$fs.Length); $rd.Close(); $fs.Close()
  $enc = [System.Text.Encoding]::UTF8
  if     ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [System.Text.Encoding]::Unicode }
  elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $enc = [System.Text.Encoding]::BigEndianUnicode }
  $text = $enc.GetString($bytes)
  if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
  $lines = $text -split "\r?\n"
  if ($lines.Count -lt 2) { $result.Source = $path; return $result }

  # local helper: register one reference row
  $add = {
    param($name, $freq, $alias)
    if (-not $name -or -not $freq) { return }
    $entry = @{ Name=$name; Freq=$freq; Tokens=(Tokenize $name) }
    $k = KeyOf $name;   if ($k  -and -not $byKey.ContainsKey($k))   { $byKey[$k]  = $entry }
    $t = TightKey $name; if ($t -and -not $byTight.ContainsKey($t)) { $byTight[$t] = $entry }
    if ($alias) { $at = TightKey $alias; if ($at -and -not $byTight.ContainsKey($at)) { $byTight[$at] = $entry } }
  }

  if ($lines[0] -match 'StatementFrequency' -and $lines[0].Contains("`t")) {
    # ---- Contact export (tab-delimited) ----
    $hdr   = $lines[0].Split("`t") | ForEach-Object { $_.Trim().Trim('"') }
    $iName = [Array]::IndexOf($hdr,'Name'); $iAli = [Array]::IndexOf($hdr,'Alias'); $iFreq = [Array]::IndexOf($hdr,'StatementFrequency')
    for ($i=1; $i -lt $lines.Count; $i++) {
      if (-not $lines[$i].Trim()) { continue }
      $c = $lines[$i].Split("`t") | ForEach-Object { $_.Trim().Trim('"') }
      if ($iFreq -lt 0 -or $iFreq -ge $c.Count) { continue }
      $nm = if ($iName -ge 0 -and $iName -lt $c.Count) { $c[$iName] } else { '' }
      $al = if ($iAli  -ge 0 -and $iAli  -lt $c.Count) { $c[$iAli]  } else { '' }
      & $add $nm $c[$iFreq] $al
    }
  } else {
    # ---- comma CSV: RxOffice(+ARName) / RxOffice / simple ----
    $objs = ($lines | Where-Object { $_.Trim() }) | ConvertFrom-Csv
    if ($objs) {
      $cols = $objs[0].PSObject.Properties.Name
      $nameCol = if ($cols -contains 'RxOfficeName') { 'RxOfficeName' } elseif ($cols -contains 'Name') { 'Name' } elseif ($cols -contains 'Customer') { 'Customer' } else { $null }
      $freqCol = if ($cols -contains 'StatementFrequency') { 'StatementFrequency' } elseif ($cols -contains 'Frequency') { 'Frequency' } else { $null }
      $aliasCol = if ($cols -contains 'ARName') { 'ARName' } else { $null }
      if ($nameCol -and $freqCol) {
        foreach ($o in $objs) {
          $nm = ([string]$o.$nameCol).Trim()
          $f  = ([string]$o.$freqCol).Trim()
          $al = if ($aliasCol) { ([string]$o.$aliasCol).Trim() } else { '' }
          & $add $nm $f $al
        }
      }
    }
  }

  $result.Source = $path
  return $result
}


# ===================== src\52-Explain.ps1 =====================
# Explain : turn an internal status string into @(Result, HumanReason) for the
# result report (Result is one of SUCCESS / SKIPPED / FAILED).
function Explain($status, $freq) {
  switch ($status) {
    'saved'            { return @('SUCCESS', "Extracted ($freq)") }
    'overwritten'      { return @('SUCCESS', "Extracted & replaced ($freq)") }
    'skipped-existing' { return @('SKIPPED', 'CSV already in output folder') }
    'no-data'          { return @('FAILED',  'No statement produced - customer name did not auto-match, or no transactions in this date range') }
    'no-report'        { return @('FAILED',  'Report did not open after clicking OK') }
    'no-saveas'        { return @('FAILED',  'Export started but the Save dialog did not appear') }
    'save-failed'      { return @('FAILED',  'Save dialog completed but no CSV file was written') }
    'meta-error'       { return @('FAILED',  'Could not read customer/date from the .xls') }
    default            { if ($status -like 'error*') { return @('FAILED', "Automation error: $status") } else { return @('FAILED', $status) } }
  }
}


# ===================== src\53-Resolve-Reference.ps1 =====================
# Resolve-Reference : match an AR (.xls) customer name to a reference row and
# return the RxOffice NAME to type in Accounting plus the frequency.
#
# Matching, most-precise first:
#   1) exact         - normalized key equal to an RxOfficeName.
#   2) tight / ARName - punctuation/spacing/case-insensitive equal to an RxOfficeName
#                       OR its ARName alias. This is how "ABALOS GUILLERMO OPTICAL"
#                       matches ARName "AbalosGuillermoOptical" -> types "ABALOS GUILLERMO".
#   3) word-prefix    - one name's words are the leading run of the other's (>= 2 words),
#                       e.g. AR "FESAR ... GREENHILLS" <-> "FESAR ... GREENHILLS SHOPPING CENTER".
# Returns @{ Name=<RxOffice name or ''>; Freq; Via='exact'|'ARname'|"ref '<name>'"|'' }.
function Resolve-Reference([string]$name, $ref) {
  if (-not $ref) { return @{ Name=''; Freq=''; Via='' } }
  $byKey = $ref.ByKey; $byTight = $ref.ByTight
  if (-not $byKey -or $byKey.Count -eq 0) { return @{ Name=''; Freq=''; Via='' } }

  $k = KeyOf $name
  if ($byKey.ContainsKey($k)) { $e = $byKey[$k]; return @{ Name=$e.Name; Freq=$e.Freq; Via='exact' } }

  $tk = TightKey $name
  if ($byTight.ContainsKey($tk)) { $e = $byTight[$tk]; return @{ Name=$e.Name; Freq=$e.Freq; Via='ARname' } }

  $ar = Tokenize $name
  if ($ar.Count -eq 0) { return @{ Name=''; Freq=''; Via='' } }
  $best = $null; $bestScore = -1
  foreach ($e in $byKey.Values) {
    $t = $e.Tokens
    if (-not $t -or $t.Count -eq 0) { continue }
    $score = -1
    if     (Test-Prefix $t $ar) { if ($t.Count  -ge 2) { $score = ($t.Count  * 1000) - ($ar.Count - $t.Count) } }   # ref shorter
    elseif (Test-Prefix $ar $t) { if ($ar.Count -ge 2) { $score = ($ar.Count * 1000) - ($t.Count - $ar.Count) } }   # AR shorter
    if ($score -gt $bestScore) { $bestScore = $score; $best = $e }
  }
  if ($best) { return @{ Name=$best.Name; Freq=$best.Freq; Via=("ref '{0}'" -f $best.Name) } }
  return @{ Name=''; Freq=''; Via='' }
}


# ===================== src\60-Log-Load.ps1 =====================
# Load-Log : read the persistent run log (_SOA_Log.csv) into an ordered hashtable
# keyed by File name. Each value is the last-known row for that .xls. Empty if the
# log doesn't exist yet (first run). This is what makes the tool resumable: files
# already marked DONE here are skipped on the next run.
function Load-Log([string]$logCsv) {
  $log = [ordered]@{}
  if (-not (Test-Path -LiteralPath $logCsv)) { return $log }
  try {
    foreach ($row in (Import-Csv -LiteralPath $logCsv -Encoding UTF8)) {
      if ($row.File) { $log[$row.File] = $row }
    }
  } catch {}
  return $log
}


# ===================== src\61-Log-Save.ps1 =====================
# Save-Log : persist the run log as BOTH a CSV (source of truth for resuming) and
# a formatted Excel workbook (_SOA_Log.xlsx) for the user to read. Rows already
# come sorted by the caller. Columns:
#   Customer | File | DateRange | Frequency | Status | Reason | Csv | UpdatedOn
# Status is DONE / ERROR / SKIPPED. Errors are visible in the Status/Reason cols.
function Save-Log($rows, [string]$logCsv, [string]$logXlsx) {
  # CSV (UTF-8, so accented names survive) - used for resume next run.
  $rows | Select-Object Customer, RxOfficeName, File, DateRange, Frequency, Status, Reason, Csv, UpdatedOn |
          Export-Csv -NoTypeInformation -LiteralPath $logCsv -Encoding UTF8

  # Excel workbook for viewing, with a bold header and coloured Status cells.
  $excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
  try {
    $wb = $excel.Workbooks.Add(); $ws = $wb.Sheets.Item(1); $ws.Name = 'SOA Log'
    $headers = @('Customer','RxOfficeName','File','DateRange','Frequency','Status','Reason','Csv','UpdatedOn')
    for ($c=0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1,$c+1) = $headers[$c]; $ws.Cells.Item(1,$c+1).Font.Bold = $true }
    $r = 2
    foreach ($row in $rows) {
      $ws.Cells.Item($r,1) = [string]$row.Customer
      $ws.Cells.Item($r,2) = [string]$row.RxOfficeName
      $ws.Cells.Item($r,3) = [string]$row.File
      $ws.Cells.Item($r,4) = [string]$row.DateRange
      $ws.Cells.Item($r,5) = [string]$row.Frequency
      $ws.Cells.Item($r,6) = [string]$row.Status
      $ws.Cells.Item($r,7) = [string]$row.Reason
      $ws.Cells.Item($r,8) = [string]$row.Csv
      $ws.Cells.Item($r,9) = [string]$row.UpdatedOn
      switch ($row.Status) {
        'DONE'    { $ws.Cells.Item($r,6).Interior.Color = 0x90EE90 }  # light green
        'ERROR'   { $ws.Cells.Item($r,6).Interior.Color = 0x9090FF }  # light red (BGR)
        'SKIPPED' { $ws.Cells.Item($r,6).Interior.Color = 0xE0E0E0 }  # grey
      }
      $r++
    }
    $ws.Columns.Item("A:I").AutoFit() | Out-Null
    try { $ws.Application.ActiveWindow.SplitRow = 1; $ws.Application.ActiveWindow.FreezePanes = $true } catch {}
    if (Test-Path -LiteralPath $logXlsx) { Remove-Item -LiteralPath $logXlsx -Force -ErrorAction SilentlyContinue }
    $wb.SaveAs($logXlsx, 51)   # 51 = .xlsx
    $wb.Close($true) | Out-Null
  } finally { $excel.Quit()|Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null }
}


# ===================== src\99-Main.ps1 =====================
# =====================  MAIN PROGRAM  =====================
# Runs after every function above is defined.
# Test hooks (optional): SOA_FOLDER, SOA_MAX, SOA_REF (reference CSV), SOA_OUT.

# ---- 1) Choose the AR folder (the .xls files) ----
if ($env:SOA_FOLDER) { $SourceFolder = $env:SOA_FOLDER } else { $SourceFolder = Pick-Folder }
$MaxFiles = if ($env:SOA_MAX) { [int]$env:SOA_MAX } else { 0 }
if (-not $SourceFolder) { Write-Host "No folder selected. Exiting."; return }
if (-not (Test-Path -LiteralPath $SourceFolder)) { Write-Host "Folder not found: $SourceFolder"; return }
$SourceFolder = (Resolve-Path -LiteralPath $SourceFolder).Path

# ---- 2) Choose which Accounting window to drive, then confirm it's logged in ----
$acct = Choose-Acct
if (-not $acct) {
  if (@(Get-AcctCandidates).Count -eq 0) {
    Show-Message "Accounting is not open.`n`nPlease open Accounting.exe and log in (user hfabros), then run this again."
  }
  Write-Host "No Accounting window selected. Exiting." -ForegroundColor Red; return
}
Write-Host ("Accounting        : PID {0}  {1}" -f $acct.Pid, $acct.Title)

# Elevation/UAC mismatch: can't automate an elevated app from a normal one.
if (Test-ElevationMismatch) {
  Write-Host "`nELEVATION MISMATCH:`n$ElevationHelp" -ForegroundColor Red
  Show-Message $ElevationHelp
  return
}

if (-not (Assert-AcctLoggedIn)) { Write-Host "Not logged in to Accounting. Exiting." -ForegroundColor Red; return }

# ---- 3) Output folder (separate from the AR data). Override: SOA_OUT ----
if ($env:SOA_OUT) {
  $OutputFolder = $env:SOA_OUT
} else {
  $parent = Split-Path -Parent $SourceFolder
  $leaf   = Split-Path -Leaf   $SourceFolder
  $OutputFolder = Join-Path $parent ($leaf + ' - SOA CSV')
}
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
New-Item -ItemType Directory -Force -Path $ShotDir      | Out-Null

# ---- 4) Reference CSV (RxOffice/Contact export): gives the customer NAME to type
#         in Accounting AND the frequency, matched from each .xls customer name. ----
if     ($env:SOA_REF)     { $refPath = $env:SOA_REF }
elseif ($env:SOA_FREQCSV) { $refPath = $env:SOA_FREQCSV }
else   { $refPath = Browse-File "Select the reference CSV (RxOffice export: Name + StatementFrequency)" }
$ref = Load-Reference $refPath

# ---- 5) Load the persistent run log (for resume) ----
$logCsv  = Join-Path $OutputFolder '_SOA_Log.csv'
$logXlsx = Join-Path $OutputFolder '_SOA_Log.xlsx'
$log     = Load-Log $logCsv
$alreadyDone = @($log.Values | Where-Object { $_.Status -eq 'DONE' }).Count

Write-Host "Input (AR data)  : $SourceFolder"
Write-Host "Output (CSV)     : $OutputFolder"
Write-Host "Reference        : $($ref.Source)  ($($ref.ByKey.Count) names)"
Write-Host "Log (resume from): $logCsv  ($alreadyDone already DONE)"
Write-Host "Fallback order   : $($FrequencyOrder -join ' -> ')`n"

# ---- 6) Process each .xls (skipping ones already extracted) ----
$xlsFiles = Get-ChildItem -LiteralPath $SourceFolder -Filter *.xls -File | Sort-Object Name
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
$did = 0; $processed = 0
foreach ($xls in $xlsFiles) {
  $base = [System.IO.Path]::GetFileNameWithoutExtension($xls.Name)
  if ($base.ToLower().EndsWith('.csv')) { $base = $base.Substring(0, $base.Length-4) }
  $csvPath = Join-Path $OutputFolder ($base + '.csv')

  $prior = $log[$xls.Name]
  $isDone = ($prior -and $prior.Status -eq 'DONE') -or (Test-Path -LiteralPath $csvPath)
  if ($isDone -and -not $Overwrite) {
    Write-Host ("SKIP (already extracted): {0}" -f $xls.Name) -ForegroundColor DarkGray
    if (-not $prior) {
      $log[$xls.Name] = [pscustomobject]@{ Customer=$base; RxOfficeName=''; File=$xls.Name; DateRange=''; Frequency=''
        Status='DONE'; Reason='CSV already present in output folder'; Csv=(Split-Path $csvPath -Leaf); UpdatedOn=$stamp }
    }
    continue
  }

  try { $meta = Read-XlsMeta $xls.FullName }
  catch {
    Write-Host ("META FAIL {0}: {1}" -f $xls.Name,$_.Exception.Message) -ForegroundColor Red
    $log[$xls.Name] = [pscustomobject]@{ Customer=''; RxOfficeName=''; File=$xls.Name; DateRange=''; Frequency=''
      Status='ERROR'; Reason=(Explain 'meta-error' '')[1]; Csv=''; UpdatedOn=$stamp }
    continue
  }

  $range = "{0}..{1}" -f $meta.From, $meta.To

  # Match the .xls (AR) customer to the reference: get the RxOffice NAME to type
  # in Accounting and the frequency. Fall back to the raw .xls name if unmatched.
  $r = Resolve-Reference $meta.Customer $ref
  $known   = $r.Freq
  # Type the matched RxOffice name (exact / normalized / word-prefix). Matching is
  # precise (>=2-word prefix), so this is the canonical name Accounting expects
  # (e.g. AR "ABALOS GUILLERMO OPTICAL" -> types "ABALOS GUILLERMO"). Only when there
  # is NO reference match do we fall back to typing the raw .xls name.
  $custName = if ($r.Name) { $r.Name } else { $meta.Customer }
  if ($known) { $tryOrder = @($known) + ($FrequencyOrder | Where-Object { $_ -ne $known }) }
  else        { $tryOrder = $FrequencyOrder }

  $nameTag = if ($r.Via -eq 'exact' -or $r.Via -eq 'normalized') { " -> ref '$custName'" } elseif ($r.Name) { " -> $($r.Via)" } else { " (no ref match; raw name)" }
  $freqTag = if ($known) { "freq=$known" } else { "freq=unknown -> guess" }
  Write-Host ("PROCESS: {0}  [{1}]{2}  {3}  {4}" -f $xls.Name,$meta.Customer,$nameTag,$range,$freqTag) -ForegroundColor Cyan

  $status='no-data'; $usedFreq=''
  foreach ($freq in $tryOrder) {
    try {
      Reset-State
      Open-StatementDialog | Out-Null
      Set-Customer $custName          # <-- type the RxOffice name from the reference
      if ($freq -ne 'Monthly') { Set-Frequency $freq }
      Set-DateField 'From' $meta.From; Set-DateField 'To' $meta.To
      Shot ("{0}_{1}_form" -f $base,$freq); Click-OK; Start-Sleep -Seconds 5
      if (-not (Get-ReportWindow)) { Write-Host ("   {0}: no report -> next" -f $freq) -ForegroundColor Yellow; Close-Dialog; continue }
      Shot ("{0}_{1}_report" -f $base,$freq)
      $status = Export-Csv-FromReport $csvPath; $usedFreq = $freq; Close-Report
      Write-Host ("   {0}: {1}" -f $freq,$status) -ForegroundColor Green
      if ($status -in @('saved','overwritten')) { break }
    } catch {
      Write-Host ("   {0}: ERROR {1}" -f $freq,$_.Exception.Message) -ForegroundColor Red
      Close-Dialog; Close-Report; $status = "error: $($_.Exception.Message)"
    }
  }

  $ex = Explain $status $usedFreq          # -> @(SUCCESS|FAILED|SKIPPED, reason)
  $logStatus = switch ($ex[0]) { 'SUCCESS' {'DONE'} 'SKIPPED' {'SKIPPED'} default {'ERROR'} }
  $reason = if ($logStatus -eq 'DONE' -and $r.Name) { "$($ex[1]); typed RxOffice name '$custName' (match: $($r.Via))" }
            else { $ex[1] }
  $log[$xls.Name] = [pscustomobject]@{ Customer=$meta.Customer; RxOfficeName=$custName; File=$xls.Name; DateRange=$range; Frequency=$usedFreq
    Status=$logStatus; Reason=$reason; Csv=$(if ($logStatus -eq 'DONE') { Split-Path $csvPath -Leaf } else { '' }); UpdatedOn=$stamp }

  $did++; $processed++
  # Save the log after every file so a crash/close still leaves an up-to-date, resumable log.
  Save-Log ($log.Values) $logCsv $logXlsx
  if ($MaxFiles -gt 0 -and $processed -ge $MaxFiles) { Write-Host "`n(Test cap reached: $MaxFiles files)"; break }
}

# ---- 7) Final save + summary ----
Save-Log ($log.Values) $logCsv $logXlsx
$done  = @($log.Values | Where-Object { $_.Status -eq 'DONE'  }).Count
$err   = @($log.Values | Where-Object { $_.Status -eq 'ERROR' }).Count
$skip  = @($log.Values | Where-Object { $_.Status -eq 'SKIPPED' }).Count

Write-Host "`n================ SUMMARY ================"
Write-Host ("Processed this run : {0}" -f $did)
Write-Host ("DONE (total)       : {0}" -f $done)
Write-Host ("ERROR (need retry) : {0}" -f $err)
Write-Host ("SKIPPED            : {0}" -f $skip)
Write-Host "`nCSV output folder  : $OutputFolder"
Write-Host "Log (Excel)        : $logXlsx"
Write-Host "Log (CSV)          : $logCsv"
Write-Host "Step screenshots   : $ShotDir"
if ($err -gt 0) { Write-Host "`nRe-run to retry the ERROR rows; DONE rows are skipped automatically." -ForegroundColor Yellow }



