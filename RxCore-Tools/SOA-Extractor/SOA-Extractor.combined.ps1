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
# Get-Acct : return the running Accounting.exe process, or throw if not running.
function Get-Acct {
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

function Get-AcctWindowTitles {
  $procs = Get-Process Accounting -ErrorAction SilentlyContinue
  if (-not $procs) { return @() }
  $titles = New-Object System.Collections.Generic.List[string]
  $desktop = $AE::RootElement
  foreach ($proc in $procs) {
    try {
      $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, [int]$proc.Id)
      $wins = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
      foreach ($w in $wins) { $titles.Add([string]$w.Current.Name) }
    } catch {}
    if ($proc.MainWindowTitle) { $titles.Add([string]$proc.MainWindowTitle) }
  }
  return $titles
}

function Assert-AcctLoggedIn {
  $procs = Get-Process Accounting -ErrorAction SilentlyContinue
  if (-not $procs) {
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


# ===================== src\30-Get-Dialog.ps1 =====================
# Get-Dialog : return the "Statement options" window element, or $null if closed.
function Get-Dialog {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Statement options")
  return $AE::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
}


# ===================== src\31-Open-StatementDialog.ps1 =====================
# Open-StatementDialog : go to the Accounting Print tab and double-click the
# Statement tile's ICON (~25px above its label - the label itself isn't clickable),
# then wait for the "Statement options" dialog. Returns the dialog element.
function Open-StatementDialog {
  $p = Focus-Acct; $root = $AE::FromHandle($p.MainWindowHandle)
  $apCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Accounting Print")
  foreach ($ap in $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $apCond)) {
    if ($ap.Current.ClassName -eq 'TextBlock' -and $ap.Current.BoundingRectangle.X -lt 200) {
      $b = $ap.Current.BoundingRectangle; Click ([int]($b.X+$b.Width/2)) ([int]($b.Y+$b.Height/2)); Start-Sleep -Milliseconds 600; break
    }
  }
  $parCond = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, "PrintAccountingReports")
  $par = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $parCond)
  if (-not $par) { throw "Accounting Print pane not found." }
  $stCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Statement")
  $st = $par.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $stCond)
  if (-not $st) { throw "'Statement' tile not found." }
  $r = $st.Current.BoundingRectangle; $cx = [int]($r.X+$r.Width/2); $iy = [int]($r.Y-25)
  for ($i=0; $i -lt 4; $i++) {
    Click $cx $iy -Double
    for ($t=0; $t -lt 15; $t++) { $d = Get-Dialog; if ($d) { return $d }; Start-Sleep -Milliseconds 200 }
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


# ===================== src\51-Load-FreqTable.ps1 =====================
# Load-FreqTable : read a reference CSV into a hashtable  KeyOf(name) -> frequency.
# Supports two formats automatically:
#   1) Contact.exe "Export" file  - UTF-16, TAB-delimited, many columns incl.
#      "Name", "Alias" and "StatementFrequency". Both Name and Alias are mapped
#      so an .xls customer name matches whichever Contact used.
#   2) Simple 2-column CSV         - "Customer","Frequency" (UTF-8), e.g. an older
#      CustomerFrequency export.
# Returns @{ Map = <hashtable>; Source = <path or '(none)'> }.
function Load-FreqTable([string]$path) {
  $map = @{}
  if (-not $path -or -not (Test-Path -LiteralPath $path)) { return @{ Map=$map; Source='(none)' } }

  # Detect encoding from the byte-order mark.
  $bytes = [System.IO.File]::ReadAllBytes($path)
  $enc = [System.Text.Encoding]::UTF8
  if     ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [System.Text.Encoding]::Unicode }           # UTF-16 LE
  elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $enc = [System.Text.Encoding]::BigEndianUnicode }   # UTF-16 BE
  $text  = $enc.GetString($bytes)
  $lines = $text -split "\r?\n"
  if ($lines.Count -lt 2) { return @{ Map=$map; Source=$path } }

  $header = $lines[0]
  if ($header -match 'StatementFrequency' -and $header.Contains("`t")) {
    # ---- Contact export (tab-delimited) ----
    $hdr   = $header.Split("`t") | ForEach-Object { $_.Trim().Trim('"') }
    $iName = [Array]::IndexOf($hdr, 'Name')
    $iAli  = [Array]::IndexOf($hdr, 'Alias')
    $iFreq = [Array]::IndexOf($hdr, 'StatementFrequency')
    for ($i = 1; $i -lt $lines.Count; $i++) {
      if (-not $lines[$i].Trim()) { continue }
      $c = $lines[$i].Split("`t") | ForEach-Object { $_.Trim().Trim('"') }
      if ($iFreq -lt 0 -or $iFreq -ge $c.Count) { continue }
      $f = $c[$iFreq]; if (-not $f) { continue }
      if ($iName -ge 0 -and $iName -lt $c.Count) { $k = KeyOf $c[$iName]; if ($k -and -not $map.ContainsKey($k)) { $map[$k] = $f } }
      if ($iAli  -ge 0 -and $iAli  -lt $c.Count -and $c[$iAli]) { $k = KeyOf $c[$iAli]; if ($k -and -not $map.ContainsKey($k)) { $map[$k] = $f } }
    }
  } else {
    # ---- Simple "Customer,Frequency" CSV ----
    foreach ($row in (Import-Csv -LiteralPath $path -Encoding UTF8)) {
      $k = KeyOf ([string]$row.Customer); $f = ([string]$row.Frequency).Trim()
      if ($k -and $f -and -not $map.ContainsKey($k)) { $map[$k] = $f }
    }
  }
  return @{ Map=$map; Source=$path }
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


# ===================== src\53-Resolve-Frequency.ps1 =====================
# Resolve-Frequency : find a customer's frequency in the reference map.
# Tries an exact (normalized) match first; if none, tries progressively shorter
# forms of the name and takes the first table entry that CONTAINS that form.
# Returns @{ Freq = <frequency or ''>; Via = 'exact' | <term used> | '' }.
function Resolve-Frequency([string]$name, [hashtable]$map) {
  if (-not $map -or $map.Count -eq 0) { return @{ Freq=''; Via='' } }
  $k = KeyOf $name
  if ($map.ContainsKey($k)) { return @{ Freq=$map[$k]; Via='exact' } }
  foreach ($term in (Get-Candidates $name)) {
    $tk = KeyOf $term
    if (-not $tk) { continue }
    $hit = $null
    foreach ($key in $map.Keys) { if ($key.Contains($tk)) { $hit = $key; break } }
    if ($hit) { return @{ Freq=$map[$hit]; Via=$term } }
  }
  return @{ Freq=''; Via='' }
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
  $rows | Select-Object Customer, File, DateRange, Frequency, Status, Reason, Csv, UpdatedOn |
          Export-Csv -NoTypeInformation -LiteralPath $logCsv -Encoding UTF8

  # Excel workbook for viewing, with a bold header and coloured Status cells.
  $excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
  try {
    $wb = $excel.Workbooks.Add(); $ws = $wb.Sheets.Item(1); $ws.Name = 'SOA Log'
    $headers = @('Customer','File','DateRange','Frequency','Status','Reason','Csv','UpdatedOn')
    for ($c=0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1,$c+1) = $headers[$c]; $ws.Cells.Item(1,$c+1).Font.Bold = $true }
    $r = 2
    foreach ($row in $rows) {
      $ws.Cells.Item($r,1) = [string]$row.Customer
      $ws.Cells.Item($r,2) = [string]$row.File
      $ws.Cells.Item($r,3) = [string]$row.DateRange
      $ws.Cells.Item($r,4) = [string]$row.Frequency
      $ws.Cells.Item($r,5) = [string]$row.Status
      $ws.Cells.Item($r,6) = [string]$row.Reason
      $ws.Cells.Item($r,7) = [string]$row.Csv
      $ws.Cells.Item($r,8) = [string]$row.UpdatedOn
      switch ($row.Status) {
        'DONE'    { $ws.Cells.Item($r,5).Interior.Color = 0x90EE90 }  # light green
        'ERROR'   { $ws.Cells.Item($r,5).Interior.Color = 0x9090FF }  # light red (BGR)
        'SKIPPED' { $ws.Cells.Item($r,5).Interior.Color = 0xE0E0E0 }  # grey
      }
      $r++
    }
    $ws.Columns.Item("A:H").AutoFit() | Out-Null
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

# ---- 2) Must be logged in to Accounting ----
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

# ---- 4) Reference CSV for frequencies (the file exported from Contact) ----
if     ($env:SOA_REF)     { $refPath = $env:SOA_REF }
elseif ($env:SOA_FREQCSV) { $refPath = $env:SOA_FREQCSV }
else   { $refPath = Browse-File "Select the Customer CSV exported from Contact (frequency reference)" }
$ref     = Load-FreqTable $refPath
$FreqMap = $ref.Map

# ---- 5) Load the persistent run log (for resume) ----
$logCsv  = Join-Path $OutputFolder '_SOA_Log.csv'
$logXlsx = Join-Path $OutputFolder '_SOA_Log.xlsx'
$log     = Load-Log $logCsv
$alreadyDone = @($log.Values | Where-Object { $_.Status -eq 'DONE' }).Count

Write-Host "Input (AR data)  : $SourceFolder"
Write-Host "Output (CSV)     : $OutputFolder"
Write-Host "Frequency ref    : $($ref.Source)  ($($FreqMap.Count) names)"
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
      $log[$xls.Name] = [pscustomobject]@{ Customer=$base; File=$xls.Name; DateRange=''; Frequency=''
        Status='DONE'; Reason='CSV already present in output folder'; Csv=(Split-Path $csvPath -Leaf); UpdatedOn=$stamp }
    }
    continue
  }

  try { $meta = Read-XlsMeta $xls.FullName }
  catch {
    Write-Host ("META FAIL {0}: {1}" -f $xls.Name,$_.Exception.Message) -ForegroundColor Red
    $log[$xls.Name] = [pscustomobject]@{ Customer=''; File=$xls.Name; DateRange=''; Frequency=''
      Status='ERROR'; Reason=(Explain 'meta-error' '')[1]; Csv=''; UpdatedOn=$stamp }
    continue
  }

  $range = "{0}..{1}" -f $meta.From, $meta.To
  $r = Resolve-Frequency $meta.Customer $FreqMap
  $known = $r.Freq
  if ($known) { $tryOrder = @($known) + ($FrequencyOrder | Where-Object { $_ -ne $known }) }
  else        { $tryOrder = $FrequencyOrder }

  $freqTag = if ($known) { "freq=$known ($($r.Via))" } else { "freq=unknown -> will guess" }
  Write-Host ("PROCESS: {0}  [{1}]  {2}  {3}" -f $xls.Name,$meta.Customer,$range,$freqTag) -ForegroundColor Cyan

  $status='no-data'; $usedFreq=''
  foreach ($freq in $tryOrder) {
    try {
      Reset-State
      Open-StatementDialog | Out-Null
      Set-Customer $meta.Customer
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
  $reason = if ($known -and $logStatus -eq 'DONE') { "$($ex[1]); frequency from reference ($($r.Via))" } else { $ex[1] }
  $log[$xls.Name] = [pscustomobject]@{ Customer=$meta.Customer; File=$xls.Name; DateRange=$range; Frequency=$usedFreq
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



