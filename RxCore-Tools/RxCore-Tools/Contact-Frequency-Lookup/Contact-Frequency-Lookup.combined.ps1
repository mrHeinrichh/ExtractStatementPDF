# ===================== src\00-Header.ps1 =====================
<#
  Contact Frequency Lookup  -  assembled source
  =============================================
  One function per file under this src\ folder. build.ps1 concatenates every
  src\*.ps1 (in filename order) into one script and compiles it to
  Contact-Frequency-Lookup.exe with ps2exe.

  Load order (by numeric filename prefix):
     00 header + assemblies (this file)   01 settings (grid offsets)
     02 Win32 interop                     10-37 one function per file
     99 main program body (runs last)

  What it does: reads the customer name from every .xls in a chosen folder,
  drives Contact.exe to look each one up (Name column filter -> select row ->
  Account tab -> read Frequency), and writes a Name+Frequency table to Excel in
  a SEPARATE "<folder> - Frequency" output folder. If an exact name finds nothing
  it shortens the search until a match appears (closest match wins).

  Requires: Contact.exe already OPEN and LOGGED IN (user hfabros); Excel installed.
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes


# ===================== src\01-Settings.ps1 =====================
# ---- Grid geometry: offsets from the MAXIMIZED window's top-left corner -----
# The results grid isn't automatable, so these two spots are clicked by offset.
# (Window maximizes to -8,-8; screen(830,792) => +838,+800 ; screen(400,820) => +408,+828.)
$NAME_FILTER_DX = 838   # the "Name" column filter box (under the Name header)
$NAME_FILTER_DY = 800
$ROW1_DX        = 408   # the first data row in the results grid
$ROW1_DY        = 828


# ===================== src\02-Interop.ps1 =====================
# ---- Win32 interop + UI Automation shortcuts ------------------------------
$sig = @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h,int n);
[DllImport("user32.dll")] public static extern bool SetWindowPos(System.IntPtr h,System.IntPtr ins,int x,int y,int cx,int cy,uint f);
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint f,uint x,uint y,uint d,int e);
[DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr h,out RECT r);
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)] public struct RECT{public int Left,Top,Right,Bottom;}
'@
if (-not ([System.Management.Automation.PSTypeName]'CF.Win').Type) {
  Add-Type -MemberDefinition $sig -Name Win -Namespace CF -PassThru | Out-Null
}
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]::Descendants


# ===================== src\10-Pick-Folder.ps1 =====================
# Pick-Folder : paste-a-path dialog (with Browse...). Returns the path or $null.
function Pick-Folder {
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Contact Frequency Lookup"; $form.Size = New-Object System.Drawing.Size(620,190)
  $form.StartPosition = "CenterScreen"; $form.FormBorderStyle = "FixedDialog"
  $form.MaximizeBox=$false; $form.MinimizeBox=$false; $form.TopMost=$true
  $l = New-Object System.Windows.Forms.Label; $l.Text="Paste the folder path that contains the .xls files:"; $l.Location='15,15'; $l.AutoSize=$true; $form.Controls.Add($l)
  $box = New-Object System.Windows.Forms.TextBox; $box.Location='15,45'; $box.Size='490,25'; $box.Anchor='Top,Left,Right'
  $box.Add_KeyDown({ if ($_.Control -and $_.KeyCode -eq 'A') { $box.SelectAll(); $_.SuppressKeyPress=$true } }); $form.Controls.Add($box)
  $br = New-Object System.Windows.Forms.Button; $br.Text="Browse..."; $br.Location='510,44'; $br.Size='75,25'; $br.Anchor='Top,Right'
  $br.Add_Click({ $fb=New-Object System.Windows.Forms.FolderBrowserDialog; if($fb.ShowDialog() -eq 'OK'){ $box.Text=$fb.SelectedPath } }); $form.Controls.Add($br)
  $ok = New-Object System.Windows.Forms.Button; $ok.Text="OK"; $ok.Location='410,100'; $ok.Size='80,30'; $ok.DialogResult='OK'; $ok.Anchor='Bottom,Right'; $form.Controls.Add($ok); $form.AcceptButton=$ok
  $cx = New-Object System.Windows.Forms.Button; $cx.Text="Cancel"; $cx.Location='500,100'; $cx.Size='80,30'; $cx.DialogResult='Cancel'; $cx.Anchor='Bottom,Right'; $form.Controls.Add($cx); $form.CancelButton=$cx
  $form.Add_Shown({ $form.Activate(); $box.Focus() })
  if ($form.ShowDialog() -ne 'OK') { return $null }
  return $box.Text.Trim().Trim('"')
}


# ===================== src\11-Get-Contact.ps1 =====================
# Get-Contact : return the running Contact.exe process, or throw if not running.
function Get-Contact {
  $p = Get-Process Contact -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $p) { throw "Contact.exe is not running. Open it and log in first." }
  return $p
}


# ===================== src\12-Normalize-Window.ps1 =====================
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


# ===================== src\13-WinRect.ps1 =====================
# WinRect : return the Contact window rectangle (Left/Top/Right/Bottom), used as
# the anchor for the grid click offsets.
function WinRect { $p=Get-Contact; $r=New-Object CF.Win+RECT; [CF.Win]::GetWindowRect($p.MainWindowHandle,[ref]$r) | Out-Null; return $r }


# ===================== src\14-Click.ps1 =====================
# Click : move the mouse to a screen coordinate and left-click.
function Click([int]$x,[int]$y){ [CF.Win]::SetCursorPos($x,$y)|Out-Null; Start-Sleep -Milliseconds 120; [CF.Win]::mouse_event(0x02,0,0,0,0); [CF.Win]::mouse_event(0x04,0,0,0,0) }


# ===================== src\15-Send.ps1 =====================
# Send : send raw keystrokes to the focused control (SendKeys syntax).
function Send([string]$k){ [System.Windows.Forms.SendKeys]::SendWait($k) }


# ===================== src\16-Paste.ps1 =====================
# Paste : put text on the clipboard and Ctrl+V it. Used for the search term so
# the whole string (incl. accented names like AVENDAÃ‘O) lands reliably - the
# search-as-you-type box drops characters when typed key by key.
function Paste([string]$t){ [System.Windows.Forms.Clipboard]::SetText($t); Start-Sleep -Milliseconds 150; Send '^v' }


# ===================== src\20-Read-XlsCustomer.ps1 =====================
# Read-XlsCustomer : open an .xls read-only (Excel COM) and return the customer
# name from cell E9.
function Read-XlsCustomer([string]$path) {
  $x = New-Object -ComObject Excel.Application; $x.Visible=$false; $x.DisplayAlerts=$false
  try {
    $wb=$x.Workbooks.Open($path,[Type]::Missing,$true); $ws=$wb.Sheets.Item(1)
    $name = ([string]$ws.Cells.Item(9,5).Text).Trim()
    $wb.Close($false) | Out-Null; return $name
  } finally { $x.Quit()|Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)|Out-Null }
}


# ===================== src\30-Get-MatchCount.ps1 =====================
# Get-MatchCount : read the "selected / TOTAL" indicator at the top-right and
# return TOTAL (the number just right of the "/"). Tells us how many grid rows
# the current filter matched (0 = not found, 1 = exact, >1 = ambiguous). -1 on error.
function Get-MatchCount {
  $root = $AE::FromHandle((Get-Contact).MainWindowHandle)
  $txtCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty,[System.Windows.Automation.ControlType]::Text)
  $txts = $root.FindAll($TS,$txtCond)
  $slash=$null
  foreach($t in $txts){ $b=$t.Current.BoundingRectangle; if($t.Current.Name -eq '/' -and $b.Y -lt 90){ $slash=$t; break } }
  if (-not $slash) { return -1 }
  $sy=$slash.Current.BoundingRectangle.Y; $sx=$slash.Current.BoundingRectangle.X
  $best=$null;$bd=99999
  foreach($t in $txts){ $b=$t.Current.BoundingRectangle; if([math]::Abs($b.Y-$sy)-le 4 -and $b.X -gt $sx){ $dx=$b.X-$sx; if($dx-lt $bd){$bd=$dx;$best=$t} } }
  if(-not $best){ return -1 }
  $n=0; if([int]::TryParse(($best.Current.Name -replace '[^\d]',''),[ref]$n)){ return $n }
  return -1
}


# ===================== src\31-Click-AccountTab.ps1 =====================
# Click-AccountTab : the Frequency field lives on the Account tab. Click the
# LEFTMOST "Account" tab (the one just right of GENERIC), then press ESC to
# dismiss any dropdown that a stray click may have opened (does not change tabs).
function Click-AccountTab {
  $root = $AE::FromHandle((Get-Contact).MainWindowHandle)
  $accs = $root.FindAll($TS,(New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty,"Account")))
  $best = $null
  foreach ($a in $accs) { if ($a.Current.ClassName -eq 'TextBlock') { if (-not $best -or $a.Current.BoundingRectangle.X -lt $best.Current.BoundingRectangle.X) { $best = $a } } }
  if ($best) {
    $b = $best.Current.BoundingRectangle
    Click ([int]($b.X + $b.Width/2)) ([int]($b.Y + $b.Height/2))
    Start-Sleep -Milliseconds 500
    Send '{ESC}'
    Start-Sleep -Milliseconds 200
  }
}


# ===================== src\32-Read-Frequency.ps1 =====================
# Read-Frequency : read the Invoice Frequency value from the Account tab. Finds
# the "Frequency" label, then the editable combo box immediately to its right and
# returns its value via the UI Automation ValuePattern. '' if not found.
function Read-Frequency {
  $root = $AE::FromHandle((Get-Contact).MainWindowHandle)
  $lbl = $root.FindFirst($TS,(New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty,"Frequency")))
  if(-not $lbl){ return '' }
  $ly=$lbl.Current.BoundingRectangle.Y; $lx=$lbl.Current.BoundingRectangle.X
  $edits=$root.FindAll($TS,(New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty,"PART_EditableTextBox")))
  $best=$null;$bd=99999
  foreach($e in $edits){ $b=$e.Current.BoundingRectangle; if([math]::Abs($b.Y-$ly)-le 6 -and $b.X -gt $lx){ $dx=$b.X-$lx; if($dx-lt $bd){$bd=$dx;$best=$e} } }
  if(-not $best){ return '' }
  try { return $best.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value } catch { return '' }
}


# ===================== src\33-Clear-NameFilter.ps1 =====================
# Clear-NameFilter : empty the Name column filter box and re-apply, so the grid
# resets to the full list before the next search.
function Clear-NameFilter {
  $r = WinRect
  Click ($r.Left + $NAME_FILTER_DX) ($r.Top + $NAME_FILTER_DY)
  Start-Sleep -Milliseconds 150; Send '^a'; Send '{DEL}'; Start-Sleep -Milliseconds 150; Send '{ENTER}'; Start-Sleep -Milliseconds 400
}


# ===================== src\34-Clear-FullTextSearch.ps1 =====================
# Clear-FullTextSearch : empty the bottom "Full Text Search" box. A leftover value
# there AND-combines with the Name column filter and hides real matches, so we
# clear it once at startup.
function Clear-FullTextSearch {
  $root = $AE::FromHandle((Get-Contact).MainWindowHandle)
  $sb = $root.FindFirst($TS,(New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty,"PART_SearchAsYouTypeTextBox")))
  if ($sb) {
    $b = $sb.Current.BoundingRectangle
    Click ([int]($b.X + $b.Width/2)) ([int]($b.Y + $b.Height/2))
    Start-Sleep -Milliseconds 150; Send '^a'; Send '{DEL}'; Start-Sleep -Milliseconds 150; Send '{ENTER}'; Start-Sleep -Milliseconds 600
  }
}


# ===================== src\35-Search-NameFilter.ps1 =====================
# Search-NameFilter : type a term into the Name column filter (paste + Enter) and
# return the resulting match count.
function Search-NameFilter([string]$term) {
  $r = WinRect
  Click ($r.Left + $NAME_FILTER_DX) ($r.Top + $NAME_FILTER_DY)
  Start-Sleep -Milliseconds 200; Send '^a'; Send '{DEL}'; Start-Sleep -Milliseconds 150
  Paste $term; Start-Sleep -Milliseconds 400; Send '{ENTER}'; Start-Sleep -Milliseconds 1400
  return (Get-MatchCount)
}


# ===================== src\36-Get-Candidates.ps1 =====================
# Get-Candidates : build progressively shorter search terms for fuzzy matching -
# the full name first, then drop trailing words, then trim characters off the
# first word. Left-anchored so the closest customer surfaces when the exact name
# isn't found (e.g. "Esca Empire Corp. -Baclaran" -> ... -> "Esca").
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


# ===================== src\37-Lookup-Frequency.ps1 =====================
# Lookup-Frequency : find a customer's frequency. Try each candidate term (see
# Get-Candidates) until one matches, select row 1, open the Account tab, and read
# the Frequency. Returns @{ Freq; Term (the term that matched); Matches (count) }.
function Lookup-Frequency([string]$name) {
  $r = WinRect
  $matchedTerm = ''; $count = 0
  foreach ($term in (Get-Candidates $name)) {
    $count = Search-NameFilter $term
    if ($count -ge 1) { $matchedTerm = $term; break }
  }
  if ($count -le 0) { return @{ Freq=''; Term=''; Matches=0 } }
  Click ($r.Left + $ROW1_DX) ($r.Top + $ROW1_DY)   # select row 1 -> form loads it
  Start-Sleep -Milliseconds 1200
  Click-AccountTab                                  # ensure Account tab is showing
  $freq = Read-Frequency
  return @{ Freq=$freq; Term=$matchedTerm; Matches=$count }
}


# ===================== src\99-Main.ps1 =====================
# =====================  MAIN PROGRAM  =====================
# Runs after every function above is defined.
# Test hook: set CF_FOLDER (+ optional CF_MAX) to skip the picker and cap file count.
if ($env:CF_FOLDER) { $SourceFolder = $env:CF_FOLDER } else { $SourceFolder = Pick-Folder }
$MaxFiles = if ($env:CF_MAX) { [int]$env:CF_MAX } else { 0 }
if (-not $SourceFolder) { Write-Host "No folder selected."; return }
if (-not (Test-Path -LiteralPath $SourceFolder)) { Write-Host "Folder not found: $SourceFolder"; return }
$SourceFolder = (Resolve-Path -LiteralPath $SourceFolder).Path

# Output goes to a SEPARATE sibling folder (AR data stays untouched). Override: CF_OUT.
$parent = Split-Path -Parent $SourceFolder; $leaf = Split-Path -Leaf $SourceFolder
$OutputFolder = if ($env:CF_OUT) { $env:CF_OUT } else { Join-Path $parent ($leaf + ' - Frequency') }
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null

try { Get-Contact | Out-Null } catch { Write-Host $_.Exception.Message -ForegroundColor Red; return }
Normalize-Window | Out-Null
Clear-FullTextSearch      # start clean so the Name filter isn't AND-limited by a stale search
Clear-NameFilter

Write-Host "Input (AR data): $SourceFolder"
Write-Host "Output (Excel) : $OutputFolder`n"

$rows = New-Object System.Collections.Generic.List[object]
$xlsFiles = Get-ChildItem -LiteralPath $SourceFolder -Filter *.xls -File | Sort-Object Name
$processed = 0
foreach ($xls in $xlsFiles) {
  try { $cust = Read-XlsCustomer $xls.FullName } catch { $cust = '' }
  if (-not $cust) {
    Write-Host ("SKIP (no name): {0}" -f $xls.Name) -ForegroundColor DarkGray
    $rows.Add([pscustomobject]@{Customer='';Frequency=''}); continue
  }
  try {
    $res = Lookup-Frequency $cust
    if ($res.Freq) {
      $tag = if ($res.Term -eq $cust) { '' } else { " (via '$($res.Term)')" }
      Write-Host ("{0,-40} -> {1}{2}" -f $cust, $res.Freq, $tag) -ForegroundColor Cyan
    } else {
      Write-Host ("{0,-40} -> (not found even after shortening)" -f $cust) -ForegroundColor Yellow
    }
    $rows.Add([pscustomobject]@{Customer=$cust;Frequency=$res.Freq})
    Clear-NameFilter
  } catch {
    Write-Host ("{0} -> ERROR {1}" -f $cust,$_.Exception.Message) -ForegroundColor Red
    $rows.Add([pscustomobject]@{Customer=$cust;Frequency=''})
    try { Clear-NameFilter } catch {}
  }
  $processed++
  if ($MaxFiles -gt 0 -and $processed -ge $MaxFiles) { Write-Host "`n(cap $MaxFiles reached)"; break }
}

# ---- write Excel (.xlsx): two columns, Customer + Frequency ----
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$xlsxPath = Join-Path $OutputFolder ("CustomerFrequency_$stamp.xlsx")
$excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
try {
  $wb = $excel.Workbooks.Add(); $ws = $wb.Sheets.Item(1); $ws.Name = 'Frequency'
  $headers = @('Customer','Frequency')
  for ($c=0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1,$c+1) = $headers[$c]; $ws.Cells.Item(1,$c+1).Font.Bold = $true }
  $rIdx = 2
  foreach ($row in $rows) {
    $ws.Cells.Item($rIdx,1) = [string]$row.Customer
    $ws.Cells.Item($rIdx,2) = [string]$row.Frequency
    $rIdx++
  }
  $ws.Columns.Item("A:B").AutoFit() | Out-Null
  $wb.SaveAs($xlsxPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
  $wb.Close($true) | Out-Null
} finally { $excel.Quit()|Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null }

# also a CSV copy (same two columns). UTF-8 so names like AVENDAÃ‘O survive.
$csvPath = Join-Path $OutputFolder ("CustomerFrequency_$stamp.csv")
$rows | Select-Object Customer, Frequency | Export-Csv -NoTypeInformation -LiteralPath $csvPath -Encoding UTF8

$found = @($rows | Where-Object { $_.Frequency -ne '' }).Count
$missing = $rows.Count - $found
Write-Host "`n================ DONE ================"
Write-Host ("TOTAL {0}   WITH FREQUENCY {1}   BLANK {2}" -f $rows.Count,$found,$missing)
Write-Host "Excel : $xlsxPath"
Write-Host "CSV   : $csvPath"



