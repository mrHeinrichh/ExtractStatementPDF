# ===================== src\00-Header.ps1 =====================
<#
  Customer Matcher  -  assembled source
  =====================================
  One function per file under src\. build.ps1 concatenates them (filename order)
  and compiles Customer-Matcher.exe with ps2exe.

  What it does (pure data - no Accounting/Contact needed):
    * Reads the customer name (cell E9) from every .xls in a folder you choose.
    * Reads the Customer CSV you exported from Contact (has Name, Alias,
      StatementFrequency).
    * Matches each .xls name to a Contact row (exact, then word-boundary fuzzy).
    * Writes "Customer_matched.csv" next to the export, identical to it but with a
      new "MatchedXlsName" column inserted right after "Name", filled with the
      .xls customer name(s) that matched that row.

  Requires: Microsoft Excel installed (to read the .xls files). Contact/Accounting
  do NOT need to be open.
#>

$ErrorActionPreference = 'Stop'


# ===================== src\02-Interop.ps1 =====================
# ---- Assemblies needed (WinForms for the dialogs). No Win32/UIA required. ----
Add-Type -AssemblyName System.Windows.Forms, System.Drawing


# ===================== src\03-Ui-Dialogs.ps1 =====================
# ---- GUI dialogs (message, folder paste-box, file browse) -----------------

function Show-Message([string]$text, [string]$title = 'Customer Matcher') {
  [System.Windows.Forms.MessageBox]::Show($text, $title,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

# Pick-Folder : paste-a-path dialog (with Browse...) for the .xls folder.
function Pick-Folder {
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Customer Matcher"; $form.Size = New-Object System.Drawing.Size(620,190)
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

# Browse-File : open-file dialog for the Contact export CSV. Starts in Documents.
function Browse-File([string]$title) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $title
  $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
  $docs = [Environment]::GetFolderPath('MyDocuments')
  if ($docs -and (Test-Path $docs)) { $dlg.InitialDirectory = $docs }
  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  return $dlg.FileName
}


# ===================== src\10-KeyOf.ps1 =====================
# KeyOf : normalize a name for matching (collapse spaces, upper-case).
function KeyOf([string]$s) { return ($s -replace '\s+',' ').Trim().ToUpperInvariant() }


# ===================== src\11-WordWrap.ps1 =====================
# WordWrap : word-bounded form of a KeyOf() value for safe fuzzy matching.
# Non-alphanumerics become spaces and it's wrapped in spaces, so .Contains(" ESCA ")
# matches the WORD "ESCA" only - never the "ESCA" inside "LESCANO".
function WordWrap([string]$k) {
  if (-not $k) { return ' ' }
  return ' ' + ((($k -replace '[^A-Z0-9]',' ') -replace '\s+',' ').Trim()) + ' '
}


# ===================== src\12-Get-Candidates.ps1 =====================
# Get-Candidates : progressively shorter forms of a name for fuzzy matching -
# full name first, then drop trailing words, then trim letters off the first word.
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


# ===================== src\20-Read-XlsNames.ps1 =====================
# Read-XlsNames : return the customer name (cell E9) from every .xls in a folder,
# read-only via Excel COM, sorted by file name.
function Read-XlsNames([string]$folder) {
  $names = New-Object System.Collections.Generic.List[string]
  $excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
  try {
    foreach ($f in (Get-ChildItem -LiteralPath $folder -Filter *.xls -File | Sort-Object Name)) {
      try {
        $wb = $excel.Workbooks.Open($f.FullName,[Type]::Missing,$true); $ws = $wb.Sheets.Item(1)
        $names.Add(([string]$ws.Cells.Item(9,5).Text).Trim()); $wb.Close($false) | Out-Null
      } catch {}
    }
  } finally { $excel.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
  return $names
}


# ===================== src\30-Read-CustomerExport.ps1 =====================
# Read-CustomerExport : read the Contact "Export" CSV (UTF-16, tab-delimited),
# even if it's open in Excel (shared read). Returns:
#   @{ HdrFields=<raw header fields>; Data=<List of raw field arrays>;
#      INameIdx; IAliasIdx; IFreqIdx }
# Raw fields keep their surrounding quotes so the file can be rewritten unchanged.
function Read-CustomerExport([string]$path) {
  $fs = New-Object System.IO.FileStream($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
  $rd = New-Object System.IO.BinaryReader($fs)
  $bytes = $rd.ReadBytes([int]$fs.Length); $rd.Close(); $fs.Close()
  $enc = [System.Text.Encoding]::UTF8
  if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [System.Text.Encoding]::Unicode }
  elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $enc = [System.Text.Encoding]::BigEndianUnicode }
  $textAll = $enc.GetString($bytes)
  if ($textAll.Length -gt 0 -and $textAll[0] -eq [char]0xFEFF) { $textAll = $textAll.Substring(1) }  # drop decoded BOM (writer adds its own)
  $lines = $textAll -split "\r?\n"
  if ($lines.Count -lt 2) { throw "The reference CSV looks empty: $path" }

  $hdrFields = $lines[0].Split("`t")
  $hdr = $hdrFields | ForEach-Object { $_.Trim().Trim('"') }
  $iName = [Array]::IndexOf($hdr,'Name')
  if ($iName -lt 0) { throw "No 'Name' column found - is this the Contact Export CSV?" }

  $data = New-Object System.Collections.Generic.List[object]
  for ($i=1; $i -lt $lines.Count; $i++) { if ($lines[$i].Trim()) { $data.Add($lines[$i].Split("`t")) } }

  return @{
    HdrFields = $hdrFields
    Data      = $data
    INameIdx  = $iName
    IAliasIdx = [Array]::IndexOf($hdr,'Alias')
    IFreqIdx  = [Array]::IndexOf($hdr,'StatementFrequency')
    IMatchIdx = [Array]::IndexOf($hdr,'MatchedXlsName')   # -1 if not annotated yet
    Encoding  = $enc
  }
}


# ===================== src\40-Match-Names.ps1 =====================
# Match-Names : match each extracted .xls name to a Contact row.
#   exact (Name or Alias key equals the .xls key), else the first row that
#   contains a progressively-shorter form of the name AS A WHOLE WORD.
# Returns @{ Matched = <hashtable rowIndex -> List[string of xls names]>;
#            Unmatched = <List[string]> }.
function Match-Names($parsed, $xlsNames) {
  $data = $parsed.Data; $iName = $parsed.INameIdx; $iAlias = $parsed.IAliasIdx
  $n = $data.Count

  # Pre-compute per-row keys + word-wrapped forms (once).
  $rowNW = New-Object string[] $n; $rowAW = New-Object string[] $n
  $nameKeyIdx = @{}; $aliasKeyIdx = @{}
  for ($j=0; $j -lt $n; $j++) {
    $nm = $data[$j][$iName].Trim().Trim('"'); $nk = KeyOf $nm
    $rowNW[$j] = WordWrap $nk
    if ($nk -and -not $nameKeyIdx.ContainsKey($nk)) { $nameKeyIdx[$nk] = $j }
    $al = ''; if ($iAlias -ge 0 -and $iAlias -lt $data[$j].Count) { $al = $data[$j][$iAlias].Trim().Trim('"') }
    $ak = KeyOf $al; $rowAW[$j] = WordWrap $ak
    if ($ak -and -not $aliasKeyIdx.ContainsKey($ak)) { $aliasKeyIdx[$ak] = $j }
  }

  $matched = @{}; $unmatched = New-Object System.Collections.Generic.List[string]
  foreach ($xn in $xlsNames) {
    if (-not $xn) { continue }
    $k = KeyOf $xn; $idx = -1
    if     ($nameKeyIdx.ContainsKey($k))  { $idx = $nameKeyIdx[$k] }
    elseif ($aliasKeyIdx.ContainsKey($k)) { $idx = $aliasKeyIdx[$k] }
    else {
      foreach ($term in (Get-Candidates $xn)) {
        $tw = WordWrap (KeyOf $term)
        if ($tw.Trim().Length -lt 4) { continue }
        for ($jj=0; $jj -lt $n; $jj++) {
          if ($rowNW[$jj].Contains($tw) -or $rowAW[$jj].Contains($tw)) { $idx = $jj; break }
        }
        if ($idx -ge 0) { break }
      }
    }
    if ($idx -ge 0) {
      if (-not $matched.ContainsKey($idx)) { $matched[$idx] = New-Object System.Collections.Generic.List[string] }
      $matched[$idx].Add($xn)
    } else { $unmatched.Add($xn) }
  }
  return @{ Matched=$matched; Unmatched=$unmatched }
}


# ===================== src\50-Write-Annotated.ps1 =====================
# Write-Annotated : write the customer file back WITH the matched names in a
# "MatchedXlsName" column placed immediately after "Name".
#   * Idempotent - if the file already has a MatchedXlsName column (from a previous
#     run) its values are replaced, not a second column added.
#   * Writes in the file's original encoding, tab-delimited.
#   * Atomic - writes a temp file then replaces, so a failure can't truncate the
#     original. Throws if the target is locked (open in Excel).
function Write-Annotated($parsed, $matched, [string]$outPath) {
  $data   = $parsed.Data
  $iName  = $parsed.INameIdx
  $iMatch = $parsed.IMatchIdx           # existing column index, or -1
  $enc    = if ($parsed.Encoding) { $parsed.Encoding } else { [System.Text.Encoding]::Unicode }
  $insertAt = $iName + 1

  function MatchValue($j) { if ($matched.ContainsKey($j)) { '"' + ($matched[$j] -join '; ') + '"' } else { '""' } }

  $sb = New-Object System.Text.StringBuilder
  if ($iMatch -ge 0) {
    # Replace the existing MatchedXlsName column in place (header stays as-is).
    [void]$sb.AppendLine($parsed.HdrFields -join "`t")
    for ($j=0; $j -lt $data.Count; $j++) {
      $row = New-Object System.Collections.Generic.List[string]; $row.AddRange([string[]]$data[$j])
      while ($row.Count -le $iMatch) { $row.Add('""') }
      $row[$iMatch] = MatchValue $j
      [void]$sb.AppendLine($row -join "`t")
    }
  } else {
    # Insert a new column right after Name.
    $nh = New-Object System.Collections.Generic.List[string]; $nh.AddRange([string[]]$parsed.HdrFields); $nh.Insert($insertAt, '"MatchedXlsName"')
    [void]$sb.AppendLine($nh -join "`t")
    for ($j=0; $j -lt $data.Count; $j++) {
      $row = New-Object System.Collections.Generic.List[string]; $row.AddRange([string[]]$data[$j])
      $row.Insert($insertAt, (MatchValue $j))
      [void]$sb.AppendLine($row -join "`t")
    }
  }

  # Atomic replace: write temp beside the target, then move over it.
  $tmp = "$outPath.tmp"
  [System.IO.File]::WriteAllText($tmp, $sb.ToString(), $enc)
  try {
    Move-Item -LiteralPath $tmp -Destination $outPath -Force
  } catch {
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    throw "Could not write '$outPath'. If it's open in Excel, close it and run again.  ($($_.Exception.Message))"
  }
}


# ===================== src\99-Main.ps1 =====================
# =====================  MAIN PROGRAM  =====================
# Test hooks (optional): CM_XLS (folder), CM_REF (customer csv), CM_OUT (write target).

# 1) folder of .xls files
if ($env:CM_XLS) { $xlsFolder = $env:CM_XLS } else { $xlsFolder = Pick-Folder }
if (-not $xlsFolder) { Write-Host "No folder selected. Exiting."; return }
if (-not (Test-Path -LiteralPath $xlsFolder)) { Write-Host "Folder not found: $xlsFolder"; return }

# 2) the Customer file exported from Contact (browse; opens in Documents)
if ($env:CM_REF) { $refPath = $env:CM_REF } else { $refPath = Browse-File "Select the Customer file exported from Contact" }
if (-not $refPath) { Write-Host "No Customer file selected. Exiting."; return }

# 3) write target = the SAME file (in place). CM_OUT can redirect it (used for tests).
$outPath = if ($env:CM_OUT) { $env:CM_OUT } else { $refPath }

Write-Host "xls folder    : $xlsFolder"
Write-Host "Customer file : $refPath"
Write-Host "Writing to    : $outPath  (in place)`n"

$parsed = Read-CustomerExport $refPath
Write-Host ("Customer rows : {0}   (Name col {1}{2})" -f $parsed.Data.Count, $parsed.INameIdx, $(if ($parsed.IMatchIdx -ge 0) { '; updating existing MatchedXlsName column' } else { '' }))

$xlsNames = Read-XlsNames $xlsFolder
Write-Host ("Extracted .xls names : {0}" -f $xlsNames.Count)

$res = Match-Names $parsed $xlsNames
Write-Host ("Matched rows  : {0}   Unmatched .xls names : {1}" -f $res.Matched.Count, $res.Unmatched.Count)

try {
  Write-Annotated $parsed $res.Matched $outPath
  Write-Host "`nDone. Matched names written into: $outPath" -ForegroundColor Green
} catch {
  Write-Host "`n$($_.Exception.Message)" -ForegroundColor Red
  Show-Message $_.Exception.Message
  return
}

if ($res.Unmatched.Count) {
  Write-Host "`nUnmatched .xls names (no Contact row found):" -ForegroundColor Yellow
  $res.Unmatched | ForEach-Object { Write-Host "  $_" }
}



