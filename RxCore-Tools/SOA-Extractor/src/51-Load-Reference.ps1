# Load-Reference : read the reference customer list into two lookups:
#   ByKey   : KeyOf(RxOfficeName) -> @{ Name=<RxOffice name>; Freq; Tokens; Id }  (exact + word-prefix)
#   ByTight : TightKey(x)         -> @{ Name=<RxOffice name>; Freq; Id }          (punctuation/spacing-
#             insensitive; includes the RxOfficeName AND its ARName alias)
# A match gives us the id to type in Accounting's Customers search box (falls back to the
# RxOffice name when no id column is present), plus the frequency.
#
# Accepts:
#   * RxOffice + ARName - comma CSV: Id,Code,RxOfficeName,StatementFrequency,ARName
#                         (ARName is the AR file/customer alias, e.g. "AbalosGuillermoOptical";
#                         Code, the 2nd column, is the id typed into the Customers search box)
#   * RxOffice          - comma CSV: Id,Code,Name,StatementFrequency
#   * Contact export    - UTF-16, TAB-delimited: Name / Alias / StatementFrequency (no id column)
#   * Simple            - comma CSV: Customer,Frequency (no id column)
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
    param($name, $freq, $alias, $id)
    if (-not $name -or -not $freq) { return }
    $entry = @{ Name=$name; Freq=$freq; Tokens=(Tokenize $name); Id=$id }
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
      & $add $nm $c[$iFreq] $al ''
    }
  } else {
    # ---- comma CSV: RxOffice(+ARName) / RxOffice / simple ----
    $objs = ($lines | Where-Object { $_.Trim() }) | ConvertFrom-Csv
    if ($objs) {
      $cols = $objs[0].PSObject.Properties.Name
      $nameCol = if ($cols -contains 'RxOfficeName') { 'RxOfficeName' } elseif ($cols -contains 'Name') { 'Name' } elseif ($cols -contains 'Customer') { 'Customer' } else { $null }
      $freqCol = if ($cols -contains 'StatementFrequency') { 'StatementFrequency' } elseif ($cols -contains 'Frequency') { 'Frequency' } else { $null }
      $aliasCol = if ($cols -contains 'ARName') { 'ARName' } else { $null }
      $idCol = if ($cols -contains 'Code') { 'Code' } else { $null }
      if ($nameCol -and $freqCol) {
        foreach ($o in $objs) {
          $nm = ([string]$o.$nameCol).Trim()
          $f  = ([string]$o.$freqCol).Trim()
          $al = if ($aliasCol) { ([string]$o.$aliasCol).Trim() } else { '' }
          $id = if ($idCol) { ([string]$o.$idCol).Trim() } else { '' }
          & $add $nm $f $al $id
        }
      }
    }
  }

  $result.Source = $path
  return $result
}
