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
