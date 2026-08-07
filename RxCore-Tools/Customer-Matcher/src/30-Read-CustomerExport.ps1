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
