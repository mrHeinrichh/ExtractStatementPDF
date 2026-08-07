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
