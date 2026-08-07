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
