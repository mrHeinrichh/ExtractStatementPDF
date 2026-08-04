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
