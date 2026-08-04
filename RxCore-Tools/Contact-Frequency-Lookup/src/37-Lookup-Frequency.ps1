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
