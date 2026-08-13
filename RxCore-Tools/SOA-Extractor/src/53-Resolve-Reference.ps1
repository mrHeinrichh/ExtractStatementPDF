# Resolve-Reference : match an AR (.xls) customer name to a reference row and
# return the RxOffice NAME (for logging/reporting), the id to type into Accounting's
# Customers search box (2nd/"Code" column of the reference CSV; empty if the reference
# has no id column), and the frequency.
#
# Matching, most-precise first:
#   1) exact         - normalized key equal to an RxOfficeName.
#   2) tight / ARName - punctuation/spacing/case-insensitive equal to an RxOfficeName
#                       OR its ARName alias. This is how "ABALOS GUILLERMO OPTICAL"
#                       matches ARName "AbalosGuillermoOptical" -> resolves to that row.
#   3) word-prefix    - one name's words are the leading run of the other's (>= 2 words),
#                       e.g. AR "FESAR ... GREENHILLS" <-> "FESAR ... GREENHILLS SHOPPING CENTER".
# Returns @{ Name=<RxOffice name or ''>; Id=<code or ''>; Freq; Via='exact'|'ARname'|"ref '<name>'"|'' }.
function Resolve-Reference([string]$name, $ref) {
  if (-not $ref) { return @{ Name=''; Id=''; Freq=''; Via='' } }
  $byKey = $ref.ByKey; $byTight = $ref.ByTight
  if (-not $byKey -or $byKey.Count -eq 0) { return @{ Name=''; Id=''; Freq=''; Via='' } }

  $k = KeyOf $name
  if ($byKey.ContainsKey($k)) { $e = $byKey[$k]; return @{ Name=$e.Name; Id=$e.Id; Freq=$e.Freq; Via='exact' } }

  $tk = TightKey $name
  if ($byTight.ContainsKey($tk)) { $e = $byTight[$tk]; return @{ Name=$e.Name; Id=$e.Id; Freq=$e.Freq; Via='ARname' } }

  $ar = Tokenize $name
  if ($ar.Count -eq 0) { return @{ Name=''; Id=''; Freq=''; Via='' } }
  $best = $null; $bestScore = -1
  foreach ($e in $byKey.Values) {
    $t = $e.Tokens
    if (-not $t -or $t.Count -eq 0) { continue }
    $score = -1
    if     (Test-Prefix $t $ar) { if ($t.Count  -ge 2) { $score = ($t.Count  * 1000) - ($ar.Count - $t.Count) } }   # ref shorter
    elseif (Test-Prefix $ar $t) { if ($ar.Count -ge 2) { $score = ($ar.Count * 1000) - ($t.Count - $ar.Count) } }   # AR shorter
    if ($score -gt $bestScore) { $bestScore = $score; $best = $e }
  }
  if ($best) { return @{ Name=$best.Name; Id=$best.Id; Freq=$best.Freq; Via=("ref '{0}'" -f $best.Name) } }
  return @{ Name=''; Id=''; Freq=''; Via='' }
}
