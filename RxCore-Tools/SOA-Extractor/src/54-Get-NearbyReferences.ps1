# Get-NearbyReferences : RxOffice names "near" a customer, tried when the matched name
# produces no statement - e.g. AR "GOCO OPTICAL" -> "GOCO OPTICAL INDIGENT".
# A nearby name starts with every word of the AR name or of the matched RxOffice name
# (>= 2 words) and has extra words after them. Closest first (fewest extra words),
# at most $MaxNearbyNames. Returns @( @{ Name; Id; Freq; Extra } ... ).
function Get-NearbyReferences([string]$arName, [string]$matchedName, $ref) {
  if (-not $ref -or -not $ref.ByKey -or $MaxNearbyNames -le 0) { return @() }
  $bases = New-Object System.Collections.Generic.List[object]
  foreach ($n in @($arName, $matchedName)) { $t = @(Tokenize $n); if ($t.Count -ge 2) { $bases.Add($t) } }
  if ($bases.Count -eq 0) { return @() }

  $skip = KeyOf $matchedName
  $found = @{}
  foreach ($e in $ref.ByKey.Values) {
    $t = @($e.Tokens)
    $k = KeyOf $e.Name
    if ($k -eq $skip -or $found.ContainsKey($k)) { continue }
    foreach ($b in $bases) {
      if ($t.Count -gt $b.Count -and (Test-Prefix $b $t)) {
        $found[$k] = [pscustomobject]@{ Name=$e.Name; Id=$e.Id; Freq=$e.Freq; Extra=($t.Count - $b.Count) }
        break
      }
    }
  }
  return @($found.Values | Sort-Object Extra, Name | Select-Object -First $MaxNearbyNames)
}
