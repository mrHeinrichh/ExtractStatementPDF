# Resolve-Frequency : find a customer's frequency in the reference map.
# Tries an exact (normalized) match first; if none, tries progressively shorter
# forms of the name and takes the first table entry that CONTAINS that form.
# Returns @{ Freq = <frequency or ''>; Via = 'exact' | <term used> | '' }.
function Resolve-Frequency([string]$name, [hashtable]$map) {
  if (-not $map -or $map.Count -eq 0) { return @{ Freq=''; Via='' } }
  $k = KeyOf $name
  if ($map.ContainsKey($k)) { return @{ Freq=$map[$k]; Via='exact' } }
  foreach ($term in (Get-Candidates $name)) {
    $tk = KeyOf $term
    if (-not $tk) { continue }
    $hit = $null
    foreach ($key in $map.Keys) { if ($key.Contains($tk)) { $hit = $key; break } }
    if ($hit) { return @{ Freq=$map[$hit]; Via=$term } }
  }
  return @{ Freq=''; Via='' }
}
