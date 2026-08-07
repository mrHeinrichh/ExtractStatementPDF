# WordWrap : normalize a KeyOf() value to a word-bounded form for safe fuzzy
# matching. Non-alphanumerics become spaces and the whole thing is wrapped in
# spaces, so .Contains(" ESCA ") matches the WORD "ESCA" only - never the "ESCA"
# buried inside "LESCANO". Used by Resolve-Frequency.
function WordWrap([string]$k) {
  if (-not $k) { return ' ' }
  return ' ' + ((($k -replace '[^A-Z0-9]',' ') -replace '\s+',' ').Trim()) + ' '
}

# Tokenize : split a name into UPPERCASE alphanumeric words (punctuation dropped),
# used for word-level prefix matching between AR names and RxOffice names.
#   "ABALOS GUILLERMO OPTICAL" -> @('ABALOS','GUILLERMO','OPTICAL')
function Tokenize([string]$s) {
  if (-not $s) { return @() }
  return @((($s.ToUpperInvariant() -replace '[^A-Z0-9]',' ') -split '\s+') | Where-Object { $_ -ne '' })
}

# Test-Prefix : is $short the leading run of words of $long? (both string[])
function Test-Prefix($short, $long) {
  if ($short.Count -eq 0 -or $short.Count -gt $long.Count) { return $false }
  for ($i = 0; $i -lt $short.Count; $i++) { if ($short[$i] -ne $long[$i]) { return $false } }
  return $true
}

# TightKey : strip EVERYTHING except letters/digits and upper-case, so different
# spacing/punctuation/casing collapse to one key. This is how an AR name matches
# the reference's ARName column:
#   "ABALOS GUILLERMO OPTICAL"  ->  "ABALOSGUILLERMOOPTICAL"
#   "AbalosGuillermoOptical"    ->  "ABALOSGUILLERMOOPTICAL"
function TightKey([string]$s) {
  if (-not $s) { return '' }
  return ($s -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
}
