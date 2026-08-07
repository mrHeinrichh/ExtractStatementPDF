# WordWrap : word-bounded form of a KeyOf() value for safe fuzzy matching.
# Non-alphanumerics become spaces and it's wrapped in spaces, so .Contains(" ESCA ")
# matches the WORD "ESCA" only - never the "ESCA" inside "LESCANO".
function WordWrap([string]$k) {
  if (-not $k) { return ' ' }
  return ' ' + ((($k -replace '[^A-Z0-9]',' ') -replace '\s+',' ').Trim()) + ' '
}
