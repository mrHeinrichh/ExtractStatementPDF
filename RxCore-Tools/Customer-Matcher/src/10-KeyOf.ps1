# KeyOf : normalize a name for matching (collapse spaces, upper-case).
function KeyOf([string]$s) { return ($s -replace '\s+',' ').Trim().ToUpperInvariant() }
