# KeyOf : normalize a customer name for matching (collapse spaces, upper-case)
# so the .xls name and the frequency-table name compare reliably.
function KeyOf([string]$s) { return ($s -replace '\s+',' ').Trim().ToUpperInvariant() }
