# Freq-Index : map a frequency name to its position in the Invoice Frequency
# dropdown (0=Monthly, 1=Daily, 2=Weekly, 3=Bi-Weekly). Returns -1 if unknown.
# Accepts spelling variants from the Contact table (e.g. "Bi-Weekly"/"BiWeekly").
function Freq-Index([string]$value) {
  switch (($value -replace '\s','').ToUpperInvariant()) {
    'MONTHLY'     { return 0 }
    'DAILY'       { return 1 }
    'WEEKLY'      { return 2 }
    'BI-WEEKLY'   { return 3 }
    'BIWEEKLY'    { return 3 }
    'FORTNIGHTLY' { return 3 }
    default       { return -1 }
  }
}
