# Paste : put text on the clipboard and Ctrl+V it. Used for the search term so
# the whole string (incl. accented names like AVENDAÑO) lands reliably - the
# search-as-you-type box drops characters when typed key by key.
function Paste([string]$t){ [System.Windows.Forms.Clipboard]::SetText($t); Start-Sleep -Milliseconds 150; Send '^v' }
