# Send-Literal : type text literally, escaping SendKeys' special characters
# ( + ^ % ~ ( ) { } [ ] ) so names/dates are entered exactly as given.
function Send-Literal([string]$t) { Send ($t -replace '([+^%~(){}\[\]])', '{$1}') }
