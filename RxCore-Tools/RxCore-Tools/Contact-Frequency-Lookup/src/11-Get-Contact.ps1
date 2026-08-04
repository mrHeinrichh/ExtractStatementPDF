# Get-Contact : return the running Contact.exe process, or throw if not running.
function Get-Contact {
  $p = Get-Process Contact -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $p) { throw "Contact.exe is not running. Open it and log in first." }
  return $p
}
