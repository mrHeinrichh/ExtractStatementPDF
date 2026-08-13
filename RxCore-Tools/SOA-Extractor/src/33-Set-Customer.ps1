# Set-Customer : type the customer id (falls back to name when no id is available) into
# the Customers box, wait for the autocomplete suggestion, then press TAB to commit it
# (binds "SOA-#|NAME").
function Set-Customer([string]$idOrName) {
  $xy = Field-XY 'Customers'; Click $xy[0] $xy[1]; Start-Sleep -Milliseconds 250
  Send '^a'; Start-Sleep -Milliseconds 80; Send '{DEL}'; Start-Sleep -Milliseconds 150
  Send-Literal $idOrName; Start-Sleep -Milliseconds 1500; Send '{TAB}'; Start-Sleep -Milliseconds 500
}
