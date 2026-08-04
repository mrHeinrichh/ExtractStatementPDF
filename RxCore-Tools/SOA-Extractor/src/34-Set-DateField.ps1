# Set-DateField : clear a date field ('From' or 'To') and type an M/D/YYYY value.
function Set-DateField([string]$name, [string]$value) {
  $xy = Field-XY $name; Click $xy[0] $xy[1]; Start-Sleep -Milliseconds 200
  Send '^a'; Start-Sleep -Milliseconds 80; Send '{DEL}'; Start-Sleep -Milliseconds 80
  Send-Literal $value; Start-Sleep -Milliseconds 200; Send '{TAB}'; Start-Sleep -Milliseconds 200
}
