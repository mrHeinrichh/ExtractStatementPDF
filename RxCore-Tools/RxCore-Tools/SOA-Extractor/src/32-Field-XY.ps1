# Field-XY : the dialog has no automatable child controls, so each field is
# located by a fixed offset from the dialog window's top-left corner.
# Returns @(screenX, screenY) for the named field/button.
function Field-XY([string]$name) {
  $d = Get-Dialog; if (-not $d) { throw "dialog not open" }
  $r = $d.Current.BoundingRectangle; $dx = [int]$r.X; $dy = [int]$r.Y
  switch ($name) {
    'Customers'        { return @(($dx+240),($dy+43)) }
    'InvoiceFrequency' { return @(($dx+200),($dy+104)) }
    'From'             { return @(($dx+200),($dy+122)) }
    'To'               { return @(($dx+200),($dy+147)) }
    'OK'               { return @(($dx+247),($dy+407)) }
    'No'               { return @(($dx+328),($dy+407)) }
    default            { throw "unknown field $name" }
  }
}
