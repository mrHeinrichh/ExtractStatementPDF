# ConvertTo-MDY : parse a statement date like "Mar. 02, ' 26" into "3/2/2026",
# the format Accounting's date fields expect.
function ConvertTo-MDY([string]$txt) {
  $t = ($txt -replace "'","" -replace "\.","" -replace ",","" -replace "\s+"," ").Trim()
  $parts = $t.Split(' '); if ($parts.Count -lt 3) { throw "Cannot parse date '$txt'" }
  $map = @{Jan=1;Feb=2;Mar=3;Apr=4;May=5;Jun=6;Jul=7;Aug=8;Sep=9;Oct=10;Nov=11;Dec=12}
  $m = $map[$parts[0].Substring(0,3)]; $d = [int]$parts[1]; $y = [int]$parts[2]
  if ($y -lt 100) { $y += 2000 }; return "$m/$d/$y"
}
