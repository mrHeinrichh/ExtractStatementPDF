# Get-MatchCount : read the "selected / TOTAL" indicator at the top-right and
# return TOTAL (the number just right of the "/"). Tells us how many grid rows
# the current filter matched (0 = not found, 1 = exact, >1 = ambiguous). -1 on error.
function Get-MatchCount {
  $root = $AE::FromHandle((Get-Contact).MainWindowHandle)
  $txtCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty,[System.Windows.Automation.ControlType]::Text)
  $txts = $root.FindAll($TS,$txtCond)
  $slash=$null
  foreach($t in $txts){ $b=$t.Current.BoundingRectangle; if($t.Current.Name -eq '/' -and $b.Y -lt 90){ $slash=$t; break } }
  if (-not $slash) { return -1 }
  $sy=$slash.Current.BoundingRectangle.Y; $sx=$slash.Current.BoundingRectangle.X
  $best=$null;$bd=99999
  foreach($t in $txts){ $b=$t.Current.BoundingRectangle; if([math]::Abs($b.Y-$sy)-le 4 -and $b.X -gt $sx){ $dx=$b.X-$sx; if($dx-lt $bd){$bd=$dx;$best=$t} } }
  if(-not $best){ return -1 }
  $n=0; if([int]::TryParse(($best.Current.Name -replace '[^\d]',''),[ref]$n)){ return $n }
  return -1
}
