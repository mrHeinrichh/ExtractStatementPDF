# Read-Frequency : read the Invoice Frequency value from the Account tab. Finds
# the "Frequency" label, then the editable combo box immediately to its right and
# returns its value via the UI Automation ValuePattern. '' if not found.
function Read-Frequency {
  $root = $AE::FromHandle((Get-Contact).MainWindowHandle)
  $lbl = $root.FindFirst($TS,(New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty,"Frequency")))
  if(-not $lbl){ return '' }
  $ly=$lbl.Current.BoundingRectangle.Y; $lx=$lbl.Current.BoundingRectangle.X
  $edits=$root.FindAll($TS,(New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty,"PART_EditableTextBox")))
  $best=$null;$bd=99999
  foreach($e in $edits){ $b=$e.Current.BoundingRectangle; if([math]::Abs($b.Y-$ly)-le 6 -and $b.X -gt $lx){ $dx=$b.X-$lx; if($dx-lt $bd){$bd=$dx;$best=$e} } }
  if(-not $best){ return '' }
  try { return $best.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value } catch { return '' }
}
