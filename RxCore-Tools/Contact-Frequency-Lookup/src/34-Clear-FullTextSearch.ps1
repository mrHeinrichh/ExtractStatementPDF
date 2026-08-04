# Clear-FullTextSearch : empty the bottom "Full Text Search" box. A leftover value
# there AND-combines with the Name column filter and hides real matches, so we
# clear it once at startup.
function Clear-FullTextSearch {
  $root = $AE::FromHandle((Get-Contact).MainWindowHandle)
  $sb = $root.FindFirst($TS,(New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty,"PART_SearchAsYouTypeTextBox")))
  if ($sb) {
    $b = $sb.Current.BoundingRectangle
    Click ([int]($b.X + $b.Width/2)) ([int]($b.Y + $b.Height/2))
    Start-Sleep -Milliseconds 150; Send '^a'; Send '{DEL}'; Start-Sleep -Milliseconds 150; Send '{ENTER}'; Start-Sleep -Milliseconds 600
  }
}
