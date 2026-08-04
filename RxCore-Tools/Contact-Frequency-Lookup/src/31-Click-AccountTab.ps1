# Click-AccountTab : the Frequency field lives on the Account tab. Click the
# LEFTMOST "Account" tab (the one just right of GENERIC), then press ESC to
# dismiss any dropdown that a stray click may have opened (does not change tabs).
function Click-AccountTab {
  $root = $AE::FromHandle((Get-Contact).MainWindowHandle)
  $accs = $root.FindAll($TS,(New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty,"Account")))
  $best = $null
  foreach ($a in $accs) { if ($a.Current.ClassName -eq 'TextBlock') { if (-not $best -or $a.Current.BoundingRectangle.X -lt $best.Current.BoundingRectangle.X) { $best = $a } } }
  if ($best) {
    $b = $best.Current.BoundingRectangle
    Click ([int]($b.X + $b.Width/2)) ([int]($b.Y + $b.Height/2))
    Start-Sleep -Milliseconds 500
    Send '{ESC}'
    Start-Sleep -Milliseconds 200
  }
}
