# Clear-NameFilter : empty the Name column filter box and re-apply, so the grid
# resets to the full list before the next search.
function Clear-NameFilter {
  $r = WinRect
  Click ($r.Left + $NAME_FILTER_DX) ($r.Top + $NAME_FILTER_DY)
  Start-Sleep -Milliseconds 150; Send '^a'; Send '{DEL}'; Start-Sleep -Milliseconds 150; Send '{ENTER}'; Start-Sleep -Milliseconds 400
}
