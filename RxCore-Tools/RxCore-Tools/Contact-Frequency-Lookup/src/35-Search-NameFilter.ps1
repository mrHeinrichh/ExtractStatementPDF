# Search-NameFilter : type a term into the Name column filter (paste + Enter) and
# return the resulting match count.
function Search-NameFilter([string]$term) {
  $r = WinRect
  Click ($r.Left + $NAME_FILTER_DX) ($r.Top + $NAME_FILTER_DY)
  Start-Sleep -Milliseconds 200; Send '^a'; Send '{DEL}'; Start-Sleep -Milliseconds 150
  Paste $term; Start-Sleep -Milliseconds 400; Send '{ENTER}'; Start-Sleep -Milliseconds 1400
  return (Get-MatchCount)
}
