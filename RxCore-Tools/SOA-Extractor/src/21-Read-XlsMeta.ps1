# Read-XlsMeta : open an .xls (read-only, via Excel COM) and return the
# customer name (cell E9) and the AS-OF From/To dates (E18 / H18) as M/D/YYYY.
function Read-XlsMeta([string]$path) {
  $x = New-Object -ComObject Excel.Application; $x.Visible=$false; $x.DisplayAlerts=$false
  try {
    $wb = $x.Workbooks.Open($path, [Type]::Missing, $true); $ws = $wb.Sheets.Item(1)
    $m = [ordered]@{
      Customer = ([string]$ws.Cells.Item(9,5).Text).Trim()
      From = ConvertTo-MDY ([string]$ws.Cells.Item(18,6).Text)
      To   = ConvertTo-MDY ([string]$ws.Cells.Item(18,8).Text)
    }
    $wb.Close($false) | Out-Null; return $m
  } finally { $x.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
}
