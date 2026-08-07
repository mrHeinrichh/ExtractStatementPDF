# Read-XlsNames : return the customer name (cell E9) from every .xls in a folder,
# read-only via Excel COM, sorted by file name.
function Read-XlsNames([string]$folder) {
  $names = New-Object System.Collections.Generic.List[string]
  $excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
  try {
    foreach ($f in (Get-ChildItem -LiteralPath $folder -Filter *.xls -File | Sort-Object Name)) {
      try {
        $wb = $excel.Workbooks.Open($f.FullName,[Type]::Missing,$true); $ws = $wb.Sheets.Item(1)
        $names.Add(([string]$ws.Cells.Item(9,5).Text).Trim()); $wb.Close($false) | Out-Null
      } catch {}
    }
  } finally { $excel.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
  return $names
}
