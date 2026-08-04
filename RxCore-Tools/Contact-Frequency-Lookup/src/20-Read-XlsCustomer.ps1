# Read-XlsCustomer : open an .xls read-only (Excel COM) and return the customer
# name from cell E9.
function Read-XlsCustomer([string]$path) {
  $x = New-Object -ComObject Excel.Application; $x.Visible=$false; $x.DisplayAlerts=$false
  try {
    $wb=$x.Workbooks.Open($path,[Type]::Missing,$true); $ws=$wb.Sheets.Item(1)
    $name = ([string]$ws.Cells.Item(9,5).Text).Trim()
    $wb.Close($false) | Out-Null; return $name
  } finally { $x.Quit()|Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)|Out-Null }
}
