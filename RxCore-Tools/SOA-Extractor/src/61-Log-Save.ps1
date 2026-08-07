# Save-Log : persist the run log as BOTH a CSV (source of truth for resuming) and
# a formatted Excel workbook (_SOA_Log.xlsx) for the user to read. Rows already
# come sorted by the caller. Columns:
#   Customer | File | DateRange | Frequency | Status | Reason | Csv | UpdatedOn
# Status is DONE / ERROR / SKIPPED. Errors are visible in the Status/Reason cols.
function Save-Log($rows, [string]$logCsv, [string]$logXlsx) {
  # CSV (UTF-8, so accented names survive) - used for resume next run.
  $rows | Select-Object Customer, RxOfficeName, File, DateRange, Frequency, Status, Reason, Csv, UpdatedOn |
          Export-Csv -NoTypeInformation -LiteralPath $logCsv -Encoding UTF8

  # Excel workbook for viewing, with a bold header and coloured Status cells.
  $excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
  try {
    $wb = $excel.Workbooks.Add(); $ws = $wb.Sheets.Item(1); $ws.Name = 'SOA Log'
    $headers = @('Customer','RxOfficeName','File','DateRange','Frequency','Status','Reason','Csv','UpdatedOn')
    for ($c=0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1,$c+1) = $headers[$c]; $ws.Cells.Item(1,$c+1).Font.Bold = $true }
    $r = 2
    foreach ($row in $rows) {
      $ws.Cells.Item($r,1) = [string]$row.Customer
      $ws.Cells.Item($r,2) = [string]$row.RxOfficeName
      $ws.Cells.Item($r,3) = [string]$row.File
      $ws.Cells.Item($r,4) = [string]$row.DateRange
      $ws.Cells.Item($r,5) = [string]$row.Frequency
      $ws.Cells.Item($r,6) = [string]$row.Status
      $ws.Cells.Item($r,7) = [string]$row.Reason
      $ws.Cells.Item($r,8) = [string]$row.Csv
      $ws.Cells.Item($r,9) = [string]$row.UpdatedOn
      switch ($row.Status) {
        'DONE'    { $ws.Cells.Item($r,6).Interior.Color = 0x90EE90 }  # light green
        'ERROR'   { $ws.Cells.Item($r,6).Interior.Color = 0x9090FF }  # light red (BGR)
        'SKIPPED' { $ws.Cells.Item($r,6).Interior.Color = 0xE0E0E0 }  # grey
      }
      $r++
    }
    $ws.Columns.Item("A:I").AutoFit() | Out-Null
    try { $ws.Application.ActiveWindow.SplitRow = 1; $ws.Application.ActiveWindow.FreezePanes = $true } catch {}
    if (Test-Path -LiteralPath $logXlsx) { Remove-Item -LiteralPath $logXlsx -Force -ErrorAction SilentlyContinue }
    $wb.SaveAs($logXlsx, 51)   # 51 = .xlsx
    $wb.Close($true) | Out-Null
  } finally { $excel.Quit()|Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null }
}
