# Explain : turn an internal status string into @(Result, HumanReason) for the
# result report (Result is one of SUCCESS / SKIPPED / FAILED).
function Explain($status, $freq) {
  switch ($status) {
    'saved'            { return @('SUCCESS', "Extracted ($freq)") }
    'overwritten'      { return @('SUCCESS', "Extracted & replaced ($freq)") }
    'skipped-existing' { return @('SKIPPED', 'CSV already in output folder') }
    'no-data'          { return @('FAILED',  'No statement produced - customer name did not auto-match, or no transactions in this date range') }
    'no-report'        { return @('FAILED',  'Report did not open after clicking OK') }
    'no-saveas'        { return @('FAILED',  'Export started but the Save dialog did not appear') }
    'save-failed'      { return @('FAILED',  'Save dialog completed but no CSV file was written') }
    'meta-error'       { return @('FAILED',  'Could not read customer/date from the .xls') }
    default            { if ($status -like 'error*') { return @('FAILED', "Automation error: $status") } else { return @('FAILED', $status) } }
  }
}
