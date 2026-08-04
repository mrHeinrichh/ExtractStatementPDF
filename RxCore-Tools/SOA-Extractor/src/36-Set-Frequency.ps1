# Set-Frequency : choose an Invoice Frequency deterministically by keyboard.
# Open the list, go to the top (Monthly), then arrow DOWN to the target index.
# (Typing the value as text left the field blank for "Bi-Weekly" - hence index nav.)
function Set-Frequency([string]$value) {
  $idx = Freq-Index $value
  if ($idx -lt 0) { return }   # unknown -> leave the dialog default (Monthly)
  $xy = Field-XY 'InvoiceFrequency'; Click $xy[0] $xy[1]; Start-Sleep -Milliseconds 300
  Send '%{DOWN}'; Start-Sleep -Milliseconds 500          # Alt+Down opens the dropdown
  Send ('{UP}' * 6); Start-Sleep -Milliseconds 200        # force selection to the top (Monthly)
  if ($idx -gt 0) { Send ('{DOWN}' * $idx); Start-Sleep -Milliseconds 200 }
  Send '{ENTER}'; Start-Sleep -Milliseconds 300
}
