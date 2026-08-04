# ---- Settings (safe to tweak) ---------------------------------------------
# Fallback order used only when a customer's frequency is NOT in the lookup table.
# These strings must match the dropdown items exactly: Monthly, Daily, Weekly, Bi-Weekly.
$FrequencyOrder = @('Monthly', 'Daily', 'Weekly', 'Bi-Weekly')

# $true  = re-generate and replace CSVs that already exist in the output folder.
# $false = skip any .xls that already has a .csv in the output folder.
$Overwrite = $false

# Where per-step screenshots and the run log are written.
$ShotDir = Join-Path $env:TEMP 'soa_shots'
