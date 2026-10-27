# ---- Settings (safe to tweak) ---------------------------------------------
# Fallback order used only when a customer's frequency is NOT in the lookup table.
# These strings must match the dropdown items exactly: Monthly, Daily, Weekly, Bi-Weekly.
$FrequencyOrder = @('Monthly', 'Daily', 'Weekly', 'Bi-Weekly')

# $true  = re-generate and replace CSVs that already exist in the output folder.
# $false = skip any .xls that already has a .csv in the output folder.
$Overwrite = $false

# Where per-step screenshots and the run log are written.
$ShotDir = Join-Path $env:TEMP 'soa_shots'

# Customers (ARName from the Customer Master Record) whose AR data is split into
# <Name>_<MMyyyy>_P1 (whole month), _P2 (1-15) and _P3 (16-end of month).
# RxOffice is extracted ONCE for the whole month as <Name>_<MMyyyy>.csv.
$WholeMonthCustomers = @('CarlosSimpao', 'SimpaoTolentino')

# When the matched customer produces no statement, try up to this many nearby RxOffice
# names (e.g. "GOCO OPTICAL" -> "GOCO OPTICAL INDIGENT"). 0 turns the fallback off.
$MaxNearbyNames = 3
