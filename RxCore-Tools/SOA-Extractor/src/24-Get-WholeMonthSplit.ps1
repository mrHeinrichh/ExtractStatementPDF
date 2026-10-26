# Get-WholeMonthSplit : for a whole-month customer's AR file name
# (e.g. "CarlosSimpao_052026_P2"), return @{ Csv='CarlosSimpao_052026'; Part=2;
# From='5/1/2026'; To='5/31/2026' } - the single whole-month CSV name and the
# calendar month range. Returns $null for every other file.
function Get-WholeMonthSplit([string]$base) {
  if ($base -notmatch '^(?<name>[^_]+)_(?<mm>\d{2})(?<yyyy>\d{4})_P(?<part>[123])$') { return $null }
  if ($WholeMonthCustomers -notcontains $Matches.name) { return $null }
  $m = [int]$Matches.mm; $y = [int]$Matches.yyyy
  return @{
    Csv  = "{0}_{1}{2}" -f $Matches.name, $Matches.mm, $Matches.yyyy
    Part = [int]$Matches.part
    From = "$m/1/$y"
    To   = "$m/$([DateTime]::DaysInMonth($y,$m))/$y"
  }
}
