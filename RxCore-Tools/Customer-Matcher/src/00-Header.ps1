<#
  Customer Matcher  -  assembled source
  =====================================
  One function per file under src\. build.ps1 concatenates them (filename order)
  and compiles Customer-Matcher.exe with ps2exe.

  What it does (pure data - no Accounting/Contact needed):
    * Reads the customer name (cell E9) from every .xls in a folder you choose.
    * Reads the Customer CSV you exported from Contact (has Name, Alias,
      StatementFrequency).
    * Matches each .xls name to a Contact row (exact, then word-boundary fuzzy).
    * Writes "Customer_matched.csv" next to the export, identical to it but with a
      new "MatchedXlsName" column inserted right after "Name", filled with the
      .xls customer name(s) that matched that row.

  Requires: Microsoft Excel installed (to read the .xls files). Contact/Accounting
  do NOT need to be open.
#>

$ErrorActionPreference = 'Stop'
