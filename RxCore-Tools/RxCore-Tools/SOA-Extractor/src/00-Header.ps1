<#
  SOA Extractor  -  assembled source
  ==================================
  This tool is written as one function per file under this src\ folder.
  build.ps1 concatenates every src\*.ps1 (in filename order) into a single
  script and compiles it to SOA-Extractor.exe with ps2exe.

  Load order is controlled by the numeric filename prefix:
     00  header + assemblies         (this file)
     01  settings
     02  Win32 interop + shortcuts
     10-52  one function per file
     99  main program body (runs last)

  What it does: for every .xls in a chosen folder, reads the customer + dates,
  drives Accounting.exe's "Accounting Print > Statement" screen, sets the
  customer's known Invoice Frequency, generates the report and exports it to a
  .csv in a SEPARATE "<folder> - SOA CSV" output folder.

  Requires: Accounting.exe already OPEN and LOGGED IN (user hfabros); Excel installed.
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
