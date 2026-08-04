<#
  Contact Frequency Lookup  -  assembled source
  =============================================
  One function per file under this src\ folder. build.ps1 concatenates every
  src\*.ps1 (in filename order) into one script and compiles it to
  Contact-Frequency-Lookup.exe with ps2exe.

  Load order (by numeric filename prefix):
     00 header + assemblies (this file)   01 settings (grid offsets)
     02 Win32 interop                     10-37 one function per file
     99 main program body (runs last)

  What it does: reads the customer name from every .xls in a chosen folder,
  drives Contact.exe to look each one up (Name column filter -> select row ->
  Account tab -> read Frequency), and writes a Name+Frequency table to Excel in
  a SEPARATE "<folder> - Frequency" output folder. If an exact name finds nothing
  it shortens the search until a match appears (closest match wins).

  Requires: Contact.exe already OPEN and LOGGED IN (user hfabros); Excel installed.
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
