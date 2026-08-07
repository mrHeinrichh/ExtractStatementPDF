# SOA Extractor

Extracts **Statement of Account** CSVs from **Accounting.exe**, one per customer
`.xls`, using each customer's Invoice Frequency taken from a reference file you
export from Contact. It **resumes** — restart it any time and it only does the
statements not yet extracted.

## The workflow

1. **In Contact.exe**, click **Export** to save all customers to a CSV (e.g.
   `Customer.csv` in your Documents). That file contains every customer's **Name**
   and **StatementFrequency** — it's the frequency reference.
2. **Run SOA-Extractor.exe.** It asks for:
   - the folder of `.xls` files (your AR data),
   - **which Accounting window to use** (if more than one is open — a picker lists each
     PID, title and exe path), then
   - the **reference CSV** you exported in step 1 (the browse dialog opens in Documents).
3. It extracts each statement with the customer's correct frequency and logs everything.

### If it says "Accounting Print pane not found"
The chosen Accounting window must be **logged in** and the tool navigates it to
**General → Accounting Print** itself. The pane is located relative to *that* window,
so a second-monitor / repositioned window is fine now. If it still can't find it, the
error prints the window rectangle, the left-column labels and the custom panes it *did*
see — send me that text and I can pinpoint the difference on that PC (e.g. a different
Accounting version with different control names).

## Requirements

- **Accounting.exe open and logged in** (user hfabros). If it isn't, the tool shows
  a message telling you to log in, and stops — it never types your password.
- **Microsoft Excel** installed.
- Leave the mouse/keyboard alone while it runs (it drives the real UI).

## What it does per `.xls`

1. Reads the customer name (**E9**) and AS-OF **From/To** dates (**E18 / H18**).
2. Looks the customer up in the reference file to get its frequency (exact match,
   then a shortened/fuzzy match if the exact name isn't present).
3. Opens **Accounting Print → Statement**, binds the customer, sets the
   **Invoice Frequency** (Monthly / Daily / Weekly / Bi-Weekly), fills the dates.
4. Clicks OK and exports the report as **CSV (comma delimited)** to the output folder.
5. Records the outcome in the log.

The known frequency is tried first; if it produces nothing, it falls back to
Monthly → Daily → Weekly → Bi-Weekly.

## Reference CSV formats accepted

- **RxOffice + ARName** (preferred) — comma CSV with
  `Id,Code,RxOfficeName,StatementFrequency,ARName`. `ARName` is your AR-side alias
  (e.g. `AbalosGuillermoOptical`); the tool types the row's `RxOfficeName`.
- **RxOffice export** — comma CSV with `Id,Code,Name,StatementFrequency`.
- **Contact "Export" file** — UTF‑16, tab-delimited, with `Name`, `Alias`, `StatementFrequency`.
- **Simple `Customer,Frequency` CSV**.

The tool auto-detects the format. It matches your `.xls` customer name to a row and
types that row's RxOffice name into Accounting, in this order:
**exact** → **ARName alias** (case/space/punctuation-insensitive, so
`AbalosGuillermoOptical` = `ABALOS GUILLERMO OPTICAL` → types `ABALOS GUILLERMO`) →
**word-prefix**. If nothing matches, it types the raw `.xls` name and guesses the
frequency. The name actually typed is recorded in the log's `RxOfficeName` column.

## Output  (folder: `<your folder> - SOA CSV`)

| File | Contents |
|------|----------|
| `<customer>_<month>.csv` | One statement CSV per extracted `.xls` |
| `_SOA_Log.xlsx` | **The log** — every file with Customer, DateRange, Frequency, **Status** (DONE / ERROR / SKIPPED, colour-coded) and **Reason** (the error text for anything that failed) |
| `_SOA_Log.csv` | Same data as the log (used to resume) |

Your AR `.xls` folder is never modified. Step screenshots go to `%TEMP%\soa_shots`.

## Resume / re-run (skip already-extracted)

The tool keeps a persistent log (`_SOA_Log.csv`) and **saves it after every file**.
On the next run it **skips anything already `DONE`** (or whose CSV already exists)
and only processes the rest — so if it's interrupted, closed, or you re-open the
`.exe`, it picks up where it left off. Rows marked **ERROR** are retried on the next
run (fix the underlying cause first if needed — see the Reason column).

To force a full re-extract, delete `_SOA_Log.csv` and the CSVs (or set
`$Overwrite = $true` in `src/01-Settings.ps1` and rebuild).

## Notes & caveats

- The frequency dropdown's four options are selected by keyboard **position**, so
  values like *Bi-Weekly* actually apply (typing them left the field blank before).
- An `ERROR` reason *"No statement produced…"* means the customer didn't auto-match
  in Accounting or there were no transactions in that date range — not a frequency
  problem.
- If a name has to be shortened to match (e.g. branch names collapsing to a common
  root), it takes the first table entry that contains the shortened term. The log's
  Reason column shows when a non-exact match was used.

## Rebuild from source

Source is one function per file under `src/`. After editing, run:

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

`build.ps1` concatenates `src\*.ps1` (filename order) into
`SOA-Extractor.combined.ps1` and compiles it with **ps2exe**.

### Source layout (`src/`)

| File | Purpose |
|------|---------|
| `00-Header.ps1` | Overview, `$ErrorActionPreference`, assembly imports |
| `01-Settings.ps1` | Fallback frequency order, `$Overwrite`, screenshot dir |
| `02-Interop.ps1` | Win32 P/Invoke + UI Automation shortcut (`$AE`) |
| `03-Ui-Dialogs.ps1` | Message box + open-file (browse) dialog |
| `04-Choose-Acct.ps1` | **Chooser** — lists running Accounting windows (PID + title + path) and remembers which one to drive |
| `10-Pick-Folder.ps1` | The paste-a-path (AR folder) dialog |
| `11-Get-Acct.ps1` / `12-Focus-Acct.ps1` | Locate (the chosen) + maximize Accounting.exe |
| `13-Reset-State.ps1` | Close leftover Report/dialog before each customer |
| `14-Click.ps1` / `15-Send.ps1` / `16-Send-Literal.ps1` / `17-Shot.ps1` | Mouse / keyboard / screenshot helpers |
| `18-Assert-Login.ps1` | **Login check** — warns if Accounting isn't open/logged in |
| `20-ConvertTo-MDY.ps1` | Parse `"Mar. 02, ' 26"` → `3/2/2026` |
| `21-Read-XlsMeta.ps1` | Read customer + From/To dates from an `.xls` |
| `22-Get-Candidates.ps1` | Shorter name forms for fuzzy matching |
| `30-Get-Dialog.ps1` / `31-Open-StatementDialog.ps1` | Find / open the Statement dialog |
| `32-Field-XY.ps1` | Fixed screen offsets for each dialog field/button |
| `33-Set-Customer.ps1` / `34-Set-DateField.ps1` | Fill customer / dates |
| `35-Freq-Index.ps1` / `36-Set-Frequency.ps1` | Map + select the frequency by position |
| `37-Click-OK.ps1` / `38-Close-Dialog.ps1` | Submit / cancel the dialog |
| `40-Get-ReportWindow.ps1` / `41-Close-Report.ps1` | Find / close the Report viewer |
| `42-Export-Csv-FromReport.ps1` | Export report → CSV via the Save As dialog |
| `50-KeyOf.ps1` | Normalize a name for matching |
| `51-Load-Reference.ps1` | Load the reference (RxOffice / Contact export / simple) → name + frequency |
| `52-Explain.ps1` | Map internal status → Result + human reason |
| `53-Resolve-Reference.ps1` | Look a customer up → canonical name + frequency (exact + fuzzy) |
| `60-Log-Load.ps1` / `61-Log-Save.ps1` | Load / save the resumable run log (CSV + Excel) |
| `99-Main.ps1` | Program body: login check, browse reference, loop, resume, log |

### Optional environment overrides (for testing)

| Variable | Effect |
|----------|--------|
| `SOA_FOLDER` | Skip the folder picker; use this folder |
| `SOA_REF` | Skip the reference browse; use this CSV |
| `SOA_MAX` | Process at most N files |
| `SOA_OUT` | Write CSVs to this folder instead of the default sibling |
