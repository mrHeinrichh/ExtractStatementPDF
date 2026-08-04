# Contact Frequency Lookup

Builds a **Customer → Invoice Frequency** table (Excel) by looking each customer up
in **Contact.exe**. This is **step 1** of the workflow — its output feeds SOA-Extractor.

## What it does

For every `.xls` in the folder you choose, it:

1. Reads the customer name from cell **E9**.
2. In Contact.exe, types the name into the grid's **Name column filter** and presses
   Enter (a precise "contains" match — not the bottom Full Text Search box, which
   matches *any* word and would load the wrong customer).
3. Selects the first matching row, clicks the **Account** tab, and reads the
   **Frequency** value.
4. Writes the results to a separate output folder.

If the exact name finds nothing, it **shortens the search** (drops trailing words,
then trims letters off the first word) until rows appear, and takes the closest
(first) match — so you still get a value. See *Caveats* below.

## Requirements

- **Contact.exe** open and logged in (user `hfabros`).
- **Microsoft Excel** installed.
- Leave the mouse/keyboard alone while it runs.

## How to run

1. Double-click **Contact-Frequency-Lookup.exe**.
2. Paste the folder path that contains the `.xls` files (or click **Browse…**), then **OK**.
3. It maximizes Contact and processes each file (progress shows in the console).

## Output

Written to a sibling folder **`<your folder> - Frequency`** (your `.xls` folder is
never touched):

| File | Contents |
|------|----------|
| `CustomerFrequency_<timestamp>.xlsx` | Two columns: **Customer**, **Frequency** |
| `CustomerFrequency_<timestamp>.csv`  | Same data (UTF-8, so names like *AVENDAÑO* are preserved) |

SOA-Extractor auto-detects the newest `CustomerFrequency_*.csv` in this folder.

## Caveats

- **Ambiguous / branch names.** When a name has to be shortened to a common root,
  every branch sharing that root collapses to the same search and takes the *first*
  matching row's frequency. Example: the seven `Esca Empire Corp. -<branch>` names
  all shorten to `Esca`, so they all receive the first `Esca` row's value. If branches
  genuinely differ, spot-check those. The console prints `(via 'Esca')` whenever a
  shortened term was used, so you can see which rows were approximate.
- Grid rows aren't readable by automation, so the tool can't visually confirm the
  matched row — it relies on the match count (`0` = not found, `1` = exact,
  `>1` = ambiguous → first row used).

## Rebuild from source

The source is split one-function-per-file under `src/` (see file headers). To
recompile after editing:

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

`build.ps1` concatenates `src\*.ps1` in filename order into
`Contact-Frequency-Lookup.combined.ps1` and compiles it with **ps2exe** (installed
for the current user automatically on first run).

### Source layout (`src/`)

| File | Purpose |
|------|---------|
| `00-Header.ps1` | File overview, `$ErrorActionPreference`, assembly imports |
| `01-Settings.ps1` | Grid click offsets (Name filter box, first row) |
| `02-Interop.ps1` | Win32 P/Invoke + UI Automation shortcuts (`$AE`, `$TS`) |
| `10-Pick-Folder.ps1` | The paste-a-path dialog |
| `11-Get-Contact.ps1` | Locate the Contact.exe process |
| `12-Normalize-Window.ps1` | Move to primary monitor + maximize |
| `13-WinRect.ps1` | Read the window rectangle (anchor for offsets) |
| `14-Click.ps1` / `15-Send.ps1` / `16-Paste.ps1` | Mouse / keyboard / clipboard helpers |
| `20-Read-XlsCustomer.ps1` | Read the customer name (E9) from an `.xls` |
| `30-Get-MatchCount.ps1` | Read the "n / TOTAL" match indicator |
| `31-Click-AccountTab.ps1` | Switch to the Account tab |
| `32-Read-Frequency.ps1` | Read the Frequency value (ValuePattern) |
| `33-Clear-NameFilter.ps1` / `34-Clear-FullTextSearch.ps1` | Reset filters |
| `35-Search-NameFilter.ps1` | Apply a Name-filter term, return match count |
| `36-Get-Candidates.ps1` | Build the progressively-shorter fuzzy terms |
| `37-Lookup-Frequency.ps1` | Full per-customer lookup (ties the above together) |
| `99-Main.ps1` | Program body: loop over `.xls`, write Excel/CSV |

### Optional environment overrides (for testing)

| Variable | Effect |
|----------|--------|
| `CF_FOLDER` | Skip the picker; use this folder |
| `CF_MAX` | Process at most N files |
| `CF_OUT` | Write output to this folder instead of the default sibling |
