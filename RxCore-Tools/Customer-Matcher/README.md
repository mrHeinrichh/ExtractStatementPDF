# Customer Matcher

Writes the SOA `.xls` names that match each customer **into the Customer file
itself** — it inserts a **`MatchedXlsName`** column right after **`Name`**, so you
can see, in the Customer file, which extracted name maps to which Contact customer
alongside its `StatementFrequency`.

It's a pure data operation — **Accounting and Contact do NOT need to be open**.
Only Microsoft Excel is required (to read the `.xls` files).

## How to run

1. **Close the Customer file in Excel** (it can't be updated while Excel holds it open).
2. Double-click **Customer-Matcher.exe**.
3. Paste the folder that contains the `.xls` files (or **Browse…**), then **OK**.
4. Browse to the **Customer file you exported from Contact** (the dialog opens in
   Documents).

It updates **that same file in place**, adding/refreshing the `MatchedXlsName`
column, and prints a summary listing any `.xls` names it couldn't match.

- **In place:** the column is written into the file you select — not a separate copy.
- **Idempotent:** run it again and it *replaces* the existing `MatchedXlsName`
  column rather than adding a second one (so it always stays right after `Name`).
- **Safe write:** it writes to a temp file first and swaps it in, so a failure can't
  truncate your file. If the file is open in Excel it stops with a message asking
  you to close it. (The Customer export is re-creatable any time from Contact's
  Export button, so there's no separate backup.)

## What "matched" means

For each `.xls`, it reads the customer name (cell **E9**) and finds the Contact row
whose **Name** or **Alias** matches:

1. **Exact** (case/space-insensitive), else
2. **Word-boundary fuzzy** — progressively shorter forms of the name, taking the
   first Contact row that contains that form *as a whole word*. This is why a
   fragment like `ESCA` matches the customer **"Esca Shop"** but never the `ESCA`
   hidden inside **"Le**sca**no"**.

Multiple `.xls` names can map to one Contact row (e.g. several branches that only
exist under one Contact record); they're joined with `; `. Names with no match are
listed at the end so you can handle them manually.

## The column it adds

| Column | Meaning |
|--------|---------|
| `MatchedXlsName` (inserted right after `Name`) | The `.xls` customer name(s) that matched this Contact row; blank if none. Multiple matches joined with `; `. |

## Rebuild from source

One function per file under `src/`. After editing:

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

### Source layout (`src/`)

| File | Purpose |
|------|---------|
| `00-Header.ps1` | Overview + `$ErrorActionPreference` |
| `02-Interop.ps1` | Load WinForms/Drawing assemblies |
| `03-Ui-Dialogs.ps1` | Message box, folder paste-box, file browse |
| `10-KeyOf.ps1` | Normalize a name for matching |
| `11-WordWrap.ps1` | Word-boundary form for safe fuzzy matching |
| `12-Get-Candidates.ps1` | Progressively shorter name forms |
| `20-Read-XlsNames.ps1` | Read customer names (E9) from the `.xls` files |
| `30-Read-CustomerExport.ps1` | Read the Contact export (UTF-16 TSV, shared read) |
| `40-Match-Names.ps1` | Match each `.xls` name to a Contact row |
| `50-Write-Annotated.ps1` | Write `Customer_matched.csv` with the new column |
| `99-Main.ps1` | Program body |

### Optional environment overrides (for testing)

| Variable | Effect |
|----------|--------|
| `CM_XLS` | Skip the folder picker; use this `.xls` folder |
| `CM_REF` | Skip the browse; use this Customer file |
| `CM_OUT` | Write to this path instead of in place (used for testing) |
