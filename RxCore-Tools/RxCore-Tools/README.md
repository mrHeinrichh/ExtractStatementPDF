# RxCore Tools

Automation for extracting Statement-of-Account CSVs from the Plastilens RxCore
suite, with each customer billed at its correct **Invoice Frequency**.

```
RxCore-Tools/
├─ README.md                      ← you are here
├─ SOA-Extractor/                 ← the main tool
│  ├─ SOA-Extractor.exe
│  ├─ README.md
│  ├─ build.ps1
│  └─ src/                        (one .ps1 per function)
└─ Contact-Frequency-Lookup/      ← optional / legacy (see note below)
   ├─ Contact-Frequency-Lookup.exe
   ├─ README.md
   ├─ build.ps1
   └─ src/
```

## The workflow (2 steps)

**Step 1 — export customers from Contact (manual, once per batch).**
In **Contact.exe**, click **Export** and save the file (e.g. `Customer.csv`) to your
**Documents**. This file lists every customer's **Name** and **StatementFrequency**
— it's the frequency reference. (Nothing to install or script; it's one click.)

**Step 2 — run SOA-Extractor.exe.**
It asks for the folder of customer `.xls` files, then for the `Customer.csv` you
exported (the browse dialog opens in Documents). It then extracts each Statement of
Account using that customer's frequency, into `<folder> - SOA CSV\`, and writes a
log you can re-run against.

## Before you run

- The relevant app must be **open and logged in** (user `hfabros`):
  - **Contact.exe** for step 1 (the Export),
  - **Accounting.exe** for step 2 (the extraction).
- If Accounting isn't open/logged in when you start SOA-Extractor, it shows a
  message asking you to log in and stops. **It never types your password** — you
  log in yourself.
- **Microsoft Excel** must be installed.
- Don't touch the mouse/keyboard while a tool runs.
- SmartScreen may warn on first launch (unsigned exe): *More info → Run anyway*.

## Resume

SOA-Extractor keeps a log (`_SOA_Log.xlsx` / `_SOA_Log.csv`) in the output folder and
saves it after every file. Re-running skips everything already **DONE** and only
does the rest, so you can stop/restart freely. Failed rows are marked **ERROR** with
the reason and are retried next run.

## About Contact-Frequency-Lookup (legacy)

This second tool looked customers up one-by-one inside Contact. It's now
**superseded** by the one-click **Export** in step 1, which is faster and gives the
whole list at once. It's kept for reference; you don't need it for the workflow
above. (SOA-Extractor reads the Contact Export directly.)

## Editing / rebuilding

The real source is the split files in each `src/` folder (one function per file,
each with a comment header). After editing, run that tool's `build.ps1` to
reassemble and recompile the `.exe`:

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

See each tool's own `README.md` for full detail.
