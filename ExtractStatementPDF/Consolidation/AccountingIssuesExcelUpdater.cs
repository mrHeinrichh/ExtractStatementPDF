using OfficeOpenXml;
using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.Consolidation
{
    public class AccountingIssuesExcelUpdater
    {
        private const string WorkbookFileName = "RxOffice Accounting Issues.xlsx";

        private const string GoogleSheetShortcutFileName = "RxOffice Accounting Issues.gsheet";

        private const string WorkbookDirectory = @"G:\.shortcut-targets-by-id\1VLaMf4DXM2APBuWDeKAJW_sNUkXgk7tk\Plastilens\05 - Testing\SOA";

        private const string WorksheetName = "Sheet1";

        public AccountingIssuesExcelUpdater()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Aaron del Rosario");
        }

        public void Update(IEnumerable<ConsolidatedStatement> statements)
        {
            var workbookPath = ResolveWorkbookPath();

            var rowsAdded = 0;
            var lastRow = 0;

            using (var package = new ExcelPackage(new FileInfo(workbookPath)))
            {
                var ws = package.Workbook.Worksheets[WorksheetName] ?? package.Workbook.Worksheets.First();
                var existingKeys = GetExistingKeys(ws);
                var nextRow = GetNextRow(ws);

                foreach (var statement in statements)
                {
                    foreach (var issue in statement.ConsolidatedOrders.Where(t => t.Remarks != "Good"))
                    {
                        var key = BuildKey(statement, issue);
                        if (!existingKeys.Add(key))
                        {
                            continue;
                        }

                        CopyRowFormatting(ws, nextRow);
                        WriteRow(ws, nextRow, statement, issue);

                        nextRow++;
                        rowsAdded++;
                    }
                }

                if (rowsAdded == 0)
                {
                    return;
                }

                lastRow = nextRow - 1;
                package.Save();
            }

            ExpandTrackedRange(workbookPath, lastRow);
        }

        private static void CopyRowFormatting(ExcelWorksheet ws, int row)
        {
            var templateRow = Math.Max(1, row - 1);

            for (var column = 1; column <= 13; column++)
            {
                ws.Cells[templateRow, column].Copy(ws.Cells[row, column]);
            }

            ws.Row(row).Hidden = false;
        }

        private static void WriteRow(
            ExcelWorksheet ws,
            int row,
            ConsolidatedStatement statement,
            ConsolidatedOrder issue)
        {
            ws.Cells[row, 1].Value = NormalizeName(statement.Customer);
            ws.Cells[row, 2].Value = ParseSoaText(statement.Month);
            ws.Cells[row, 3].Value = issue.Reference;
            ws.Cells[row, 4].Value = issue.Remarks;
            ws.Cells[row, 5].Value = DateTime.Today;
            ws.Cells[row, 6].Value = string.Empty;
            ws.Cells[row, 7].Value = false;
            ws.Cells[row, 8].Value = string.Empty;
            ws.Cells[row, 9].Value = string.Empty;
            ws.Cells[row, 10].Value = string.Empty;
            ws.Cells[row, 11].Value = string.Empty;
            ws.Cells[row, 12].Value = string.Empty;
            ws.Cells[row, 13].Value = string.Empty;
        }

        private static string BuildKey(ConsolidatedStatement statement, ConsolidatedOrder issue)
        {
            return $"{NormalizeName(statement.Customer)}|{issue.Reference}|{issue.Remarks}";
        }

        private static HashSet<string> GetExistingKeys(ExcelWorksheet ws)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (ws.Dimension == null)
            {
                return keys;
            }

            for (var row = 2; row <= ws.Dimension.End.Row; row++)
            {
                var customer = ws.Cells[row, 1].Text.Trim();
                var reference = ws.Cells[row, 3].Text.Trim();
                var issueType = ws.Cells[row, 4].Text.Trim();

                if (customer.Length == 0 || reference.Length == 0 || issueType.Length == 0)
                {
                    continue;
                }

                keys.Add($"{customer}|{reference}|{issueType}");
            }

            return keys;
        }

        private static int GetNextRow(ExcelWorksheet ws)
        {
            if (ws.Dimension == null)
            {
                return 2;
            }

            return ws.Dimension.End.Row + 1;
        }

        private static string NormalizeName(string input)
        {
            return Regex.Replace(input, "([a-z0-9])([A-Z])", "$1 $2").Trim().ToUpper();
        }

        private static string ParseSoaText(string input)
        {
            if (DateTime.TryParseExact(input, "MMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            {
                return date.ToString("MMMM yyyy", CultureInfo.InvariantCulture);
            }

            return input;
        }

        private static string ResolveWorkbookPath()
        {
            var workbookPath = Path.Combine(WorkbookDirectory, WorkbookFileName);
            if (File.Exists(workbookPath))
            {
                return workbookPath;
            }

            var googleSheetShortcutPath = Path.Combine(WorkbookDirectory, GoogleSheetShortcutFileName);
            if (File.Exists(googleSheetShortcutPath))
            {
                throw new InvalidOperationException(
                    $"Target '{googleSheetShortcutPath}' is a Google Sheets shortcut (.gsheet), not an Excel workbook. " +
                    $"This app can only update '{WorkbookFileName}' with EPPlus. Use a real .xlsx file in that folder or add Google Sheets API support.");
            }

            throw new FileNotFoundException(
                $"Could not find '{WorkbookFileName}' in '{WorkbookDirectory}'.",
                workbookPath);
        }

        private static void ExpandTrackedRange(string workbookPath, int lastRow)
        {
            UpdateZipEntry(
                workbookPath,
                "xl/workbook.xml",
                text => Regex.Replace(
                    text,
                    @"(Sheet1!\$A\$1:\$M\$)\d+",
                    m => $"{m.Groups[1].Value}{lastRow}"));

            UpdateZipEntry(
                workbookPath,
                "xl/worksheets/sheet1.xml",
                text => Regex.Replace(
                    text,
                    "(<autoFilter ref=\"\\$A\\$1:\\$M\\$)\\d+(\"/>)",
                    m => $"{m.Groups[1].Value}{lastRow}{m.Groups[2].Value}"));

            UpdateZipEntry(
                workbookPath,
                "xl/pivotCache/pivotCacheDefinition1.xml",
                text => Regex.Replace(
                    text,
                    "(<worksheetSource ref=\"A1:M)\\d+(\" sheet=\"Sheet1\"/>)",
                    m => $"{m.Groups[1].Value}{lastRow}{m.Groups[2].Value}"));

            UpdateZipEntry(
                workbookPath,
                "xl/worksheets/sheet1.xml",
                text => Regex.Replace(
                    text,
                    "(<dataValidation type=\"list\" allowBlank=\"1\" showErrorMessage=\"1\" sqref=\"I2:I)\\d+(\">)",
                    m => $"{m.Groups[1].Value}{lastRow}{m.Groups[2].Value}"));
        }

        private static void UpdateZipEntry(
            string workbookPath,
            string entryName,
            Func<string, string> update)
        {
            using var archive = ZipFile.Open(workbookPath, ZipArchiveMode.Update);
            var entry = archive.GetEntry(entryName);
            if (entry == null)
            {
                return;
            }

            string updatedText;
            using (var reader = new StreamReader(entry.Open()))
            {
                updatedText = update(reader.ReadToEnd());
            }

            entry.Delete();
            var newEntry = archive.CreateEntry(entryName);
            using var writer = new StreamWriter(newEntry.Open());
            writer.Write(updatedText);
        }
    }
}
