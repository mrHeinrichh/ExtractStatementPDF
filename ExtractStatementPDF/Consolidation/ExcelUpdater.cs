using OfficeOpenXml;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.Consolidation
{
    public class ExcelUpdater
    {
        public ExcelUpdater()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Aaron del Rosario");
        }

        public void Update(List<ConsolidatedStatement> statements)
        {
            var fileInfo = new FileInfo("G:\\.shortcut-targets-by-id\\1VLaMf4DXM2APBuWDeKAJW_sNUkXgk7tk\\Plastilens\\05 - Testing\\SOA\\SOA Results.xlsx");
            var package = new ExcelPackage(fileInfo);

            var ws = package.Workbook.Worksheets.First();
            var nextRow = GetNextRow(ws);

            foreach (var statement in statements)
            {
                ws.InsertRow(nextRow, 1, 2);

                ws.Cells[nextRow, 1].Value = NormalizeName(statement.Customer);
                ws.Cells[nextRow - 1, 2].Copy(ws.Cells[nextRow, 2]);
                ws.Cells[nextRow - 1, 3].Copy(ws.Cells[nextRow, 3]);
                ws.Cells[nextRow, 4].Value = ParseDate(statement.Month);
                ws.Cells[nextRow, 5].Value = statement.Period;
                ws.Cells[nextRow, 6].Value = statement.ARStatement.TotalNet();
                ws.Cells[nextRow, 7].Value = statement.RxOfficeStatement.TotalNet();
                ws.Cells[nextRow, 8].Value = statement.NetVariance();
                ws.Cells[nextRow, 9].Value = statement.CumulativeVariance();

                nextRow++;
            }

            package.Save();
        }

        private string NormalizeName(string input)
        {
            var text = Regex.Replace(input, "([a-z0-9])([A-Z])", "$1 $2");

            return text;
        }

        private DateTime ParseDate(string input)
        {
            if (DateTime.TryParseExact(input, "MMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            {
                return date;
            }

            return DateTime.Today;
        }

        private int GetNextRow(ExcelWorksheet ws)
        {
            int nextRow = 1;

            if (ws.Dimension != null)
                nextRow = ws.Dimension.End.Row + 1;

            return nextRow;
        }
    }
}
