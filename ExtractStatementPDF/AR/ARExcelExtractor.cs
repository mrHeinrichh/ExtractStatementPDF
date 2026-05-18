using ExcelDataReader;
using ExtractStatementPDF.Models;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.AR
{
    public class ARExcelExtractor
    {
        public ARExcelExtractor()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Aaron del Rosario");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public ARStatement Extract(IEnumerable<string> filenames)
        {
            var arStatement = new ARStatement(filenames);
            
            foreach (var filename in filenames)
            {
                var lines = GetWorksheetLines(filename);
                if (lines.Count == 0)
                {
                    continue;
                }

                arStatement.AddOrders(ParseLines(lines));
                arStatement.AddPages([string.Join(Environment.NewLine, lines)]);
            }

            return arStatement;
        }

        public ARStatement Extract(string filename)
        {
            return Extract([filename]);
        }

        private List<AROrder> ParseLines(IEnumerable<string> lines)
        {
            var arOrders = new List<AROrder>();

            foreach (var line in lines)
            {
                var arOrder = ParseLine(line);
                if (arOrder != null)
                {
                    arOrders.Add(arOrder);
                }
            }

            return arOrders;
        }

        
        private static AROrder? ParseLine(string line)
        {
            var pattern = @"(\d{2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2})\s+([A-Za-z\d]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)";
            var match = Regex.Match(line, pattern);

            if (!match.Success)
            {
                return null;
            }

            return new AROrder()
            {
                Date = match.Groups[1].Value,
                Reference = match.Groups[2].Value,
                Gross = decimal.Parse(match.Groups[3].Value, NumberStyles.Number, CultureInfo.InvariantCulture),
                Discount = decimal.Parse(match.Groups[4].Value, NumberStyles.Number, CultureInfo.InvariantCulture),
                Net = decimal.Parse(match.Groups[5].Value, NumberStyles.Number, CultureInfo.InvariantCulture),
            };
        }

        private List<string> GetWorksheetLines(string filename)
        {
            try
            {
                return Path.GetExtension(filename).ToLowerInvariant() switch
                {
                    ".xlsx" => GetOpenXmlWorksheetLines(filename),
                    ".xls" => GetLegacyWorksheetLines(filename),
                    _ => [],
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error reading excel: {filename}. {ex.Message}");
                return [];
            }
        }

        private static List<string> GetOpenXmlWorksheetLines(string filename)
        {
            using var package = new ExcelPackage(new FileInfo(filename));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();

            if (worksheet?.Dimension == null)
            {
                return [];
            }

            var lines = new List<string>();

            for (var row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                var values = Enumerable
                    .Range(1, worksheet.Dimension.End.Column)
                    .Select(column => worksheet.Cells[row, column].Text.Trim())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .ToList();

                if (values.Count > 0)
                {
                    lines.Add(string.Join(" ", values));
                }
            }

            return lines;
        }

        private static List<string> GetLegacyWorksheetLines(string filename)
        {
            using var stream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateBinaryReader(stream);

            var lines = new List<string>();

            while (reader.Read())
            {
                var values = Enumerable
                    .Range(0, reader.FieldCount)
                    .Select(i => reader.IsDBNull(i) ? null : GetCellText(reader, i))
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .ToList();

                if (values.Count > 0)
                {
                    lines.Add(string.Join(" ", values));
                }
            }

            return lines;
        }

        private static string? GetCellText(IExcelDataReader reader, int columnIndex)
        {
            var value = reader.GetValue(columnIndex);
            return value switch
            {
                DateTime dt => dt.ToString("dd-MMM-yy", CultureInfo.InvariantCulture),
                double d => d.ToString(CultureInfo.InvariantCulture),
                _ => value?.ToString()?.Trim(),
            };
        }
    }
}
