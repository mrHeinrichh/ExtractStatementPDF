using ExcelDataReader;
using ExtractStatementPDF.Models;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

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
                    ".ods" => GetOdsWorksheetLines(filename),
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

        private static List<string> GetOdsWorksheetLines(string filename)
        {
            using var archive = ZipFile.OpenRead(filename);
            var entry = archive.GetEntry("content.xml");
            if (entry == null)
            {
                return new List<string>();
            }

            using var stream = entry.Open();
            var document = XDocument.Load(stream);
            var tableNs = XNamespace.Get("urn:oasis:names:tc:opendocument:xmlns:table:1.0");
            var officeNs = XNamespace.Get("urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            var textNs = XNamespace.Get("urn:oasis:names:tc:opendocument:xmlns:text:1.0");

            var table = document.Descendants(tableNs + "table").FirstOrDefault();
            if (table == null)
            {
                return new List<string>();
            }

            var lines = new List<string>();

            foreach (var rowElement in table.Elements(tableNs + "table-row"))
            {
                var values = new List<string>();

                foreach (var cellElement in rowElement.Elements(tableNs + "table-cell"))
                {
                    var valueType = (string?)cellElement.Attribute(officeNs + "value-type");

                    if (valueType == "date")
                    {
                        var dateValue = (string?)cellElement.Attribute(officeNs + "date-value");
                        if (dateValue != null && DateTime.TryParse(dateValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                        {
                            values.Add(date.ToString("dd-MMM-yy", CultureInfo.InvariantCulture));
                        }
                    }
                    else
                    {
                        var text = string.Concat(cellElement.Descendants(textNs + "p").Select(p => p.Value)).Trim();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            values.Add(text);
                        }
                    }
                }

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
