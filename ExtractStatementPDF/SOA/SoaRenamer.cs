using ExcelDataReader;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ExtractStatementPDF.SOA
{
    public class SoaFile
    {
        public required string FullPath { get; init; }

        public string FileName => Path.GetFileName(FullPath);

        public string? CustomerName { get; set; }

        public string? DocumentNo { get; set; }

        public DateTime? PeriodFrom { get; set; }

        public DateTime? PeriodTo { get; set; }

        public string? NewName { get; set; }

        public string? Error { get; set; }

        public bool CanRename => Error == null
            && NewName != null
            && !string.Equals(NewName, FileName, StringComparison.Ordinal);
    }

    public class SoaRenamer
    {
        private const int HeaderRowLimit = 30;

        private static readonly string[] SupportedExtensions = [".xls", ".ods"];

        private static readonly Dictionary<string, string> WordOverrides = new(StringComparer.OrdinalIgnoreCase)
        {
            ["DR"] = "Dr",
            ["DRA"] = "Dra",
            ["DE"] = "De",
            ["OF"] = "Of",
        };

        // Initialisms that stay fully uppercase in file names.
        private static readonly HashSet<string> UppercaseWords = new(StringComparer.OrdinalIgnoreCase)
        {
            "SM", "JP", "LC", "BF", "GMA", "MTC", "DLSV", "OOS", "APECS",
        };

        private static readonly XNamespace TableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        private static readonly XNamespace OfficeNs = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        private static readonly XNamespace TextNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

        public SoaRenamer()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public List<SoaFile> Scan(string directory)
        {
            var files = Directory
                .EnumerateFiles(directory, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => SupportedExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .Select(ReadHeader)
                .ToList();

            AssignNewNames(files);

            return files;
        }

        public List<string> Apply(IEnumerable<SoaFile> files)
        {
            var log = new List<string>();

            foreach (var file in files.Where(f => f.CanRename))
            {
                var target = Path.Combine(Path.GetDirectoryName(file.FullPath)!, file.NewName!);

                try
                {
                    if (File.Exists(target))
                    {
                        log.Add($"SKIPPED {file.FileName} -> {file.NewName} (target already exists)");
                        continue;
                    }

                    File.Move(file.FullPath, target);
                    log.Add($"RENAMED {file.FileName} -> {file.NewName}");
                }
                catch (Exception ex)
                {
                    log.Add($"FAILED {file.FileName} -> {file.NewName} ({ex.Message})");
                }
            }

            return log;
        }

        private static SoaFile ReadHeader(string path)
        {
            var file = new SoaFile { FullPath = path };

            List<List<object>> rows;

            try
            {
                rows = Path.GetExtension(path).ToLowerInvariant() switch
                {
                    ".ods" => ReadOdsRows(path),
                    _ => ReadXlsRows(path),
                };
            }
            catch (Exception ex)
            {
                file.Error = $"cannot read file ({ex.Message})";
                return file;
            }

            ParseHeader(file, rows);

            if (string.IsNullOrWhiteSpace(file.CustomerName))
            {
                file.Error = "CUSTOMER NAME not found";
            }
            else if (file.PeriodFrom == null || file.PeriodTo == null)
            {
                file.Error = "statement period (AS OF) not found";
            }

            return file;
        }

        // Reads each header row into a list of cell values (string or DateTime),
        // independent of the underlying file format.
        private static void ParseHeader(SoaFile file, List<List<object>> rows)
        {
            string? previousText = null;

            foreach (var cells in rows.Take(HeaderRowLimit))
            {
                var texts = cells
                    .OfType<string>()
                    .Select(t => t.Trim())
                    .Where(t => t.Length > 0)
                    .ToList();

                if (texts.Count == 0 && cells.Count == 0)
                {
                    continue;
                }

                if (file.CustomerName == null)
                {
                    if (texts.Any(t => t.StartsWith("CUSTOMER NAME", StringComparison.OrdinalIgnoreCase)))
                    {
                        // The customer name is printed in the row above the label
                        // (merged cells), but fall back to a value on the label row itself.
                        file.CustomerName = texts.FirstOrDefault(t => !t.Contains(':') && !IsDocumentNo(t)) ?? previousText;
                        file.DocumentNo = texts.LastOrDefault(IsDocumentNo);
                    }
                    else
                    {
                        previousText = texts.LastOrDefault(t => !IsLetterheadText(t)) ?? previousText;
                    }
                }

                if (texts.Any(t => t.StartsWith("AS OF", StringComparison.OrdinalIgnoreCase)))
                {
                    var dates = cells.OfType<DateTime>().ToList();

                    if (dates.Count == 0)
                    {
                        dates = texts
                            .Where(t => DateTime.TryParse(t, out _))
                            .Select(DateTime.Parse)
                            .ToList();
                    }

                    if (dates.Count > 0)
                    {
                        file.PeriodFrom = dates.First();
                        file.PeriodTo = dates.Last();
                    }
                }

                if (file.CustomerName != null && file.PeriodFrom != null)
                {
                    break;
                }
            }
        }

        private static List<List<object>> ReadXlsRows(string path)
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateBinaryReader(stream);

            var rows = new List<List<object>>();

            while (reader.Read() && rows.Count < HeaderRowLimit)
            {
                var cells = Enumerable
                    .Range(0, reader.FieldCount)
                    .Select(reader.GetValue)
                    .Where(v => v != null)
                    .Select(v => v!)
                    .ToList();

                rows.Add(cells);
            }

            return rows;
        }

        // An .ods file is a ZIP whose content.xml holds the spreadsheet. Date cells
        // carry the real value in office:date-value (the visible text is unparseable),
        // so those are read as DateTime; everything else is read as trimmed text.
        private static List<List<object>> ReadOdsRows(string path)
        {
            using var archive = ZipFile.OpenRead(path);
            var entry = archive.GetEntry("content.xml")
                ?? throw new InvalidDataException("content.xml not found in .ods");

            using var stream = entry.Open();
            var document = XDocument.Load(stream);

            var table = document.Descendants(TableNs + "table").FirstOrDefault()
                ?? throw new InvalidDataException("no table found in .ods");

            var rows = new List<List<object>>();

            foreach (var rowElement in table.Elements(TableNs + "table-row"))
            {
                if (rows.Count >= HeaderRowLimit)
                {
                    break;
                }

                var cells = new List<object>();

                foreach (var cellElement in rowElement.Elements(TableNs + "table-cell"))
                {
                    var valueType = (string?)cellElement.Attribute(OfficeNs + "value-type");

                    if (valueType == "date")
                    {
                        var dateValue = (string?)cellElement.Attribute(OfficeNs + "date-value");
                        if (DateTime.TryParse(dateValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                        {
                            cells.Add(date);
                        }
                    }
                    else
                    {
                        var text = string.Concat(cellElement.Descendants(TextNs + "p").Select(p => p.Value)).Trim();
                        if (text.Length > 0)
                        {
                            cells.Add(text);
                        }
                    }
                }

                rows.Add(cells);
            }

            return rows;
        }

        private static void AssignNewNames(List<SoaFile> files)
        {
            var valid = files.Where(f => f.Error == null).ToList();

            foreach (var file in valid)
            {
                var extension = Path.GetExtension(file.FullPath).ToLowerInvariant();
                var suffix = GetPeriodSuffix(file.PeriodFrom!.Value, file.PeriodTo!.Value);
                file.NewName = $"{ToPascalCase(file.CustomerName!)}_{file.PeriodFrom:MMyyyy}{suffix}{extension}";
            }

            // Same customer, period and format more than once (e.g. two full-month
            // statements with different SOA numbers) gets _P1.._Pn ordered by SOA number.
            foreach (var group in valid.GroupBy(f => f.NewName!, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
            {
                var ordered = group
                    .OrderBy(f => f.DocumentNo ?? f.FileName, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (group.Key.Contains("_1st") || group.Key.Contains("_2nd"))
                {
                    foreach (var file in ordered.Skip(1))
                    {
                        file.Error = $"duplicate target name {group.Key}";
                        file.NewName = null;
                    }

                    continue;
                }

                for (var i = 0; i < ordered.Count; i++)
                {
                    var extension = Path.GetExtension(ordered[i].NewName!);
                    var baseName = Path.GetFileNameWithoutExtension(ordered[i].NewName!);
                    ordered[i].NewName = $"{baseName}_P{i + 1}{extension}";
                }
            }
        }

        private static string GetPeriodSuffix(DateTime from, DateTime to)
        {
            if (to.Day <= 15) return "_1st";
            if (from.Day >= 16) return "_2nd";
            return "";
        }

        private static string ToPascalCase(string name)
        {
            var words = Regex.Split(RemoveDiacritics(name).Replace("&", " & "), @"[\s\-/.]+");
            var result = new StringBuilder();

            foreach (var word in words)
            {
                if (word == "&")
                {
                    result.Append('&');
                    continue;
                }

                var cleaned = Regex.Replace(word, "[^A-Za-z0-9]", "");
                if (cleaned.Length == 0)
                {
                    continue;
                }

                if (WordOverrides.TryGetValue(cleaned, out var replacement))
                {
                    result.Append(replacement);
                }
                else if (cleaned.Length == 1 || UppercaseWords.Contains(cleaned))
                {
                    result.Append(cleaned.ToUpperInvariant());
                }
                else
                {
                    result.Append(char.ToUpperInvariant(cleaned[0]));
                    result.Append(cleaned[1..].ToLowerInvariant());
                }
            }

            return result.ToString();
        }

        private static string RemoveDiacritics(string text)
        {
            var normalized = text.Normalize(NormalizationForm.FormD);
            var result = new StringBuilder();

            foreach (var c in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    result.Append(c);
                }
            }

            return result.ToString().Normalize(NormalizationForm.FormC);
        }

        private static bool IsDocumentNo(string text)
        {
            return Regex.IsMatch(text, @"^\d{3,}[A-Z]{0,2}$");
        }

        private static bool IsLetterheadText(string text)
        {
            return text.StartsWith("PLASTILENS", StringComparison.OrdinalIgnoreCase)
                || text.StartsWith("#")
                || text.StartsWith("Tel No", StringComparison.OrdinalIgnoreCase)
                || text.StartsWith("Fax No", StringComparison.OrdinalIgnoreCase)
                || text.StartsWith("STATEMENT", StringComparison.OrdinalIgnoreCase);
        }
    }
}
