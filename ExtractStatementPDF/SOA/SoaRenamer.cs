using ExcelDataReader;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

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

        public SoaRenamer()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public List<SoaFile> Scan(string directory)
        {
            var files = Directory
                .EnumerateFiles(directory, "*.xls", SearchOption.TopDirectoryOnly)
                .Where(f => Path.GetExtension(f).Equals(".xls", StringComparison.OrdinalIgnoreCase))
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

            try
            {
                using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var reader = ExcelReaderFactory.CreateBinaryReader(stream);

                string? previousText = null;
                var row = 0;

                while (reader.Read() && row < HeaderRowLimit)
                {
                    row++;

                    var cells = Enumerable
                        .Range(0, reader.FieldCount)
                        .Select(reader.GetValue)
                        .Where(v => v != null)
                        .ToList();

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
            catch (Exception ex)
            {
                file.Error = $"cannot read file ({ex.Message})";
                return file;
            }

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

        private static void AssignNewNames(List<SoaFile> files)
        {
            var valid = files.Where(f => f.Error == null).ToList();

            foreach (var file in valid)
            {
                var suffix = GetPeriodSuffix(file.PeriodFrom!.Value, file.PeriodTo!.Value);
                file.NewName = $"{ToPascalCase(file.CustomerName!)}_{file.PeriodFrom:MMyyyy}{suffix}.xls";
            }

            // Same customer and period more than once (e.g. two full-month statements
            // with different SOA numbers) gets _P1.._Pn ordered by SOA number.
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
                    ordered[i].NewName = group.Key.Replace(".xls", $"_P{i + 1}.xls");
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
