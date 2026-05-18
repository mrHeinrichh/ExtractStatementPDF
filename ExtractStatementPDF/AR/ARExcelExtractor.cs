using ExtractStatementPDF.Models;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.AR
{
    public class ARExcelExtractor
    {
        public ARExcelExtractor()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Aaron del Rosario");
        }

        public ARStatement Extract(IEnumerable<string> filenames)
        {
            var arStatement = new ARStatement();

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

        public List<string> ExtractReferences(IEnumerable<string> filenames)
        {
            var references = new List<string>();

            foreach (var filename in filenames)
            {
                var lines = GetWorksheetLines(filename);
                if (lines.Count == 0)
                {
                    continue;
                }

                references.AddRange(ParseReferences(lines));
            }

            return references;
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

        private static IEnumerable<string> ParseReferences(IEnumerable<string> lines)
        {
            foreach (var line in lines)
            {
                var reference = ParseReference(line);
                if (!string.IsNullOrWhiteSpace(reference))
                {
                    yield return reference;
                }
            }
        }

        private static AROrder? ParseLine(string line)
        {
            var pattern = @"(\d{2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2})\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)";
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

        private static string? ParseReference(string line)
        {
            var pattern = @"(\d{2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2})\s+([A-Za-z0-9-]+)\s+";
            var match = Regex.Match(line, pattern);

            if (!match.Success)
            {
                return null;
            }

            return match.Groups[2].Value;
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
            var commandText = BuildLegacyExcelReadCommand(filename);
            var encodedCommand = Convert.ToBase64String(Encoding.Unicode.GetBytes(commandText));
            var powershellPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "WindowsPowerShell",
                "v1.0",
                "powershell.exe");

            using var process = new Process();
            process.StartInfo.FileName = powershellPath;
            process.StartInfo.Arguments = $"-NoProfile -NonInteractive -EncodedCommand {encodedCommand}";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
            process.StartInfo.StandardErrorEncoding = Encoding.UTF8;

            process.Start();

            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();

            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                Debug.WriteLine($"Error reading legacy excel: {filename}. {error}");
                return [];
            }

            var lines = JsonSerializer.Deserialize<List<string>>(output.Trim());
            return lines ?? [];
        }

        private static string BuildLegacyExcelReadCommand(string filename)
        {
            var escapedFilename = filename.Replace("'", "''");

            return $$"""
$path = '{{escapedFilename}}'
Add-Type -AssemblyName System.Data
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$path;Extended Properties='Excel 8.0;HDR=NO;IMEX=1'"
$connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
try {
    $connection.Open()
    $schema = $connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    $sheet = @($schema.Rows | ForEach-Object { $_["TABLE_NAME"].ToString() } | Where-Object { $_ -match '\$' } | Select-Object -First 1)
    if ($sheet.Count -eq 0) {
        @() | ConvertTo-Json -Compress
        exit 0
    }

    $command = $connection.CreateCommand()
    $command.CommandText = "SELECT * FROM [$($sheet[0])]"
    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
    $table = New-Object System.Data.DataTable
    [void]$adapter.Fill($table)

    $lines = foreach ($row in $table.Rows) {
        $values = foreach ($value in $row.ItemArray) {
            if ($null -ne $value) {
                $text = $value.ToString().Trim()
                if ($text.Length -gt 0) {
                    $text
                }
            }
        }

        if ($values) {
            $values -join ' '
        }
    }

    @($lines) | ConvertTo-Json -Compress
}
finally {
    if ($connection.State -eq 'Open') {
        $connection.Close()
    }
}
""";
        }
    }
}
