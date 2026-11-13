using CsvHelper;
using MySqlConnector;
using OfficeOpenXml;
using System.Data;
using System.Globalization;
using System.Text;

namespace ExtractStatementPDF.Sql
{
    public class SqlExportOptions
    {
        // Empty = every table in the database.
        public List<string> Tables { get; init; } = [];

        public bool Structure { get; init; } = true;

        public bool Data { get; init; } = true;

        public bool DropTable { get; init; } = true;

        public bool Routines { get; init; } = true;
    }

    // Database dump/restore (MySQL / MariaDB, via MySqlBackup.NET) and saving a
    // result grid as CSV or Excel. Blocking calls: run them off the UI thread.
    public static class SqlImportExport
    {
        public static void ExportDatabase(SqlSession session, string database, string path, SqlExportOptions options,
            Action<string> progress, CancellationToken token)
        {
            using var connection = (MySqlConnection)session.CreateConnection(database);
            connection.Open();

            using var command = connection.CreateCommand();
            using var backup = new MySqlBackup(command);

            backup.ExportInfo.TablesToBeExportedList = options.Tables;
            backup.ExportInfo.ExportTableStructure = options.Structure;
            backup.ExportInfo.ExportRows = options.Data;
            backup.ExportInfo.AddDropTable = options.DropTable;
            backup.ExportInfo.ExportViews = options.Routines;
            backup.ExportInfo.ExportProcedures = options.Routines;
            backup.ExportInfo.ExportFunctions = options.Routines;
            backup.ExportInfo.ExportTriggers = options.Routines;
            backup.ExportInfo.ExportEvents = options.Routines;
            backup.ExportInfo.AddCreateDatabase = false;
            backup.ExportInfo.IntervalForProgressReport = 250;

            backup.ExportProgressChanged += (_, e) => progress(
                $"Exporting {e.CurrentTableName} (table {e.CurrentTableIndex:N0} of {e.TotalTables:N0}), " +
                $"row {e.CurrentRowIndexInAllTables:N0} of {e.TotalRowsInAllTables:N0}");

            using var stop = token.Register(backup.StopAllProcess);

            backup.ExportToFile(path);

            token.ThrowIfCancellationRequested();
            if (backup.LastError != null) throw new InvalidOperationException(backup.LastError.Message, backup.LastError);
        }

        public static void ImportMySqlFile(SqlSession session, string database, string path, Action<string> progress, CancellationToken token)
        {
            using var connection = (MySqlConnection)session.CreateConnection(database);
            connection.Open();

            using var command = connection.CreateCommand();
            using var backup = new MySqlBackup(command);

            backup.ImportInfo.IgnoreSqlError = false;   // stop at the first failing statement
            backup.ImportInfo.IntervalForProgressReport = 250;

            backup.ImportProgressChanged += (_, e) => progress(
                $"Importing {Path.GetFileName(path)}: {e.PercentageCompleted:0}% ({e.CurrentBytes / 1024.0 / 1024.0:N1} of {e.TotalBytes / 1024.0 / 1024.0:N1} MB)");

            using var stop = token.Register(backup.StopAllProcess);

            backup.ImportFromFile(path);

            token.ThrowIfCancellationRequested();

            if (backup.LastError != null)
            {
                var sql = backup.LastErrorSQL ?? "";
                if (sql.Length > 300) sql = sql[..300] + "...";
                throw new InvalidOperationException($"{backup.LastError.Message}{Environment.NewLine}Statement: {sql}", backup.LastError);
            }
        }

        public static void ExportCsv(DataTable table, string path)
        {
            // UTF-8 with BOM so Excel shows accented names (e.g. "Ñ") correctly.
            using var writer = new StreamWriter(path, false, new UTF8Encoding(true));
            using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

            foreach (DataColumn column in table.Columns) csv.WriteField(column.ColumnName);
            csv.NextRecord();

            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn column in table.Columns) csv.WriteField(FormatValue(row[column]));
                csv.NextRecord();
            }
        }

        public static void ExportExcel(DataTable table, string path, string sheetName)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Aaron del Rosario");

            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add(SafeSheetName(sheetName));

            sheet.Cells["A1"].LoadFromDataTable(table, true);
            sheet.Row(1).Style.Font.Bold = true;
            sheet.View.FreezePanes(2, 1);

            for (var i = 0; i < table.Columns.Count; i++)
            {
                if (table.Columns[i].DataType == typeof(DateTime))
                {
                    sheet.Column(i + 1).Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
                }
            }

            // Auto-fitting very large sheets is slow; the first rows give good widths.
            if (sheet.Dimension != null)
            {
                var sample = Math.Min(sheet.Dimension.End.Row, 2000);
                sheet.Cells[1, 1, sample, Math.Max(1, table.Columns.Count)].AutoFitColumns(8, 60);
            }

            package.SaveAs(new FileInfo(path));
        }

        private static string FormatValue(object value) => value switch
        {
            DBNull => "",
            DateTime date => date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString() ?? "",
        };

        private static string SafeSheetName(string name)
        {
            var cleaned = new string(name.Where(c => !"[]:*?/\\".Contains(c)).ToArray()).Trim();
            if (cleaned.Length == 0) cleaned = "Result";
            return cleaned.Length > 31 ? cleaned[..31] : cleaned;
        }
    }
}
