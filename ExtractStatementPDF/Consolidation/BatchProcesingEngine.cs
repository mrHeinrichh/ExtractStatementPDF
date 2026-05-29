using ExtractStatementPDF.AR;
using ExtractStatementPDF.RxOffice;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.Consolidation
{
    public class BatchProcesingEngine
    {
        private ARExtractor _arExtractor = new ARExtractor();

        private RxOfficeExtractor _rxOfficeExtractor = new RxOfficeExtractor();

        private ConsolidatedStatementExcel _excelGenerator = new ConsolidatedStatementExcel();

        public BatchProcesingEngine() { }
       
        public void Verify(string directory)
        {
            var directoryInfo = new DirectoryInfo(directory);
            var arCopies = new List<FileInfo>();
            var csvs = new List<FileInfo>();

            foreach (var file in directoryInfo.GetFiles($"{Subdirectories.Verification}/*", SearchOption.AllDirectories))
            {
                switch (file.Extension.ToLowerInvariant())
                {
                    case ".pdf":
                    case ".xls":
                        arCopies.Add(file);
                        break;
                }
            }

            var archiveDirectory = EnsureDirectoryExists(directoryInfo.FullName, Subdirectories.Archived);
            var verifiedDirectory = EnsureDirectoryExists(directoryInfo.FullName, Subdirectories.Verified);

            var invalidStatements = new List<ARStatement>();

            foreach (var file in arCopies)
            {
              
                var statement = _arExtractor.Extract(file.FullName);

                if (!statement.IsValid())
                {
                    invalidStatements.Add(statement);

                    MoveFilesTo(null, [file], archiveDirectory);
                }
                else
                {
                    MoveFilesTo(null, [file], verifiedDirectory);

                }
            }
        }

        public void Process(string directory)
        {
            var directoryInfo = new DirectoryInfo(directory);

            var archivedDirectory = EnsureDirectoryExists(directoryInfo.FullName, Subdirectories.Archived);
            var processedDirectory = EnsureDirectoryExists(directoryInfo.FullName, Subdirectories.Processed);

            var arCopies = new List<FileInfo>();
            var csvs = new List<FileInfo>();

            foreach (var file in directoryInfo.GetFiles($"{Subdirectories.Matched}/*", SearchOption.AllDirectories))
            {
                switch (file.Extension.ToLowerInvariant())
                {
                    case ".pdf":
                    case ".xls":
                        arCopies.Add(file);
                        break;
                    case ".csv":
                        csvs.Add(file);
                        break;
                }
            }

            var matches = MatchFiles(arCopies, csvs);
            var statements = new List<ConsolidatedStatement>();

            foreach (var match in matches)
            {
                var statement = Reconciliate(match.Key, match.Value);

                if (statement.IsValid())
                {
                    statements.Add(statement);
                    MoveFilesTo(match.Key, match.Value, processedDirectory);
                }
                else
                {
                    MoveFilesTo(match.Key, match.Value, archivedDirectory);
                }
            }

            foreach (var statement in statements)
            {
                var bytes = _excelGenerator.GenerateExcel(statement);
                var filename = Path.GetFileNameWithoutExtension(statement.Filename) + ".xlsx";
                using var filestream = new FileStream($"{directory}/{filename}", FileMode.CreateNew, FileAccess.Write);
                filestream.Write(bytes);
            }

            Update(statements);
        }

        private static string EnsureDirectoryExists(string rootDirectory, string subdirectory)
        {
            var directory = Path.Combine(rootDirectory, subdirectory);
            Directory.CreateDirectory(directory);

            return directory;
        }

        private static void MoveFilesTo(FileInfo? csv, IEnumerable<FileInfo> arFiles, string directory)
        {
            if (csv is not null)
                MoveToDirectory(csv, directory);

            foreach (var file in arFiles)
            {
                MoveToDirectory(file, directory);
            }
        }

        private static void MoveToDirectory(FileInfo file, string directory)
        {
            var dest = Path.Combine(directory, file.Name);
            file.MoveTo(dest, overwrite: true);
        }

        private static void Update(List<ConsolidatedStatement> s)
        {
            var excel = new ExcelUpdater();
            var accountingIssues = new AccountingIssuesExcelUpdater();

            excel.Update(s);
            accountingIssues.Update(s);
        }

        private static Dictionary<FileInfo, IEnumerable<FileInfo>> MatchFiles(List<FileInfo> arCopies, List<FileInfo> csvs)
        {
            var matches = new Dictionary<FileInfo, IEnumerable<FileInfo>>();
            foreach (var currentCsv in csvs)
            {
                var csvKey = BuildLookupKey(currentCsv.FullName);
                var matchingPdfs = arCopies
                    .Where(t => BuildLookupKey(t.FullName) == csvKey)
                    .OrderByDescending(t => string.Equals(t.Extension, ".pdf", StringComparison.OrdinalIgnoreCase))
                    .ThenBy(t => t.Name);

                matches.Add(currentCsv, matchingPdfs);
            }

            return matches;
        }

        private static string BuildLookupKey(string fullName)
        {
            var name = Path.GetFileNameWithoutExtension(fullName);
            var regex = new Regex(@"^([\w&]+?)(?=_P\d+|$)");
            var matches = regex.Match(name);

            if (matches.Success) return matches.Groups[1].Value;
            return "";
        }

        private ConsolidatedStatement Reconciliate(FileInfo csv, IEnumerable<FileInfo> pdfs)
        {
            var arStatement = _arExtractor.Extract(pdfs.Select(t => t.FullName));
            var rxOfficeStatement = _rxOfficeExtractor.Extract(csv.FullName);

            var statement = new ConsolidatedStatement(arStatement, rxOfficeStatement);

            return statement;
        }
    }
}
