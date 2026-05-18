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
       
        public void Process(string directory)
        {
            var directoryInfo = new DirectoryInfo(directory);
            var archiveDirectory = Path.Combine(directoryInfo.FullName, "Archive");

            var arCopies = new List<FileInfo>();
            var csvs = new List<FileInfo>();

            foreach (var file in directoryInfo.GetFiles("*", SearchOption.AllDirectories))
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

            Directory.CreateDirectory(archiveDirectory);

            foreach (var match in matches)
            {
                var statement = Reconciliate(match.Key, match.Value);

                if (statement.IsValid())
                {
                    statements.Add(statement);
                }
                else
                {
                    ArchiveFiles(match.Key, match.Value, archiveDirectory);
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

        private static void ArchiveFiles(FileInfo csv, IEnumerable<FileInfo> arFiles, string archiveDirectory)
        {
            MoveToArchive(csv, archiveDirectory);
            foreach (var file in arFiles)
            {
                MoveToArchive(file, archiveDirectory);
            }
        }

        private static void MoveToArchive(FileInfo file, string archiveDirectory)
        {
            var dest = Path.Combine(archiveDirectory, file.Name);
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
