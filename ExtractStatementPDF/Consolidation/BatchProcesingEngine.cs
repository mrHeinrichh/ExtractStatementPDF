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

            var arCopies = new List<FileInfo>();
            var csvs = new List<FileInfo>();

            foreach (var file in directoryInfo.GetFiles("*", SearchOption.AllDirectories))
            {
                switch (file.Extension.ToLowerInvariant())
                {
                    case ".pdf":
                    case ".xls":
                    case ".xlsx":
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
                statements.Add(statement);

                var bytes = _excelGenerator.GenerateExcel(statement);

                var filename = Path.GetFileNameWithoutExtension(statement.Filename) + ".xlsx";
                using var filestream = new FileStream($"{directory}/{filename}", FileMode.CreateNew, FileAccess.Write);
                filestream.Write(bytes);
            }

            Update(statements);
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
            var regex = new Regex(@"^([\w]+?)(?=_P\d+|$)");
            var matches = regex.Match(name);

            if (matches.Success) return matches.Groups[1].Value;
            return "";
        }

        private static bool IsMonthToken(string value)
        {
            return value.Length == 6
                && int.TryParse(value[..2], out var month)
                && month >= 1
                && month <= 12
                && int.TryParse(value[2..], out _);
        }

        private ConsolidatedStatement Reconciliate(FileInfo csv, IEnumerable<FileInfo> pdfs)
        {
            var arStatement = _arExtractor.Extract(pdfs.Select(t => t.FullName));

            RxOfficeStatement rxOfficeStatement = _rxOfficeExtractor.Extract(csv.FullName);

            var statement = new ConsolidatedStatement(arStatement, rxOfficeStatement);

            return statement;
        }
    }
}
