using ExtractStatementPDF.AR;
using ExtractStatementPDF.RxOffice;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.Consolidation
{
    public record FileMatchResult(
        string Name,
        IEnumerable<FileInfo> CsvFiles,
        IEnumerable<FileInfo> ArFiles
    );

    public record ProgressUpdate(string Message);

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
                    case ".ods":
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

                    MoveFilesTo([], [file], archiveDirectory);
                }
                else
                {
                    MoveFilesTo([], [file], verifiedDirectory);

                }
            }
        }

        public void VerifyMatches(string directory)
        {
            var directoryInfo = new DirectoryInfo(directory);
            var matchedDirectory = EnsureDirectoryExists(directoryInfo.FullName, Subdirectories.Matched);

            var arCopies = new List<FileInfo>();
            var csvs = new List<FileInfo>();

            foreach (var file in directoryInfo.GetFiles($"{Subdirectories.Verified}/*", SearchOption.AllDirectories))
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

            foreach (var match in matches)
            {
                if (match.ArFiles.Count() > 0)
                {
                    MoveFilesTo(match.CsvFiles, match.ArFiles, matchedDirectory);
                }
            }
        }

        public IObservable<ProgressUpdate> Process(string directory)
        {
            var subject = new Subject<ProgressUpdate>();

            Task.Run(() =>
            {
                try
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
                        var statement = Reconciliate(match);

                        if (statement.IsValid())
                        {
                            statements.Add(statement);
                            MoveFilesTo(match.CsvFiles, match.ArFiles, processedDirectory);
                        }
                        else
                        {
                            MoveFilesTo(match.CsvFiles, match.ArFiles, archivedDirectory);
                        }

                        subject.OnNext(new ProgressUpdate(match.Name));
                    }

                    foreach (var statement in statements)
                    {
                        var bytes = _excelGenerator.GenerateExcel(statement);
                        var filename = Path.GetFileNameWithoutExtension(statement.Filename) + ".xlsx";
                        using var filestream = new FileStream($"{directory}/{filename}", FileMode.CreateNew, FileAccess.Write);
                        filestream.Write(bytes);
                    }

                    Update(statements);

                    subject.OnCompleted();
                }
                catch (Exception ex)
                {
                    subject.OnError(ex);
                }
                finally
                {
                    subject.Dispose();
                }
            });

            return subject;
        }

        private static string EnsureDirectoryExists(string rootDirectory, string subdirectory)
        {
            var directory = Path.Combine(rootDirectory, subdirectory);
            Directory.CreateDirectory(directory);

            return directory;
        }

        private static void MoveFilesTo(IEnumerable<FileInfo> csvs, IEnumerable<FileInfo> arFiles, string directory)
        {
            foreach (var file in csvs)
            {
                MoveToDirectory(file, directory);
            }

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

        private static IEnumerable<FileMatchResult> MatchFiles(List<FileInfo> arCopies, List<FileInfo> csvs)
        {
            var filenames = new List<FileInfo>().Concat(arCopies).Concat(csvs);
            var lookupKeys = filenames.Select(t => BuildLookupKey(t.FullName)).Distinct().Order();
            
            var matches = new List<FileMatchResult>();

            foreach (var lookupKey in lookupKeys)
            {
                var matchingCsvs = csvs.Where(t => BuildLookupKey(t.FullName) == lookupKey);
                var matchingPdfs = arCopies.Where(t => BuildLookupKey(t.FullName) == lookupKey);

                matches.Add(new FileMatchResult(lookupKey, matchingCsvs, matchingPdfs));
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

        private ConsolidatedStatement Reconciliate(FileMatchResult fileMatchResult)
        {
            var arStatement = _arExtractor.Extract(fileMatchResult.ArFiles.Select(t => t.FullName));
            var rxOfficeStatement = _rxOfficeExtractor.Extract(fileMatchResult.CsvFiles.Select(t => t.FullName));

            var statement = new ConsolidatedStatement(fileMatchResult.Name, arStatement, rxOfficeStatement);

            return statement;
        }
    }
}
