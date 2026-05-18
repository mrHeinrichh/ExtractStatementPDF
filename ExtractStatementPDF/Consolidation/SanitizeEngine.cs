using ExtractStatementPDF.AR;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.Consolidation
{
    public sealed class SanitizeResult
    {
        public string ScannedDirectory { get; set; } = string.Empty;

        public int ArchivedMatchCount { get; set; }

        public int MovedFileCount { get; set; }

        public List<string> ArchiveDirectories { get; } = [];

        public List<string> MovedFiles { get; } = [];
    }

    public class SanitizeEngine
    {
        private readonly ARPDFExtractor _arPdfExtractor = new();
        private readonly ARExcelExtractor _arExcelExtractor = new();

        public SanitizeResult Process(string directory)
        {
            var result = new SanitizeResult();
            var directoryInfo = new DirectoryInfo(directory);
            result.ScannedDirectory = directoryInfo.FullName;

            var arSources = new List<FileInfo>();

            foreach (var file in directoryInfo.GetFiles("*", SearchOption.AllDirectories))
            {
                if (ShouldSkip(file))
                {
                    continue;
                }

                switch (file.Extension.ToLowerInvariant())
                {
                    case ".pdf":
                    case ".xls":
                    case ".xlsx":
                        arSources.Add(file);
                        break;
                }
            }

            var groupedSources = arSources
                .GroupBy(file => StatementFileMatcher.GetLookupKey(file.FullName), StringComparer.OrdinalIgnoreCase);

            foreach (var group in groupedSources)
            {
                var sourceFiles = group
                    .OrderBy(file => file.Name)
                    .ToList();

                if (!ShouldArchive(sourceFiles))
                {
                    continue;
                }

                ArchiveMatch(directoryInfo.FullName, sourceFiles, result);
            }

            return result;
        }

        private static bool ShouldSkip(FileInfo file)
        {
            return StatementFileMatcher.IsInFolder(file, "Archive")
                || StatementFileMatcher.IsInFolder(file, "Sanitized");
        }

        private bool ShouldArchive(IEnumerable<FileInfo> sourceFiles)
        {
            var references = new List<string>();
            var pdfFiles = sourceFiles
                .Where(file => string.Equals(file.Extension, ".pdf", StringComparison.OrdinalIgnoreCase))
                .Select(file => file.FullName);
            var excelFiles = sourceFiles
                .Where(file => string.Equals(file.Extension, ".xls", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(file.Extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
                .Select(file => file.FullName);

            references.AddRange(_arPdfExtractor.ExtractReferences(pdfFiles));
            references.AddRange(_arExcelExtractor.ExtractReferences(excelFiles));

            if (references.Count == 0)
            {
                return false;
            }

            return references.Any(reference =>
                !string.IsNullOrWhiteSpace(reference) &&
                Regex.IsMatch(reference, "[A-Za-z]"));
        }

        private static void ArchiveMatch(string rootDirectory, IEnumerable<FileInfo> sourceFiles, SanitizeResult result)
        {
            var archiveRoot = ResolveArchiveRoot(rootDirectory, sourceFiles);
            var filesToArchive = sourceFiles
                .DistinctBy(t => t.FullName, StringComparer.OrdinalIgnoreCase);

            result.ArchivedMatchCount++;
            if (!result.ArchiveDirectories.Contains(archiveRoot, StringComparer.OrdinalIgnoreCase))
            {
                result.ArchiveDirectories.Add(archiveRoot);
            }

            foreach (var file in filesToArchive)
            {
                var destination = MoveToArchive(rootDirectory, archiveRoot, file);
                result.MovedFileCount++;
                result.MovedFiles.Add(destination);
            }
        }

        private static string ResolveArchiveRoot(string rootDirectory, IEnumerable<FileInfo> arFiles)
        {
            var sourceDirectory = arFiles
                .Select(file => FindContainingFolder(file, "Source"))
                .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path));

            return string.IsNullOrWhiteSpace(sourceDirectory)
                ? Path.Combine(rootDirectory, "Archive")
                : Path.Combine(sourceDirectory, "Archive");
        }

        private static string? FindContainingFolder(FileInfo file, string folderName)
        {
            var directory = file.Directory;
            while (directory != null)
            {
                if (string.Equals(directory.Name, folderName, StringComparison.OrdinalIgnoreCase))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            return null;
        }

        private static string MoveToArchive(string rootDirectory, string archiveRoot, FileInfo file)
        {
            var relativePath = TryGetArchiveRelativePath(file, archiveRoot)
                ?? Path.GetRelativePath(rootDirectory, file.FullName);
            var destination = Path.Combine(archiveRoot, relativePath);
            var destinationDirectory = Path.GetDirectoryName(destination);

            if (!string.IsNullOrWhiteSpace(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
            }

            try
            {
                File.Move(file.FullName, destination, overwrite: true);
            }
            catch (IOException)
            {
                FileLockResolver.KillLockingProcesses(file.FullName);
                if (File.Exists(destination))
                {
                    FileLockResolver.KillLockingProcesses(destination);
                }

                File.Move(file.FullName, destination, overwrite: true);
            }

            return destination;
        }

        private static string? TryGetArchiveRelativePath(FileInfo file, string archiveRoot)
        {
            var sourceRoot = Directory.GetParent(archiveRoot)?.FullName;
            if (string.IsNullOrWhiteSpace(sourceRoot))
            {
                return null;
            }

            if (!file.FullName.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
            {
                return file.Name;
            }

            return Path.GetRelativePath(sourceRoot, file.FullName);
        }
    }
}
