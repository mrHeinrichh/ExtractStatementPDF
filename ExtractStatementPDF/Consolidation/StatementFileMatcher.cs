using System.Text.RegularExpressions;

namespace ExtractStatementPDF.Consolidation
{
    internal static class StatementFileMatcher
    {
        public static string GetLookupKey(string fullName)
        {
            return BuildLookupKey(fullName);
        }

        public static Dictionary<FileInfo, IEnumerable<FileInfo>> MatchFiles(List<FileInfo> arCopies, List<FileInfo> csvs)
        {
            var matches = new Dictionary<FileInfo, IEnumerable<FileInfo>>();
            foreach (var currentCsv in csvs)
            {
                var csvKey = BuildLookupKey(currentCsv.FullName);
                var matchingCandidates = arCopies
                    .Where(t => BuildLookupKey(t.FullName) == csvKey)
                    .ToList();

                var preferredCandidates = matchingCandidates
                    .Where(IsSourceFile)
                    .ToList();

                var matchingFiles = (preferredCandidates.Count > 0 ? preferredCandidates : matchingCandidates)
                    .OrderByDescending(t => string.Equals(t.Extension, ".pdf", StringComparison.OrdinalIgnoreCase))
                    .ThenBy(t => t.Name);

                matches.Add(currentCsv, matchingFiles);
            }

            return matches;
        }

        public static bool IsInFolder(FileInfo file, string folderName)
        {
            var directory = file.Directory;
            while (directory != null)
            {
                if (string.Equals(directory.Name, folderName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                directory = directory.Parent;
            }

            return false;
        }

        private static string BuildLookupKey(string fullName)
        {
            var name = Path.GetFileNameWithoutExtension(fullName);

            var underscoreParts = name.Split("_", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var monthTokenIndex = Array.FindIndex(underscoreParts, IsMonthToken);
            var monthToken = monthTokenIndex >= 0 ? underscoreParts[monthTokenIndex] : string.Empty;
            var periodToken = underscoreParts.LastOrDefault(IsPeriodToken) ?? string.Empty;

            if (monthTokenIndex > 0)
            {
                name = string.Join("_", underscoreParts.Take(monthTokenIndex));
            }

            var normalizedName = Regex.Replace(name, @"[^A-Za-z0-9]+", "").ToLowerInvariant();
            if (periodToken.Length == 0)
            {
                return monthToken.Length == 0
                    ? normalizedName
                    : $"{normalizedName}|{monthToken.ToLowerInvariant()}";
            }

            return monthToken.Length == 0
                ? $"{normalizedName}|{periodToken.ToLowerInvariant()}"
                : $"{normalizedName}|{monthToken.ToLowerInvariant()}|{periodToken.ToLowerInvariant()}";
        }

        private static bool IsMonthToken(string value)
        {
            return value.Length == 6
                && int.TryParse(value[..2], out var month)
                && month >= 1
                && month <= 12
                && int.TryParse(value[2..], out _);
        }

        private static bool IsPeriodToken(string value)
        {
            return Regex.IsMatch(value, @"^\d+(?:st|nd|rd|th)$", RegexOptions.IgnoreCase);
        }

        private static bool IsSourceFile(FileInfo file)
        {
            return IsInFolder(file, "Source");
        }
    }
}
