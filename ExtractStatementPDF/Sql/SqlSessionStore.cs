using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace ExtractStatementPDF.Sql
{
    // Saves SQL sessions and the last query under %AppData%\ExtractStatementPDF.
    // Passwords and custom connection strings are encrypted with Windows DPAPI
    // (current user only), so the file is useless on another account or PC.
    public static class SqlSessionStore
    {
        private static readonly string Folder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExtractStatementPDF");

        private static readonly string SessionsFile = Path.Combine(Folder, "sql-sessions.json");

        private static readonly string QueryFile = Path.Combine(Folder, "sql-last-query.sql");

        private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

        private class StoredSessions
        {
            public List<SqlSession> Sessions { get; set; } = [];

            public string? LastSession { get; set; }
        }

        public static (List<SqlSession> Sessions, string? LastSession) Load()
        {
            try
            {
                if (!File.Exists(SessionsFile)) return ([], null);

                var stored = JsonSerializer.Deserialize<StoredSessions>(File.ReadAllText(SessionsFile)) ?? new StoredSessions();

                foreach (var session in stored.Sessions)
                {
                    session.Password = Unprotect(session.ProtectedPassword);
                    session.CustomConnectionString = Unprotect(session.ProtectedConnectionString);
                }

                return (stored.Sessions, stored.LastSession);
            }
            catch
            {
                return ([], null);
            }
        }

        public static void Save(IEnumerable<SqlSession> sessions, string? lastSession)
        {
            var stored = new StoredSessions { LastSession = lastSession };

            foreach (var session in sessions)
            {
                var copy = session.Clone();
                copy.ProtectedPassword = session.SavePassword ? Protect(session.Password) : null;
                copy.ProtectedConnectionString = session.UseCustomConnectionString ? Protect(session.CustomConnectionString) : null;
                stored.Sessions.Add(copy);
            }

            Directory.CreateDirectory(Folder);
            File.WriteAllText(SessionsFile, JsonSerializer.Serialize(stored, JsonOptions));
        }

        public static string LoadQuery()
        {
            try
            {
                return File.Exists(QueryFile) ? File.ReadAllText(QueryFile) : "";
            }
            catch
            {
                return "";
            }
        }

        public static void SaveQuery(string query)
        {
            try
            {
                Directory.CreateDirectory(Folder);
                File.WriteAllText(QueryFile, query);
            }
            catch
            {
                // Losing the scratch query is not worth an error on close.
            }
        }

        private static string? Protect(string value)
        {
            if (string.IsNullOrEmpty(value)) return null;

            var bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(value), null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(bytes);
        }

        private static string Unprotect(string? value)
        {
            if (string.IsNullOrEmpty(value)) return "";

            try
            {
                var bytes = ProtectedData.Unprotect(Convert.FromBase64String(value), null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(bytes);
            }
            catch
            {
                return "";
            }
        }
    }
}
