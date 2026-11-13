using Microsoft.Data.SqlClient;
using MySqlConnector;
using System.Data.Common;
using System.Text.Json.Serialization;

namespace ExtractStatementPDF.Sql
{
    public enum SqlEngine
    {
        MySql,
        SqlServer,
    }

    public enum SqlAuthentication
    {
        SqlLogin,
        Windows,
    }

    // One saved connection ("session" in HeidiSQL terms). The password and a custom
    // connection string are never written in plain text; see SqlSessionStore.
    public class SqlSession
    {
        public string Name { get; set; } = "New session";

        public SqlEngine Engine { get; set; } = SqlEngine.MySql;

        public string Host { get; set; } = "localhost";

        public int Port { get; set; } = DefaultPort(SqlEngine.MySql);

        public SqlAuthentication Authentication { get; set; } = SqlAuthentication.SqlLogin;

        public string User { get; set; } = "root";

        [JsonIgnore]
        public string Password { get; set; } = "";

        public bool SavePassword { get; set; } = true;

        public string Database { get; set; } = "";

        public bool Encrypt { get; set; }

        public bool TrustServerCertificate { get; set; } = true;

        public int TimeoutSeconds { get; set; } = 15;

        public bool UseCustomConnectionString { get; set; }

        [JsonIgnore]
        public string CustomConnectionString { get; set; } = "";

        // DPAPI-protected copies, filled in by SqlSessionStore.
        public string? ProtectedPassword { get; set; }

        public string? ProtectedConnectionString { get; set; }

        public static int DefaultPort(SqlEngine engine) => engine == SqlEngine.MySql ? 3306 : 1433;

        public static string EngineName(SqlEngine engine) => engine == SqlEngine.MySql ? "MySQL / MariaDB" : "Microsoft SQL Server";

        public SqlSession Clone() => (SqlSession)MemberwiseClone();

        public override string ToString() => Name;

        public string DescribeTarget()
        {
            if (UseCustomConnectionString) return $"{EngineName(Engine)} (custom connection string)";

            var who = Engine == SqlEngine.SqlServer && Authentication == SqlAuthentication.Windows ? "Windows login" : User;
            return $"{who}@{DataSource()}";
        }

        // Extra advice shown under a connection error.
        public string ConnectionHint()
        {
            var host = Host.Trim().ToLowerInvariant();
            var local = host is "localhost" or "127.0.0.1" or "::1" or "." or "(local)";

            if (!UseCustomConnectionString && local)
            {
                return "\"localhost\" means this PC. If the database is on another computer "
                    + "(e.g. the RxOffice server 192.168.1.30), enter that IP address under Hostname / IP.";
            }

            return $"Check that the server is switched on and reachable from this PC (VPN / ZeroTier connected) and that port {Port} is open.";
        }

        // Builds the connection string from the fields (or uses the custom one).
        // `database` overrides the default database; `maskPassword` is for display only.
        public string BuildConnectionString(string? database = null, bool maskPassword = false)
        {
            var password = maskPassword && Password.Length > 0 ? "********" : Password;

            if (UseCustomConnectionString)
            {
                if (database == null) return CustomConnectionString;

                return Engine == SqlEngine.MySql
                    ? new MySqlConnectionStringBuilder(CustomConnectionString) { Database = database }.ConnectionString
                    : new SqlConnectionStringBuilder(CustomConnectionString) { InitialCatalog = database }.ConnectionString;
            }

            if (Engine == SqlEngine.MySql)
            {
                var builder = new MySqlConnectionStringBuilder
                {
                    Server = Host.Trim(),
                    Port = (uint)Port,
                    UserID = User,
                    Password = password,
                    ConnectionTimeout = (uint)TimeoutSeconds,
                    // Same as RxOffice's own connection: zero dates read as DateTime.MinValue.
                    ConvertZeroDateTime = true,
                    AllowUserVariables = true,
                };

                var db = database ?? Database;
                if (!string.IsNullOrWhiteSpace(db)) builder.Database = db.Trim();

                return builder.ConnectionString;
            }
            else
            {
                var builder = new SqlConnectionStringBuilder
                {
                    DataSource = DataSource(),
                    ConnectTimeout = TimeoutSeconds,
                    Encrypt = Encrypt ? SqlConnectionEncryptOption.Mandatory : SqlConnectionEncryptOption.Optional,
                    TrustServerCertificate = TrustServerCertificate,
                    ApplicationName = "ExtractStatementPDF",
                };

                if (Authentication == SqlAuthentication.Windows)
                {
                    builder.IntegratedSecurity = true;
                }
                else
                {
                    builder.UserID = User;
                    builder.Password = password;
                }

                var db = database ?? Database;
                if (!string.IsNullOrWhiteSpace(db)) builder.InitialCatalog = db.Trim();

                return builder.ConnectionString;
            }
        }

        public DbConnection CreateConnection(string? database = null)
        {
            var connectionString = BuildConnectionString(database);

            return Engine == SqlEngine.MySql
                ? new MySqlConnection(connectionString)
                : new SqlConnection(connectionString);
        }

        // SQL Server takes "host,port"; a named instance ("host\SQLEXPRESS") or port 0
        // lets the server's browser service pick the port.
        private string DataSource()
        {
            var host = Host.Trim();

            if (Engine == SqlEngine.MySql) return $"{host}:{Port}";
            if (Port <= 0 || host.Contains('\\') || host.StartsWith("(localdb)", StringComparison.OrdinalIgnoreCase)) return host;

            return $"{host},{Port}";
        }
    }
}
