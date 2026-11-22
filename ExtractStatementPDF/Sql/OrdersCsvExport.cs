using System.Data.Common;

namespace ExtractStatementPDF.Sql
{
    // "Download Orders CSV": the order lines of every order created in a period,
    // read from the RxOffice database and streamed straight into a CSV file.
    public static class OrdersCsvExport
    {
        public const string DefaultDatabase = "rxoffice";

        public const string Query = """
            SELECT oh.Id, ol.Code, ol.Qty, ol.UnitListPrice, ol.UnitDiscountedPrice
            FROM orderheaders oh
            INNER JOIN orderlines ol
              ON oh.Id = ol.OrderHeaderId
            WHERE oh.CreatedOn
              BETWEEN @from AND @to;
            """;

        // Whole days: From 00:00:00 through To 23:59:59.
        public static (DateTime From, DateTime To) Period(DateTime from, DateTime to) =>
            (from.Date, to.Date.AddDays(1).AddSeconds(-1));

        // The query as it will run, with the dates filled in (for display only).
        public static string Preview(DateTime from, DateTime to)
        {
            var (start, end) = Period(from, to);
            return Query
                .Replace("@from", $"'{start:yyyy-MM-dd HH:mm:ss}'")
                .Replace("@to", $"'{end:yyyy-MM-dd HH:mm:ss}'");
        }

        // Returns the number of rows written.
        public static long Download(SqlSession session, string database, DateTime from, DateTime to, string path)
        {
            var (start, end) = Period(from, to);

            using var connection = session.CreateConnection(string.IsNullOrWhiteSpace(database) ? null : database.Trim());
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText = Query;
            command.CommandTimeout = 600;
            AddParameter(command, "@from", start);
            AddParameter(command, "@to", end);

            using var reader = command.ExecuteReader();
            return SqlImportExport.ExportCsv(reader, path);
        }

        private static void AddParameter(DbCommand command, string name, object value)
        {
            var parameter = command.CreateParameter();
            parameter.ParameterName = name;
            parameter.Value = value;
            command.Parameters.Add(parameter);
        }
    }
}
