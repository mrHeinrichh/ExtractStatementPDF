using CsvHelper;
using System.Globalization;

namespace ExtractStatementPDF.RxOffice
{
    public class RxOfficeExtractor
    {
        public RxOfficeStatement Extract(string fullpath)
        {
            using var reader = new StreamReader(fullpath);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            csv.Context.RegisterClassMap<RxOfficeOrderMap>();

            var records = csv.GetRecords<RxOfficeOrder>();

            var filename = Path.GetFileName(fullpath);
            var statement = new RxOfficeStatement(filename);
            statement.AddOrders(records);

            return statement;
        }
    }
}
