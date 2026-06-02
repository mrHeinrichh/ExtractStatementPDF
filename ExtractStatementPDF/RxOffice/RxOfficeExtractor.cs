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

        public RxOfficeStatement Extract(IEnumerable<string> fullpaths)
        {
            if (!fullpaths.Any()) return new RxOfficeStatement("");

            var filename = Path.GetFileName(fullpaths.First());

            var statement = new RxOfficeStatement(filename);

            foreach (var fullpath in fullpaths)
            {
                using var reader = new StreamReader(fullpath);
                using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
                csv.Context.RegisterClassMap<RxOfficeOrderMap>();

                var records = csv.GetRecords<RxOfficeOrder>();

                statement.AddOrders(records);
            }

            return statement;
        }
    }
}
