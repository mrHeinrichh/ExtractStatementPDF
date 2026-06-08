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
            try
            {
                if (!fullpaths.Any()) return new RxOfficeStatement("");

                var filename = Path.GetFileName(fullpaths.First());

                var statement = new RxOfficeStatement(filename);

                foreach (var fullpath in fullpaths)
                {
                    using var reader = new StreamReader(fullpath);
                    using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

                    csv.Read();
                    csv.ReadHeader();
                    var count = csv.Parser.Count;

                    if (count == 16)
                    {
                        csv.Context.RegisterClassMap<RxOfficeOrderMap>();
                    }
                    else if (count == 27)
                    {
                        csv.Context.RegisterClassMap<RxOfficeOrderMapV2>();
                    }
                    else
                    {
                        Console.WriteLine($"Unexpected column count ({count}) in file: {fullpath}");
                        continue; // Skip this file and move to the next one
                    }

                    var records = csv.GetRecords<RxOfficeOrder>();

                    statement.AddOrders(records);
                }

                return statement;
            }
            catch
            {
                // Log the exception or handle it as needed
                Console.WriteLine($"Failed to extract csv: {string.Join(" ,", fullpaths)}");
                return new RxOfficeStatement("");
            }
        }
    }
}
