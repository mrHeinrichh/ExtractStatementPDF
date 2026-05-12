using CsvHelper.Configuration;

namespace ExtractStatementPDF.RxOffice
{
    public class RxOfficeOrderMap : ClassMap<RxOfficeOrder>
    {
        public RxOfficeOrderMap()
        {
            Map(t => t.Date).Index(11);

            Map(t => t.Reference).Index(12);

            Map(t => t.Gross).Convert(t =>
            {
                var value = t.Row.GetField(13);
                return decimal.Parse(value ?? "");
            });

            Map(t => t.Discount).Convert(t =>
            {
                var value = t.Row.GetField(14);
                return decimal.Parse(value ?? "");
            });

            Map(t => t.Net).Convert(t =>
            {
                var value = t.Row.GetField(15);
                return decimal.Parse(value ?? "");
            });
        }
    }
}
