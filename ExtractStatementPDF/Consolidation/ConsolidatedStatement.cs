using ExtractStatementPDF.AR;
using ExtractStatementPDF.RxOffice;

namespace ExtractStatementPDF.Consolidation
{
    public class ConsolidatedStatement
    {
        public ARStatement ARStatement;

        public RxOfficeStatement RxOfficeStatement;

        public List<ConsolidatedOrder> ConsolidatedOrders { get; private set; }

        public string Filename => RxOfficeStatement.Filename;

        public string Customer { get; private set; }

        public string Month { get; private set; }

        public string Period { get; private set; }

        public ConsolidatedStatement(
            ARStatement arStatement,
            RxOfficeStatement rxOfficeStatement)
        {
            ARStatement = arStatement;
            RxOfficeStatement = rxOfficeStatement;

            ParseFilename(RxOfficeStatement.Filename);

            var arOrders = arStatement.Orders;
            var rxOfficeOrders = rxOfficeStatement.Orders;

            ConsolidatedOrders = (
                from a in arOrders
                join b in rxOfficeOrders on a.Reference equals b.Reference into gj
                from subB in gj.DefaultIfEmpty()
                select new ConsolidatedOrder { AROrder = a, RxOfficeOrder = subB }
            )
            .Union(
                from b in rxOfficeOrders
                join a in arOrders on b.Reference equals a.Reference into gj
                from subA in gj.DefaultIfEmpty()
                where subA == null
                select new ConsolidatedOrder { AROrder = subA, RxOfficeOrder = b }
            )
            .OrderBy(t => DateTime.Parse(t.Date))
            .ToList();
        }

        private void ParseFilename(string fileName)
        {
            var parts = Path
                .GetFileNameWithoutExtension(fileName)
                .Split("_", StringSplitOptions.RemoveEmptyEntries);

            Customer = parts.ElementAtOrDefault(0) ?? "";

            Month = parts.FirstOrDefault(IsMonthToken) ?? "";

            Period = parts
                .Skip(1)
                .FirstOrDefault(t => !string.Equals(t, Month, StringComparison.OrdinalIgnoreCase))
                ?? "";
        }

        private static bool IsMonthToken(string value)
        {
            return value.Length == 6
                && int.TryParse(value[..2], out var month)
                && month >= 1
                && month <= 12
                && int.TryParse(value[2..], out _);
        }

        public decimal NetVariance()
        {
            return Math.Abs(ARStatement.TotalNet() - RxOfficeStatement.TotalNet());
        }

        public decimal CumulativeVariance()
        {
            return ConsolidatedOrders.Sum(t => t.Variance);
        }
    }
}
