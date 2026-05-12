namespace ExtractStatementPDF.RxOffice
{
    public class RxOfficeStatement
    {
        public string Filename { get; }

        public RxOfficeStatement(string filename)
        {
            Filename = filename;
        }

        public List<RxOfficeOrder> Orders { get; set; } = [];

        public void AddOrders(IEnumerable<RxOfficeOrder> orders)
        {
            Orders.AddRange(orders);
        }

        public decimal TotalNet()
        {
            return Orders.Select(t => t.Net).Aggregate(0m, (t1, t2) => t1 + t2);
        }
    }
}
