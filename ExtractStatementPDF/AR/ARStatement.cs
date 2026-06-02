using ExtractStatementPDF.Models;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.AR
{
    public class ARStatement
    {
        public List<string> Filenames { get; set; } = [];

        public List<string> Pages { get; set; } = [];

        public List<AROrder> Orders { get; set; } = [];

        public ARStatement(IEnumerable<string> filenames)
        {
            Filenames = Filenames;
        }

        public bool IsValid()
        {
            var manualOrders = Orders.Where(t => Regex.IsMatch(t.Reference, "[A-Za-z]")).Count();

            return manualOrders <= 1;
        }

        public void AddOrders(List<AROrder> orders)
        {
            foreach (var order in orders)
            {
                if (Orders.Any(t => t.Reference == order.Reference))
                {
                    continue;
                }

                Orders.Add(order);
            }
        }

        public void AddPages(List<string> pages)
        {
            Pages.AddRange(pages);
        }

        public decimal TotalNet()
        {
            return Orders.Select(t => t.Net).Aggregate(0m, (t1, t2) => t1 + t2);
        }
    }
}
