namespace ExtractStatementPDF.Models
{
    public class AROrder
    {
        public string Date { get; set; }

        public string Reference { get; set; }

        public decimal Gross { get; set; }

        public decimal Discount { get; set; }

        public decimal Net { get; set; }
    }
}
