using ExtractStatementPDF.Models;
using ExtractStatementPDF.RxOffice;

namespace ExtractStatementPDF.Consolidation
{
    public class ConsolidatedOrder
    {
        public AROrder? AROrder { get; set; }

        public RxOfficeOrder? RxOfficeOrder { get; set; }

        public string Reference => AROrder?.Reference ?? RxOfficeOrder?.Reference ?? "-";

        public string Date => AROrder?.Date ?? RxOfficeOrder?.Date ?? "-";

        public decimal Variance => Math.Abs((AROrder?.Net ?? 0m) - (RxOfficeOrder?.Net ?? 0m));

        public string Remarks
        {
            get
            {
                if (AROrder == null)
                {
                    if (RxOfficeOrder?.Net == 0m)
                    {
                        return "Warranty Replacement";
                    }

                    return "AR Order is missing";
                }
                if (RxOfficeOrder == null)
                {
                    return "RxOffice Order is missing";
                }
                if (RxOfficeOrder.Net == 0m && AROrder.Net > 0)
                {
                    return "RxOffice price didn't compute";
                }
                if (RxOfficeOrder.Net != AROrder.Net)
                {
                    return "NET discrepancy";
                }

                return "Good";
            }
        }
    }
}
