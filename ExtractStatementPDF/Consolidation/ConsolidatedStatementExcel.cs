using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExtractStatementPDF.Consolidation
{
    public class ConsolidatedStatementExcel
    {
        public ConsolidatedStatementExcel()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Aaron del Rosario");
        }

        public byte[] GenerateExcel(ConsolidatedStatement statement)
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Statement");

            ws.Cells[1, 1].Value = "Reference";
            ws.Cells[1, 2].Value = "Date";
            ws.Cells[1, 3].Value = "Rx Gross";
            ws.Cells[1, 4].Value = "Rx Discount";
            ws.Cells[1, 5].Value = "Rx NET";
            ws.Cells[1, 6].Value = "AR Gross";
            ws.Cells[1, 7].Value = "AR Discount";
            ws.Cells[1, 8].Value = "AR NET";
            ws.Cells[1, 9].Value = "Variance";
            ws.Cells[1, 10].Value = "Auto Remarks";
            ws.Cells[1, 11].Value = "Comments";

            var firstRow = 2;
            var lastRow = 0;

            var row = firstRow;
            foreach (var order in statement.ConsolidatedOrders)
            {
                lastRow = row;

                ws.Cells[row, 1].Value = order.Reference;
                ws.Cells[row, 2].Value = order.Date;
                ws.Cells[row, 3].Value = order.RxOfficeOrder?.Gross;
                ws.Cells[row, 4].Value = order.RxOfficeOrder?.Discount;
                ws.Cells[row, 5].Value = order.RxOfficeOrder?.Net;
                ws.Cells[row, 6].Value = order.AROrder?.Gross;
                ws.Cells[row, 7].Value = order.AROrder?.Discount;
                ws.Cells[row, 8].Value = order.AROrder?.Net;
                ws.Cells[row, 9].Value = order.Variance;
                ws.Cells[row, 10].Value = order.Remarks;

                row++;
            }

            var footerRow = lastRow + 1;

            ws.Cells[footerRow, 3].Formula = $"SUM(C{firstRow}:C{lastRow})";
            ws.Cells[footerRow, 4].Formula = $"SUM(D{firstRow}:D{lastRow})";
            ws.Cells[footerRow, 5].Formula = $"SUM(E{firstRow}:E{lastRow})";
            ws.Cells[footerRow, 6].Formula = $"SUM(F{firstRow}:F{lastRow})";
            ws.Cells[footerRow, 7].Formula = $"SUM(G{firstRow}:G{lastRow})";
            ws.Cells[footerRow, 8].Formula = $"SUM(H{firstRow}:H{lastRow})";
            ws.Cells[footerRow, 9].Formula = $"SUM(I{firstRow}:I{lastRow})";

            ws.Cells["A1:K1"].Style.Font.Bold = true;
            ws.Cells["A1:K1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            ws.Cells[$"A{footerRow}:J{footerRow}"].Style.Font.Bold = true;
            ws.Cells["A1:J1"].Style.Border.Top.Style = ExcelBorderStyle.Thick;

            ws.Cells[$"A{firstRow}:I{footerRow}"].Style.Numberformat.Format = "###,###,##0.00";

            ws.Cells.AutoFitColumns();

            return package.GetAsByteArray();
        }
    }
}
