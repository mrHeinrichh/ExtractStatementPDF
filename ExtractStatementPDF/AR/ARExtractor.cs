using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractStatementPDF.AR
{
    public class ARExtractor
    {
        private readonly ARPDFExtractor _pdfExtractor = new();

        private readonly ARExcelExtractor _excelExtractor = new();

        public ARStatement Extract(IEnumerable<string> filenames)
        {
            var candidates = filenames
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var pdfFiles = candidates
                .Where(t => string.Equals(Path.GetExtension(t), ".pdf", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (pdfFiles.Count > 0)
            {
                return _pdfExtractor.Extract(pdfFiles);
            }

            var excelFiles = candidates
                .Where(t =>
                    string.Equals(Path.GetExtension(t), ".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(Path.GetExtension(t), ".xls", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (excelFiles.Count > 0)
            {
                return _excelExtractor.Extract(excelFiles);
            }

            return new ARStatement([]);
        }

        public ARStatement Extract(string filename)
        {
            return Extract([filename]);
        }
    }
}
