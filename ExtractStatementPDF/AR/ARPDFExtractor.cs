using ExtractStatementPDF.Models;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.AR
{
    public class ARPDFExtractor
    {
        public ARStatement Extract(IEnumerable<string> filenames)
        {
            var arStatement = new ARStatement();

            foreach (var filename in filenames)
            {
                var pages = GetPageContents(filename);

                foreach (var page in pages)
                {
                    var orders = ParsePage(page);
                    arStatement.AddOrders(orders);
                }

                arStatement.AddPages(pages);
            }

            return arStatement;
        }

        public ARStatement Extract(string filename)
        {
            return Extract([filename]);
        }

        public List<string> ExtractReferences(IEnumerable<string> filenames)
        {
            var references = new List<string>();

            foreach (var filename in filenames)
            {
                var pages = GetPageContents(filename);
                foreach (var page in pages)
                {
                    references.AddRange(ParseReferences(page));
                }
            }

            return references;
        }

        private List<AROrder> ParsePage(string pageContent)
        {
            using var reader = new StringReader(pageContent);

            string? line;
            var arOrders = new List<AROrder>();
            while ((line = reader.ReadLine()) != null)
            {
                var arorder = ParseLine(line);

                if (arorder != null) arOrders.Add(arorder);
            }

            return arOrders;
        }

        private IEnumerable<string> ParseReferences(string pageContent)
        {
            using var reader = new StringReader(pageContent);

            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                var reference = ParseReference(line);
                if (!string.IsNullOrWhiteSpace(reference))
                {
                    yield return reference;
                }
            }
        }

        private AROrder? ParseLine(string line)
        {
            var pattern = @"(\d{2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2})\s(\d+)\s([\d,.]+)\s([\d,.]+)\s([\d,.]+)";

            var match = Regex.Match(line, pattern);

            if (match.Success)
            {
                var arOrder = new AROrder()
                {
                    Date = match.Groups[1].Value,
                    Reference = match.Groups[2].Value,
                    Gross = decimal.Parse(match.Groups[3].Value, NumberStyles.Number, CultureInfo.InvariantCulture),
                    Discount = decimal.Parse(match.Groups[4].Value, NumberStyles.Number, CultureInfo.InvariantCulture),
                    Net = decimal.Parse(match.Groups[5].Value, NumberStyles.Number, CultureInfo.InvariantCulture),
                };

                return arOrder;
            }

            return null;
        }

        private static string? ParseReference(string line)
        {
            var pattern = @"(\d{2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2})\s([A-Za-z0-9-]+)\s";
            var match = Regex.Match(line, pattern);

            if (!match.Success)
            {
                return null;
            }

            return match.Groups[2].Value;
        }

        private List<string> GetPageContents(string filename)
        {
            try
            {
                using var reader = new PdfReader(filename);
                using var document = new PdfDocument(reader);
                var strategy = new SimpleTextExtractionStrategy();
                var content = new List<string>();

                for (var index = 1; index <= document.GetNumberOfPages(); index++)
                {
                    var page = document.GetPage(index);

                    content.Add(PdfTextExtractor.GetTextFromPage(page, strategy));
                }

                return content;
            }
            catch (Exception e)
            {
                Debug.WriteLine($"Error reading pdf: {filename}");
                return [];
            }
        }
    }
}
