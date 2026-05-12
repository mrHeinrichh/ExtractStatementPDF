using ExtractStatementPDF.AR;
using ExtractStatementPDF.Consolidation;
using ExtractStatementPDF.RxOffice;

namespace ExtractStatementPDF
{
    public partial class Form1 : Form
    {
        private ARStatement? arStatement;

        private RxOfficeStatement? rxOfficeStatement;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var fullpath = SelectARFile();
            if (string.IsNullOrEmpty(fullpath)) return;

            var extractor = new ARExtractor();
            arStatement = extractor.Extract(fullpath);

            richTextBox1.Text = string.Join(Environment.NewLine, arStatement.Pages);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var fullpath = SelectCSV();
            if (string.IsNullOrEmpty(fullpath)) return;

            var extractor = new RxOfficeExtractor();
            rxOfficeStatement = extractor.Extract(fullpath);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (arStatement == null) return;
                if (rxOfficeStatement == null) return;

                var statement = new ConsolidatedStatement(arStatement, rxOfficeStatement);

                var excel = new ConsolidatedStatementExcel();
                var bytes = excel.GenerateExcel(statement);

                using var saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FileName = statement.Filename;
                saveDialog.AddExtension = true;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    using var stream = saveDialog.OpenFile();

                    stream.Write(bytes);

                    var accountingIssuesUpdater = new AccountingIssuesExcelUpdater();
                    accountingIssuesUpdater.Update([statement]);
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var engine = new BatchProcesingEngine();

            var directory = SelectDirectory();

            if (directory == string.Empty) return;

            engine.Process(directory);
        }

        private static string SelectCSV()
        {
            return SelectFile("CSV files (*.csv)|*.csv|All files (*.*)|*.*");
        }

        private static string SelectARFile()
        {
            return SelectFile("AR files (*.pdf;*.xlsx;*.xls)|*.pdf;*.xlsx;*.xls|All files (*.*)|*.*");
        }

        private static string SelectFile(string filter)
        {
            var fileDialog = new OpenFileDialog();
            fileDialog.Filter = filter;
            fileDialog.RestoreDirectory = true;

            var result = fileDialog.ShowDialog();
            var filename = string.Empty;
            if (result == DialogResult.OK)
            {
                filename = fileDialog.FileName;
            }

            return filename;
        }

        private static string SelectDirectory()
        {
            var browserDialog = new FolderBrowserDialog();
            var result = browserDialog.ShowDialog();

            var directory = string.Empty;

            if (result == DialogResult.OK)
            {
                directory = browserDialog.SelectedPath;
            }

            return directory;
        }
    }
}
