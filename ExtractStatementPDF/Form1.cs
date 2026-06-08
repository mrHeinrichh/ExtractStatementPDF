using ExtractStatementPDF.AR;
using ExtractStatementPDF.Consolidation;
using ExtractStatementPDF.RxOffice;
using System.Reactive.Linq;

namespace ExtractStatementPDF
{
    public partial class Form1 : Form
    {
        private ARStatement? arStatement;

        private RxOfficeStatement? rxOfficeStatement;

        private IDisposable? subscription;

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

            txtUpdate.Text = string.Join(Environment.NewLine, arStatement.Pages);
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

                var statement = new ConsolidatedStatement("", arStatement, rxOfficeStatement);

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

        private async void button4_Click(object sender, EventArgs e)
        {
            try
            {
                var engine = new BatchProcesingEngine();

                var directory = SelectDirectory();
                if (directory == string.Empty) return;

                txtUpdate.AppendText("Processing started...\n");

                subscription = engine.Process(directory)
                    .Subscribe((t) =>
                    {
                        BeginInvoke(() =>
                        {
                            txtUpdate.AppendText($"{t.Message}\n");
                        });
                    },
                    ex =>
                    {
                        BeginInvoke(() =>
                        {
                            txtUpdate.AppendText(
                            $"ERROR: {ex.Message}{Environment.NewLine}");
                        });
                    },
                    () =>
                    {
                        BeginInvoke(() =>
                        {
                            txtUpdate.AppendText("Completed" + Environment.NewLine);
                        });
                    });
            }
            catch
            {
                txtUpdate.AppendText("Processing aborted due to error.");
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            subscription?.Dispose();
            base.OnFormClosed(e);
        }

        private async void button5_Click(object sender, EventArgs e)
        {
            try
            {
                var engine = new BatchProcesingEngine();

                var directory = SelectDirectory();
                if (directory == string.Empty) return;

                txtUpdate.AppendText("Verifying started...");

                await Task.Run(() =>
                {
                    engine.Verify(directory);
                });

                txtUpdate.AppendText("Verifying complete.");
            }
            catch
            {
                txtUpdate.AppendText("Verifying aborted due to error.");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var engine = new BatchProcesingEngine();

            var directory = SelectDirectory();
            if (directory == string.Empty) return;

            engine.VerifyMatches(directory);
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
