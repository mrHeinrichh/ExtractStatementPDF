namespace ExtractStatementPDF.Sql
{
    // Asks which saved connection, database and period "Download Orders CSV" uses,
    // and shows the exact query that will run.
    public class OrdersCsvForm : Form
    {
        private List<SqlSession> sessions = [];

        private readonly ComboBox cmbSession = new() { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly Button btnConnections = new() { Text = "Connections...", AutoSize = true };
        private readonly TextBox txtDatabase = new() { Dock = DockStyle.Fill };
        private readonly DateTimePicker dtFrom = new() { Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Anchor = AnchorStyles.Left };
        private readonly DateTimePicker dtTo = new() { Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Anchor = AnchorStyles.Left };
        private readonly Button btnLastMonth = new() { Text = "Last month", AutoSize = true };
        private readonly Button btnThisMonth = new() { Text = "This month", AutoSize = true };
        private readonly TextBox txtQuery = new()
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            BackColor = SystemColors.Window,
            Font = new Font("Consolas", 9.5f),
        };
        private readonly Button btnDownload = new() { Text = "Download...", AutoSize = true };
        private readonly Button btnCancel = new() { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };

        public SqlSession? Session { get; private set; }

        public string Database => txtDatabase.Text.Trim();

        public DateTime From => dtFrom.Value.Date;

        public DateTime To => dtTo.Value.Date;

        public OrdersCsvForm()
        {
            SuspendLayout();
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;

            Text = "Download Orders CSV";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(640, 470);
            MinimumSize = new Size(560, 420);
            ShowInTaskbar = false;
            MinimizeBox = false;
            AcceptButton = btnDownload;
            CancelButton = btnCancel;
            Padding = new Padding(10);

            var fields = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            void AddRow(string label, Control control, Control? extra = null)
            {
                var row = fields.RowCount++;
                fields.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                fields.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 12, 3) }, 0, row);
                fields.Controls.Add(control, 1, row);
                if (extra != null) fields.Controls.Add(extra, 2, row);
            }

            var monthButtons = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty };
            monthButtons.Controls.AddRange([btnLastMonth, btnThisMonth]);

            AddRow("Connection", cmbSession, btnConnections);
            AddRow("Database", txtDatabase);
            AddRow("From", dtFrom, monthButtons);
            AddRow("To", dtTo);

            var queryLabel = new Label { Text = "Query that will run", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 10, 0, 4) };

            var buttons = new TableLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, ColumnCount = 3, Padding = new Padding(0, 8, 0, 0) };
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            buttons.Controls.Add(btnDownload, 1, 0);
            buttons.Controls.Add(btnCancel, 2, 0);

            // Fill first, then the docked rows, so they take their edges first.
            Controls.Add(txtQuery);
            Controls.Add(queryLabel);
            Controls.Add(fields);
            Controls.Add(buttons);

            cmbSession.SelectedIndexChanged += (_, _) =>
            {
                if (cmbSession.SelectedItem is SqlSession session)
                {
                    txtDatabase.Text = string.IsNullOrWhiteSpace(session.Database) ? OrdersCsvExport.DefaultDatabase : session.Database;
                }
            };
            btnConnections.Click += (_, _) => EditConnections();
            dtFrom.ValueChanged += (_, _) => UpdatePreview();
            dtTo.ValueChanged += (_, _) => UpdatePreview();
            btnLastMonth.Click += (_, _) => SetMonth(DateTime.Today.AddMonths(-1));
            btnThisMonth.Click += (_, _) => SetMonth(DateTime.Today);
            btnDownload.Click += (_, _) => Accept();

            Load += (_, _) =>
            {
                // Fixed-size pickers collapse when the form auto-scales; size them afterwards.
                dtFrom.Width = dtTo.Width = LogicalToDeviceUnits(130);
                ActiveControl = cmbSession;
            };

            ResumeLayout(false);
            PerformLayout();

            LoadSessions(null);
            SetMonth(DateTime.Today.AddMonths(-1));
        }

        private void LoadSessions(string? select)
        {
            var (saved, lastSession) = SqlSessionStore.Load();
            sessions = saved;

            cmbSession.Items.Clear();
            foreach (var session in sessions) cmbSession.Items.Add(session);

            var name = select ?? lastSession;
            var index = sessions.FindIndex(s => s.Name == name);
            if (cmbSession.Items.Count > 0) cmbSession.SelectedIndex = index >= 0 ? index : 0;
        }

        private void EditConnections()
        {
            using var dialog = new SqlSessionForm((cmbSession.SelectedItem as SqlSession)?.Name);
            var chosen = dialog.ShowDialog(this) == DialogResult.OK ? dialog.SelectedSession?.Name : null;

            // The dialog saves sessions itself; reload so new or edited ones show up.
            LoadSessions(chosen ?? (cmbSession.SelectedItem as SqlSession)?.Name);
        }

        private void SetMonth(DateTime month)
        {
            var first = new DateTime(month.Year, month.Month, 1);
            dtFrom.Value = first;
            dtTo.Value = first.AddMonths(1).AddDays(-1);
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            txtQuery.Text = OrdersCsvExport.Preview(From, To).Replace("\n", Environment.NewLine);
        }

        private void Accept()
        {
            string? problem = null;

            if (cmbSession.SelectedItem is not SqlSession) problem = "Choose a connection. Use Connections... to add one (e.g. the RxOffice server 192.168.1.30).";
            else if (From > To) problem = "From must be on or before To.";

            if (problem != null)
            {
                MessageBox.Show(this, problem, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Session = ((SqlSession)cmbSession.SelectedItem!).Clone();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
