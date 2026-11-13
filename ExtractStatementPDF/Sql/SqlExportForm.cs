namespace ExtractStatementPDF.Sql
{
    // Options for "Export database to SQL file": which tables, and whether to write
    // the structure, the data, DROP TABLE lines and views/routines.
    public class SqlExportForm : Form
    {
        private readonly CheckedListBox lstTables = new() { Dock = DockStyle.Fill, CheckOnClick = true, IntegralHeight = false };
        private readonly Button btnAll = new() { Text = "Select all", AutoSize = true };
        private readonly Button btnNone = new() { Text = "Select none", AutoSize = true };
        private readonly CheckBox chkStructure = new() { Text = "Table structure (CREATE TABLE)", AutoSize = true, Checked = true };
        private readonly CheckBox chkData = new() { Text = "Data (INSERT rows)", AutoSize = true, Checked = true };
        private readonly CheckBox chkDropTable = new() { Text = "Add DROP TABLE before each table", AutoSize = true, Checked = true };
        private readonly CheckBox chkRoutines = new() { Text = "Views, procedures, functions, triggers and events", AutoSize = true, Checked = true };
        private readonly Label lblCount = new() { AutoSize = true, ForeColor = SystemColors.GrayText, Anchor = AnchorStyles.Left };
        private readonly Button btnExport = new() { Text = "Export...", AutoSize = true };
        private readonly Button btnCancel = new() { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };

        public SqlExportOptions? Options { get; private set; }

        public SqlExportForm(string database, IEnumerable<string> tables)
        {
            SuspendLayout();
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;

            Text = $"Export database {database}";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(560, 540);
            MinimumSize = new Size(480, 440);
            ShowInTaskbar = false;
            MinimizeBox = false;
            AcceptButton = btnExport;
            CancelButton = btnCancel;
            Padding = new Padding(10);

            foreach (var table in tables) lstTables.Items.Add(table, true);

            var tableButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            tableButtons.Controls.AddRange([btnAll, btnNone, lblCount]);

            var options = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(0, 6, 0, 0) };
            options.Controls.AddRange([chkStructure, chkData, chkDropTable, chkRoutines]);

            var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 3, Padding = new Padding(0, 6, 0, 0) };
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            buttons.Controls.Add(btnExport, 1, 0);
            buttons.Controls.Add(btnCancel, 2, 0);

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));   // fit the window, not the content
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "Tables to export", AutoSize = true, Margin = new Padding(3, 3, 3, 6) }, 0, 0);
            layout.Controls.Add(lstTables, 0, 1);
            layout.Controls.Add(tableButtons, 0, 2);
            layout.Controls.Add(options, 0, 3);
            layout.Controls.Add(buttons, 0, 4);
            Controls.Add(layout);

            btnAll.Click += (_, _) => SetAll(true);
            btnNone.Click += (_, _) => SetAll(false);
            lstTables.ItemCheck += (_, _) => BeginInvoke(UpdateCount);
            btnExport.Click += (_, _) => Accept();

            ResumeLayout(false);
            PerformLayout();
            UpdateCount();
        }

        private void SetAll(bool check)
        {
            for (var i = 0; i < lstTables.Items.Count; i++) lstTables.SetItemChecked(i, check);
            UpdateCount();
        }

        private void UpdateCount()
        {
            lblCount.Text = $"{lstTables.CheckedItems.Count:N0} of {lstTables.Items.Count:N0} selected";
        }

        private void Accept()
        {
            var selected = lstTables.CheckedItems.Cast<string>().ToList();

            if (lstTables.Items.Count > 0 && selected.Count == 0 && !chkRoutines.Checked)
            {
                MessageBox.Show(this, "Select at least one table.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!chkStructure.Checked && !chkData.Checked && !chkRoutines.Checked)
            {
                MessageBox.Show(this, "Choose what to export (structure, data or routines).", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // No table ticked = routines only (an empty table list would mean "all tables").
            var noTables = selected.Count == 0;

            Options = new SqlExportOptions
            {
                // All tables selected = export everything (also tables created meanwhile).
                Tables = selected.Count == lstTables.Items.Count ? [] : selected,
                Structure = chkStructure.Checked && !noTables,
                Data = chkData.Checked && !noTables,
                DropTable = chkDropTable.Checked,
                Routines = chkRoutines.Checked,
            };

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
