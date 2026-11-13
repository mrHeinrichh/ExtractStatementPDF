namespace ExtractStatementPDF.Sql
{
    // Connection settings dialog (HeidiSQL's "Session manager"): saved sessions on the
    // left, the selected session's settings on the right. Kept separate from the query
    // window so configuring a connection never touches the query editor.
    public class SqlSessionForm : Form
    {
        private readonly List<SqlSession> sessions;

        private SqlSession? current;

        private bool loading;

        private readonly ListBox lstSessions = new() { Dock = DockStyle.Fill, IntegralHeight = false };
        private readonly Button btnNew = new() { Text = "New", AutoSize = true };
        private readonly Button btnDelete = new() { Text = "Delete", AutoSize = true };

        private readonly TextBox txtName = new() { Dock = DockStyle.Fill };
        private readonly ComboBox cmbEngine = new() { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox txtHost = new() { Dock = DockStyle.Fill };
        private readonly NumericUpDown numPort = new() { Anchor = AnchorStyles.Left, Minimum = 0, Maximum = 65535 };
        private readonly ComboBox cmbAuth = new() { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox txtUser = new() { Dock = DockStyle.Fill };
        private readonly TextBox txtPassword = new() { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
        private readonly CheckBox chkSavePassword = new() { Text = "Save password (encrypted for your Windows user)", AutoSize = true };
        private readonly TextBox txtDatabase = new() { Dock = DockStyle.Fill, PlaceholderText = "optional, e.g. rxoffice" };
        private readonly CheckBox chkEncrypt = new() { Text = "Encrypt connection", AutoSize = true };
        private readonly CheckBox chkTrust = new() { Text = "Trust server certificate", AutoSize = true };
        private readonly NumericUpDown numTimeout = new() { Anchor = AnchorStyles.Left, Minimum = 1, Maximum = 600 };
        private readonly CheckBox chkCustom = new() { Text = "Edit connection string manually", AutoSize = true };
        private readonly TextBox txtConnectionString = new()
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Height = 64,
            Font = new Font("Consolas", 9f),
        };

        private readonly Label lblInfo = new() { AutoSize = true, ForeColor = SystemColors.GrayText, Anchor = AnchorStyles.Left };
        private readonly Button btnTest = new() { Text = "Test connection", AutoSize = true };
        private readonly Button btnSave = new() { Text = "Save", AutoSize = true };
        private readonly Button btnOpen = new() { Text = "Open", AutoSize = true };
        private readonly Button btnCancel = new() { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };

        public SqlSession? SelectedSession { get; private set; }

        public SqlSessionForm(string? selectName = null)
        {
            var (saved, lastSession) = SqlSessionStore.Load();
            sessions = saved;

            if (sessions.Count == 0)
            {
                // First run: the RxOffice server Accounting/Contact use (password left blank).
                sessions.Add(new SqlSession { Name = "RxOffice", Host = "192.168.1.30", Database = "rxoffice" });
            }

            // Scale with the screen's DPI like the designer-made forms do.
            SuspendLayout();
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;

            Text = "SQL Connection";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(820, 600);
            MinimumSize = new Size(720, 580);
            ShowInTaskbar = false;
            AcceptButton = btnOpen;
            CancelButton = btnCancel;

            BuildLayout();
            WireEvents();

            ResumeLayout(false);
            PerformLayout();

            foreach (var session in sessions) lstSessions.Items.Add(session);

            var select = selectName ?? lastSession;
            var index = sessions.FindIndex(s => s.Name == select);
            lstSessions.SelectedIndex = index >= 0 ? index : 0;
        }

        private void BuildLayout()
        {
            cmbEngine.Items.AddRange([SqlSession.EngineName(SqlEngine.MySql), SqlSession.EngineName(SqlEngine.SqlServer)]);
            cmbAuth.Items.AddRange(["SQL Server authentication", "Windows authentication"]);

            var left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(0, 0, 8, 0) };
            left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            left.Controls.Add(new Label { Text = "Saved sessions", AutoSize = true, Margin = new Padding(3, 3, 3, 6) }, 0, 0);
            left.Controls.Add(lstSessions, 0, 1);
            var leftButtons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
            leftButtons.Controls.AddRange([btnNew, btnDelete]);
            left.Controls.Add(leftButtons, 0, 2);

            var fields = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoScroll = true };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            void AddRow(string label, Control control)
            {
                var row = fields.RowCount++;
                fields.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                fields.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 12, 3) }, 0, row);
                fields.Controls.Add(control, 1, row);
            }

            AddRow("Session name", txtName);
            AddRow("Network type", cmbEngine);
            AddRow("Hostname / IP", txtHost);
            AddRow("Port", numPort);
            AddRow("Authentication", cmbAuth);
            AddRow("User", txtUser);
            AddRow("Password", txtPassword);
            AddRow("", chkSavePassword);
            AddRow("Database", txtDatabase);
            AddRow("", chkEncrypt);
            AddRow("", chkTrust);
            AddRow("Timeout (seconds)", numTimeout);
            AddRow("", chkCustom);
            AddRow("Connection string", txtConnectionString);

            var main = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1 };
            main.Panel1.Controls.Add(left);
            main.Panel2.Controls.Add(fields);

            var bottom = new TableLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, ColumnCount = 5, RowCount = 2, Padding = new Padding(0, 8, 0, 0) };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var i = 0; i < 4; i++) bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.Controls.Add(lblInfo, 0, 0);
            bottom.SetColumnSpan(lblInfo, 5);
            bottom.Controls.Add(btnTest, 1, 1);
            bottom.Controls.Add(btnSave, 2, 1);
            bottom.Controls.Add(btnOpen, 3, 1);
            bottom.Controls.Add(btnCancel, 4, 1);

            Padding = new Padding(10);
            Controls.Add(main);
            Controls.Add(bottom);

            Load += (_, _) =>
            {
                main.SplitterDistance = LogicalToDeviceUnits(200);

                // NumericUpDown collapses when the form auto-scales; size it afterwards.
                numPort.Width = numTimeout.Width = LogicalToDeviceUnits(100);
            };
        }

        private void WireEvents()
        {
            lstSessions.SelectedIndexChanged += (_, _) =>
            {
                current = lstSessions.SelectedItem as SqlSession;
                WriteFields();
            };

            foreach (var box in new[] { txtName, txtHost, txtUser, txtPassword, txtDatabase }) box.TextChanged += (_, _) => ReadFields();
            foreach (var check in new[] { chkSavePassword, chkEncrypt, chkTrust }) check.CheckedChanged += (_, _) => ReadFields();
            numPort.ValueChanged += (_, _) => ReadFields();
            numTimeout.ValueChanged += (_, _) => ReadFields();
            cmbAuth.SelectedIndexChanged += (_, _) => ReadFields();

            cmbEngine.SelectedIndexChanged += (_, _) =>
            {
                if (loading || current == null) return;

                // Follow the engine's default port unless the user typed a custom one.
                var engine = (SqlEngine)cmbEngine.SelectedIndex;
                if ((int)numPort.Value == SqlSession.DefaultPort(current.Engine))
                {
                    loading = true;
                    numPort.Value = SqlSession.DefaultPort(engine);
                    loading = false;
                }

                ReadFields();
            };

            chkCustom.CheckedChanged += (_, _) =>
            {
                if (loading || current == null) return;

                if (chkCustom.Checked && string.IsNullOrWhiteSpace(current.CustomConnectionString))
                {
                    current.CustomConnectionString = current.BuildConnectionString();
                }

                ReadFields();
            };

            txtConnectionString.TextChanged += (_, _) =>
            {
                if (loading || current == null || !current.UseCustomConnectionString) return;
                current.CustomConnectionString = txtConnectionString.Text;
            };

            btnNew.Click += (_, _) =>
            {
                var session = new SqlSession { Name = UniqueName("New session") };
                sessions.Add(session);
                lstSessions.Items.Add(session);
                lstSessions.SelectedItem = session;
                txtName.Focus();
                txtName.SelectAll();
            };

            btnDelete.Click += (_, _) =>
            {
                if (current == null) return;

                var answer = MessageBox.Show(this, $"Delete the session \"{current.Name}\"?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer != DialogResult.Yes) return;

                var index = lstSessions.SelectedIndex;
                sessions.Remove(current);
                lstSessions.Items.Remove(current);

                if (sessions.Count == 0)
                {
                    var session = new SqlSession { Name = UniqueName("New session") };
                    sessions.Add(session);
                    lstSessions.Items.Add(session);
                }

                lstSessions.SelectedIndex = Math.Min(index, lstSessions.Items.Count - 1);
                SaveSessions();
            };

            btnTest.Click += async (_, _) => await TestConnectionAsync();

            btnSave.Click += (_, _) => SaveSessions();

            btnOpen.Click += (_, _) =>
            {
                if (current == null || !ValidateSession(current)) return;

                SaveSessions();
                SelectedSession = current.Clone();
                DialogResult = DialogResult.OK;
                Close();
            };
        }

        private void WriteFields()
        {
            if (current == null) return;

            loading = true;
            txtName.Text = current.Name;
            cmbEngine.SelectedIndex = (int)current.Engine;
            txtHost.Text = current.Host;
            numPort.Value = Math.Clamp(current.Port, 0, 65535);
            cmbAuth.SelectedIndex = (int)current.Authentication;
            txtUser.Text = current.User;
            txtPassword.Text = current.Password;
            chkSavePassword.Checked = current.SavePassword;
            txtDatabase.Text = current.Database;
            chkEncrypt.Checked = current.Encrypt;
            chkTrust.Checked = current.TrustServerCertificate;
            numTimeout.Value = Math.Clamp(current.TimeoutSeconds, 1, 600);
            chkCustom.Checked = current.UseCustomConnectionString;
            loading = false;

            UpdateUi();
        }

        private void ReadFields()
        {
            if (loading || current == null) return;

            var renamed = current.Name != txtName.Text;

            current.Name = txtName.Text;
            current.Engine = (SqlEngine)Math.Max(0, cmbEngine.SelectedIndex);
            current.Host = txtHost.Text;
            current.Port = (int)numPort.Value;
            current.Authentication = (SqlAuthentication)Math.Max(0, cmbAuth.SelectedIndex);
            current.User = txtUser.Text;
            current.Password = txtPassword.Text;
            current.SavePassword = chkSavePassword.Checked;
            current.Database = txtDatabase.Text;
            current.Encrypt = chkEncrypt.Checked;
            current.TrustServerCertificate = chkTrust.Checked;
            current.TimeoutSeconds = (int)numTimeout.Value;
            current.UseCustomConnectionString = chkCustom.Checked;

            if (renamed)
            {
                // Re-assigning the item makes the ListBox repaint the new name.
                var index = lstSessions.SelectedIndex;
                if (index >= 0) lstSessions.Items[index] = current;
            }

            UpdateUi();
        }

        private void UpdateUi()
        {
            if (current == null) return;

            var custom = current.UseCustomConnectionString;
            var sqlServer = current.Engine == SqlEngine.SqlServer;
            var windowsLogin = sqlServer && current.Authentication == SqlAuthentication.Windows;

            foreach (var control in new Control[] { txtHost, numPort, txtDatabase, numTimeout }) control.Enabled = !custom;
            cmbAuth.Enabled = !custom && sqlServer;
            txtUser.Enabled = !custom && !windowsLogin;
            txtPassword.Enabled = !custom && !windowsLogin;
            chkEncrypt.Enabled = !custom && sqlServer;
            chkTrust.Enabled = !custom && sqlServer;

            loading = true;
            txtConnectionString.ReadOnly = !custom;
            txtConnectionString.Text = custom ? current.CustomConnectionString : current.BuildConnectionString(maskPassword: true);
            loading = false;

            lblInfo.Text = sqlServer
                ? "SQL Server: use host,port or a named instance like localhost\\SQLEXPRESS (port 0)."
                : "MySQL / MariaDB: e.g. 192.168.1.30, port 3306, user root.";
        }

        private async Task TestConnectionAsync()
        {
            if (current == null || !ValidateSession(current)) return;

            btnTest.Enabled = false;
            UseWaitCursor = true;

            try
            {
                await using var connection = current.CreateConnection();
                await connection.OpenAsync();

                MessageBox.Show(this,
                    $"Connected to {current.DescribeTarget()}.{Environment.NewLine}Server version: {connection.ServerVersion}",
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    $"Connection failed:{Environment.NewLine}{ex.Message}{Environment.NewLine}{Environment.NewLine}{current.ConnectionHint()}",
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
                btnTest.Enabled = true;
            }
        }

        private bool ValidateSession(SqlSession session)
        {
            string? problem = null;

            if (string.IsNullOrWhiteSpace(session.Name)) problem = "Give the session a name.";
            else if (session.UseCustomConnectionString && string.IsNullOrWhiteSpace(session.CustomConnectionString)) problem = "Enter a connection string.";
            else if (!session.UseCustomConnectionString && string.IsNullOrWhiteSpace(session.Host)) problem = "Enter a hostname or IP address (e.g. localhost).";

            if (problem == null)
            {
                try
                {
                    session.BuildConnectionString();
                }
                catch (Exception ex)
                {
                    problem = $"The connection string is not valid: {ex.Message}";
                }
            }

            if (problem != null) MessageBox.Show(this, problem, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);

            return problem == null;
        }

        private void SaveSessions()
        {
            try
            {
                SqlSessionStore.Save(sessions, current?.Name);
                lblInfo.Text = $"Saved {sessions.Count} session(s).";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not save sessions: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string UniqueName(string name)
        {
            var candidate = name;
            for (var i = 2; sessions.Any(s => s.Name == candidate); i++) candidate = $"{name} {i}";
            return candidate;
        }
    }
}
