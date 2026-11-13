using Microsoft.Data.SqlClient;
using MySqlConnector;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtractStatementPDF.Sql
{
    // A small HeidiSQL-style query window: connection toolbar on top, databases/tables
    // on the left, the query editor (with its own Run/Stop bar) and the results below.
    public class SqlQueryForm : Form
    {
        private const int MaxRowsPerResult = 100_000;

        private const int PreviewRows = 1000;

        private static readonly Regex GoSeparator = new(@"^\s*GO\s*(\d+)?\s*;?\s*$", RegexOptions.IgnoreCase | RegexOptions.Multiline);

        private SqlSession? session;

        private DbConnection? connection;

        private CancellationTokenSource? running;

        private bool updatingDatabaseList;

        // ---- Connection toolbar (configuration) ----
        private readonly ToolStrip toolbar = new() { GripStyle = ToolStripGripStyle.Hidden, Padding = new Padding(6, 2, 6, 2) };
        private readonly ToolStripButton btnConnection = new("Connection...") { ToolTipText = "Choose, edit or create a connection (localhost or IP address)" };
        private readonly ToolStripButton btnDisconnect = new("Disconnect") { Enabled = false };
        private readonly ToolStripComboBox cmbDatabase = new() { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = false, Width = 220, Enabled = false };
        private readonly ToolStripButton btnRefresh = new("Refresh") { Enabled = false, ToolTipText = "Reload databases and tables" };
        private readonly ToolStripDropDownButton btnImportExport = new("Import / Export");
        private readonly ToolStripMenuItem mnuExportDatabase = new("Export database to SQL file...");
        private readonly ToolStripMenuItem mnuImportFile = new("Import SQL file...");
        private readonly ToolStripMenuItem mnuExportCsv = new("Export current result to CSV...");
        private readonly ToolStripMenuItem mnuExportExcel = new("Export current result to Excel...");

        // ---- Left: databases / tables ----
        private readonly TreeView tree = new() { Dock = DockStyle.Fill, HideSelection = false, BorderStyle = BorderStyle.None };
        private readonly ContextMenuStrip treeMenu = new();

        // ---- Query editor and its own Run bar ----
        private readonly TextBox txtQuery = new()
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            AcceptsReturn = true,
            AcceptsTab = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            HideSelection = false,
            MaxLength = 0,
            Font = new Font("Consolas", 10.5f),
        };
        private readonly Button btnRun = new() { Text = "▶  Run  (F9)", AutoSize = true, Enabled = false };
        private readonly Button btnStop = new() { Text = "■  Stop", AutoSize = true, Enabled = false };
        private readonly Button btnClear = new() { Text = "Clear", AutoSize = true };
        private readonly Label lblRunHint = new()
        {
            Text = "Runs the selected text, or the whole editor when nothing is selected.  F9 / F5 / Ctrl+Enter",
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
            Margin = new Padding(12, 8, 3, 3),
        };

        // ---- Output ----
        private readonly TabControl tabs = new() { Dock = DockStyle.Fill };
        private readonly TabPage messagesPage = new("Messages");
        private readonly TextBox txtMessages = new()
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            BackColor = SystemColors.Window,
            Font = new Font("Consolas", 9.5f),
        };

        private readonly StatusStrip status = new();
        private readonly ToolStripStatusLabel lblConnection = new("Not connected") { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
        private readonly ToolStripStatusLabel lblResult = new("");

        private readonly SplitContainer mainSplit = new() { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1 };
        private readonly SplitContainer editorSplit = new() { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal };

        private enum ItemKind { Database, Table, Placeholder }

        private sealed class TreeItem
        {
            public ItemKind Kind { get; init; }

            public string Database { get; init; } = "";

            public string Schema { get; init; } = "";

            public string Name { get; init; } = "";

            public bool IsView { get; init; }

            public bool Loaded { get; set; }
        }

        private sealed record QueryResult(DataTable Table, bool Truncated);

        private sealed class QueryOutcome
        {
            public List<QueryResult> Results { get; } = [];

            public List<string> Messages { get; } = [];

            public bool Failed { get; set; }
        }

        public SqlQueryForm()
        {
            // Scale with the screen's DPI like the designer-made forms do.
            SuspendLayout();
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;

            Text = "SQL Query";
            ClientSize = new Size(1200, 760);
            MinimumSize = new Size(800, 500);
            StartPosition = FormStartPosition.CenterScreen;
            KeyPreview = true;

            BuildLayout();
            WireEvents();

            ResumeLayout(false);
            PerformLayout();
        }

        private void BuildLayout()
        {
            toolbar.Items.AddRange(
            [
                btnConnection,
                btnDisconnect,
                new ToolStripSeparator(),
                new ToolStripLabel("Database:"),
                cmbDatabase,
                btnRefresh,
                new ToolStripSeparator(),
                btnImportExport,
            ]);

            btnImportExport.DropDownItems.AddRange(
            [
                mnuExportDatabase,
                mnuImportFile,
                new ToolStripSeparator(),
                mnuExportCsv,
                mnuExportExcel,
            ]);

            treeMenu.Items.Add($"Show first {PreviewRows:N0} rows", null, async (_, _) => await PreviewSelectedTableAsync());
            treeMenu.Items.Add("Insert name into query", null, (_, _) => InsertSelectedName());
            treeMenu.Items.Add(new ToolStripSeparator());
            treeMenu.Items.Add("Refresh", null, async (_, _) => await LoadDatabasesAsync());
            tree.ContextMenuStrip = treeMenu;

            var runBar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(2, 4, 2, 4) };
            runBar.Controls.AddRange([btnRun, btnStop, btnClear, lblRunHint]);

            editorSplit.Panel1.Controls.Add(txtQuery);
            editorSplit.Panel1.Controls.Add(runBar);

            messagesPage.Controls.Add(txtMessages);
            tabs.TabPages.Add(messagesPage);
            editorSplit.Panel2.Controls.Add(tabs);

            var treeHeader = new Label { Text = "Databases", Dock = DockStyle.Top, Height = 24, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };
            mainSplit.Panel1.Controls.Add(tree);
            mainSplit.Panel1.Controls.Add(treeHeader);
            mainSplit.Panel2.Controls.Add(editorSplit);

            status.Items.AddRange([lblConnection, lblResult]);

            // Fill first, then the docked bars, so the bars take their edges first.
            Controls.Add(mainSplit);
            Controls.Add(toolbar);
            Controls.Add(status);
        }

        private void WireEvents()
        {
            btnConnection.Click += async (_, _) => await ChooseConnectionAsync();
            btnDisconnect.Click += async (_, _) => await DisconnectAsync();
            btnRefresh.Click += async (_, _) => await LoadDatabasesAsync();
            cmbDatabase.SelectedIndexChanged += async (_, _) => await ChangeDatabaseAsync();

            btnImportExport.DropDownOpening += (_, _) =>
            {
                var connected = connection != null && session != null;
                var mySql = session?.Engine == SqlEngine.MySql;

                mnuExportDatabase.Enabled = connected && mySql;
                mnuExportDatabase.Text = connected && !mySql
                    ? "Export database to SQL file... (MySQL / MariaDB only)"
                    : "Export database to SQL file...";
                mnuImportFile.Enabled = connected;
                mnuExportCsv.Enabled = mnuExportExcel.Enabled = CurrentResultTable() != null;
            };
            mnuExportDatabase.Click += async (_, _) => await ExportDatabaseAsync();
            mnuImportFile.Click += async (_, _) => await ImportFileAsync();
            mnuExportCsv.Click += async (_, _) => await ExportResultAsync(excel: false);
            mnuExportExcel.Click += async (_, _) => await ExportResultAsync(excel: true);

            btnRun.Click += async (_, _) => await RunEditorAsync();
            btnStop.Click += (_, _) => running?.Cancel();
            btnClear.Click += (_, _) => { txtQuery.Clear(); txtQuery.Focus(); };

            tree.BeforeExpand += async (_, e) => await LoadTablesAsync(e.Node);
            tree.AfterSelect += (_, e) =>
            {
                if (e.Node?.Tag is TreeItem { Kind: ItemKind.Database } item) SelectDatabaseInList(item.Database, apply: true);
            };
            tree.NodeMouseClick += (_, e) =>
            {
                if (e.Button == MouseButtons.Right) tree.SelectedNode = e.Node;
            };
            tree.NodeMouseDoubleClick += async (_, e) =>
            {
                if (e.Node?.Tag is TreeItem { Kind: ItemKind.Table }) await PreviewSelectedTableAsync();
            };
            treeMenu.Opening += (_, e) =>
            {
                var isTable = tree.SelectedNode?.Tag is TreeItem { Kind: ItemKind.Table };
                treeMenu.Items[0].Enabled = isTable && connection != null && running == null;
                treeMenu.Items[1].Enabled = tree.SelectedNode?.Tag is TreeItem { Kind: not ItemKind.Placeholder };
                treeMenu.Items[3].Enabled = session != null;
            };

            Load += (_, _) =>
            {
                mainSplit.SplitterDistance = LogicalToDeviceUnits(250);
                editorSplit.SplitterDistance = (int)(editorSplit.Height * 0.42);
                txtQuery.Text = SqlSessionStore.LoadQuery();
            };

            // Like HeidiSQL, ask for a connection as soon as the window opens.
            Shown += async (_, _) => await ChooseConnectionAsync();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            var run = e.KeyCode is Keys.F9 or Keys.F5 || (e.Control && e.KeyCode == Keys.Enter);

            if (run)
            {
                e.Handled = e.SuppressKeyPress = true;
                _ = RunEditorAsync();
                return;
            }

            if (e.Control && e.KeyCode == Keys.A && ActiveControl is TextBox box)
            {
                e.Handled = e.SuppressKeyPress = true;
                box.SelectAll();
                return;
            }

            base.OnKeyDown(e);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SqlSessionStore.SaveQuery(txtQuery.Text);
            running?.Cancel();
            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            connection?.Dispose();
            connection = null;
            base.OnFormClosed(e);
        }

        // ---------------------------------------------------------------- connection

        private async Task ChooseConnectionAsync()
        {
            if (running != null) return;

            using var dialog = new SqlSessionForm(session?.Name);
            if (dialog.ShowDialog(this) != DialogResult.OK || dialog.SelectedSession == null) return;

            await ConnectAsync(dialog.SelectedSession);
        }

        private async Task ConnectAsync(SqlSession target)
        {
            await DisconnectAsync();

            lblConnection.Text = $"Connecting to {target.DescribeTarget()}...";
            UseWaitCursor = true;

            DbConnection? opened = null;

            try
            {
                opened = target.CreateConnection();
                await opened.OpenAsync();
            }
            catch (Exception ex)
            {
                opened?.Dispose();
                UseWaitCursor = false;
                lblConnection.Text = "Not connected";
                MessageBox.Show(this,
                    $"Could not connect to {target.DescribeTarget()}:{Environment.NewLine}{ex.Message}{Environment.NewLine}{Environment.NewLine}{target.ConnectionHint()}",
                    "SQL Connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            UseWaitCursor = false;

            session = target;
            connection = opened;

            Text = $"SQL Query - {target.Name}";
            lblConnection.Text = $"Connected: {target.Name}  ·  {target.DescribeTarget()}  ·  {SqlSession.EngineName(target.Engine)} {opened.ServerVersion}";
            AppendMessage($"Connected to {target.DescribeTarget()} ({SqlSession.EngineName(target.Engine)} {opened.ServerVersion}).");

            UpdateButtons();
            await LoadDatabasesAsync();
        }

        private async Task DisconnectAsync()
        {
            running?.Cancel();

            if (connection != null)
            {
                await connection.DisposeAsync();
                connection = null;
                AppendMessage("Disconnected.");
            }

            session = null;
            tree.Nodes.Clear();

            updatingDatabaseList = true;
            cmbDatabase.Items.Clear();
            updatingDatabaseList = false;

            Text = "SQL Query";
            lblConnection.Text = "Not connected";
            UpdateButtons();
        }

        private async Task<bool> EnsureOpenAsync()
        {
            if (connection == null || session == null) return false;
            if (connection.State == ConnectionState.Open) return true;

            // The server dropped us (timeout, restart, network); reconnect quietly once.
            try
            {
                var database = connection.Database;
                connection.Dispose();
                connection = session.CreateConnection(string.IsNullOrEmpty(database) ? null : database);
                await connection.OpenAsync();
                AppendMessage("Reconnected.");
                return true;
            }
            catch (Exception ex)
            {
                AppendMessage($"Reconnect failed: {ex.Message}");
                return false;
            }
        }

        private void UpdateButtons()
        {
            var connected = connection != null;
            var busy = running != null;

            btnConnection.Enabled = !busy;
            btnDisconnect.Enabled = connected && !busy;
            btnRefresh.Enabled = connected && !busy;
            cmbDatabase.Enabled = connected && !busy;
            btnRun.Enabled = connected && !busy;
            btnStop.Enabled = busy;
            btnImportExport.Enabled = !busy;
        }

        // ---------------------------------------------------------------- databases / tables

        private async Task LoadDatabasesAsync()
        {
            if (session == null || connection == null) return;

            tree.BeginUpdate();
            tree.Nodes.Clear();
            tree.EndUpdate();

            List<string> databases;

            try
            {
                var sql = session.Engine == SqlEngine.MySql
                    ? "SHOW DATABASES"
                    : "SELECT name FROM sys.databases WHERE HAS_DBACCESS(name) = 1 ORDER BY name";

                databases = (await QueryMetadataAsync(sql, null)).Select(r => r[0]).ToList();
            }
            catch (Exception ex)
            {
                AppendMessage($"Could not list databases: {ex.Message}");
                return;
            }

            tree.BeginUpdate();
            foreach (var database in databases)
            {
                var node = new TreeNode(database) { Tag = new TreeItem { Kind = ItemKind.Database, Database = database } };
                node.Nodes.Add(PlaceholderNode("Loading..."));
                tree.Nodes.Add(node);
            }
            tree.EndUpdate();

            updatingDatabaseList = true;
            cmbDatabase.Items.Clear();
            foreach (var database in databases) cmbDatabase.Items.Add(database);
            updatingDatabaseList = false;

            SelectDatabaseInList(connection.Database, apply: false);
        }

        private async Task LoadTablesAsync(TreeNode? node)
        {
            if (node?.Tag is not TreeItem { Kind: ItemKind.Database, Loaded: false } item || session == null) return;

            item.Loaded = true;

            try
            {
                var rows = session.Engine == SqlEngine.MySql
                    ? await QueryMetadataAsync(
                        "SELECT TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE FROM information_schema.TABLES WHERE TABLE_SCHEMA = @db ORDER BY TABLE_NAME",
                        null, ("@db", item.Database))
                    : await QueryMetadataAsync(
                        "SELECT TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_SCHEMA, TABLE_NAME",
                        item.Database);

                tree.BeginUpdate();
                node.Nodes.Clear();

                foreach (var row in rows)
                {
                    var table = new TreeItem
                    {
                        Kind = ItemKind.Table,
                        Database = item.Database,
                        Schema = row[0],
                        Name = row[1],
                        IsView = row[2].Contains("VIEW", StringComparison.OrdinalIgnoreCase),
                    };

                    var text = session.Engine == SqlEngine.SqlServer ? $"{table.Schema}.{table.Name}" : table.Name;
                    if (table.IsView) text += "  (view)";

                    node.Nodes.Add(new TreeNode(text) { Tag = table, ForeColor = table.IsView ? Color.DarkCyan : tree.ForeColor });
                }

                if (rows.Count == 0) node.Nodes.Add(PlaceholderNode("(no tables)"));
                tree.EndUpdate();
            }
            catch (Exception ex)
            {
                item.Loaded = false;
                node.Nodes.Clear();
                node.Nodes.Add(PlaceholderNode($"(error: {ex.Message})"));
            }
        }

        // Metadata runs on its own short-lived (pooled) connection so browsing the tree
        // never collides with a query running on the main connection.
        private async Task<List<string[]>> QueryMetadataAsync(string sql, string? database, params (string Name, object Value)[] parameters)
        {
            await using var metadata = session!.CreateConnection(database);
            await metadata.OpenAsync();

            await using var command = metadata.CreateCommand();
            command.CommandText = sql;

            foreach (var (name, value) in parameters)
            {
                var parameter = command.CreateParameter();
                parameter.ParameterName = name;
                parameter.Value = value;
                command.Parameters.Add(parameter);
            }

            var rows = new List<string[]>();
            await using var reader = await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                var row = new string[reader.FieldCount];
                for (var i = 0; i < row.Length; i++) row[i] = reader.IsDBNull(i) ? "" : Convert.ToString(reader.GetValue(i)) ?? "";
                rows.Add(row);
            }

            return rows;
        }

        private async Task ChangeDatabaseAsync()
        {
            if (updatingDatabaseList || connection == null || cmbDatabase.SelectedItem is not string database) return;

            if (running != null)
            {
                SelectDatabaseInList(connection.Database, apply: false);
                return;
            }

            if (database == connection.Database || !await EnsureOpenAsync()) return;

            try
            {
                await connection.ChangeDatabaseAsync(database);
                AppendMessage($"Using database {database}.");
            }
            catch (Exception ex)
            {
                AppendMessage($"Could not switch to {database}: {ex.Message}");
                SelectDatabaseInList(connection.Database, apply: false);
            }
        }

        private void SelectDatabaseInList(string? database, bool apply)
        {
            if (string.IsNullOrEmpty(database) || !cmbDatabase.Items.Contains(database)) return;

            updatingDatabaseList = !apply;
            cmbDatabase.SelectedItem = database;
            updatingDatabaseList = false;
        }

        private async Task PreviewSelectedTableAsync()
        {
            if (tree.SelectedNode?.Tag is not TreeItem { Kind: ItemKind.Table } table || session == null) return;

            var sql = session.Engine == SqlEngine.MySql
                ? $"SELECT * FROM {QuoteMySql(table.Database)}.{QuoteMySql(table.Name)} LIMIT {PreviewRows};"
                : $"SELECT TOP ({PreviewRows}) * FROM {QuoteSqlServer(table.Database)}.{QuoteSqlServer(table.Schema)}.{QuoteSqlServer(table.Name)};";

            await RunAsync(sql, table.Name);
        }

        private void InsertSelectedName()
        {
            if (tree.SelectedNode?.Tag is not TreeItem item || session == null) return;

            var name = item.Kind switch
            {
                ItemKind.Database => session.Engine == SqlEngine.MySql ? QuoteMySql(item.Database) : QuoteSqlServer(item.Database),
                ItemKind.Table when session.Engine == SqlEngine.MySql => $"{QuoteMySql(item.Database)}.{QuoteMySql(item.Name)}",
                ItemKind.Table => $"{QuoteSqlServer(item.Database)}.{QuoteSqlServer(item.Schema)}.{QuoteSqlServer(item.Name)}",
                _ => "",
            };

            txtQuery.SelectedText = name;
            txtQuery.Focus();
        }

        private static TreeNode PlaceholderNode(string text) =>
            new(text) { Tag = new TreeItem { Kind = ItemKind.Placeholder }, ForeColor = SystemColors.GrayText };

        private static string QuoteMySql(string name) => $"`{name.Replace("`", "``")}`";

        private static string QuoteSqlServer(string name) => $"[{name.Replace("]", "]]")}]";

        // ---------------------------------------------------------------- running queries

        private async Task RunEditorAsync()
        {
            var sql = txtQuery.SelectionLength > 0 ? txtQuery.SelectedText : txtQuery.Text;
            await RunAsync(sql, null);
        }

        private async Task RunAsync(string sql, string? title)
        {
            if (running != null) return;

            if (connection == null || session == null)
            {
                await ChooseConnectionAsync();
                if (connection == null) return;
            }

            var batches = SplitBatches(sql, session!.Engine);
            if (batches.Count == 0) return;

            running = new CancellationTokenSource();
            UpdateButtons();
            lblResult.Text = "Running...";

            var outcome = new QueryOutcome();
            var stopwatch = Stopwatch.StartNew();

            try
            {
                if (!await EnsureOpenAsync())
                {
                    outcome.Failed = true;
                    outcome.Messages.Add("Not connected. Use Connection... to connect again.");
                }
                else
                {
                    var target = connection!;
                    var engine = session.Engine;
                    var token = running.Token;

                    await Task.Run(() => ExecuteAsync(target, engine, batches, outcome, token));
                }
            }
            finally
            {
                stopwatch.Stop();
                running.Dispose();
                running = null;
            }

            if (IsDisposed) return;

            ShowOutcome(outcome, title, stopwatch.Elapsed);
            UpdateButtons();

            // A "USE db" in the query changes the session's database; reflect it.
            if (connection?.State == ConnectionState.Open) SelectDatabaseInList(connection.Database, apply: false);
        }

        private static async Task ExecuteAsync(DbConnection connection, SqlEngine engine, List<(string Sql, int Line)> batches, QueryOutcome outcome, CancellationToken token)
        {
            void Message(string text)
            {
                lock (outcome.Messages) outcome.Messages.Add(text);
            }

            // Server messages (PRINT, warnings) and per-statement row counts.
            SqlInfoMessageEventHandler? sqlInfo = null;
            MySqlInfoMessageEventHandler? mySqlInfo = null;

            if (connection is SqlConnection sqlConnection)
            {
                sqlInfo = (_, e) => Message(e.Message);
                sqlConnection.InfoMessage += sqlInfo;
            }
            else if (connection is MySqlConnection mySqlConnection)
            {
                mySqlInfo = (_, e) => { foreach (var error in e.Errors) Message($"{error.Level}: {error.Message}"); };
                mySqlConnection.InfoMessage += mySqlInfo;
            }

            try
            {
                foreach (var (sql, line) in batches)
                {
                    token.ThrowIfCancellationRequested();

                    await using var command = connection.CreateCommand();
                    command.CommandText = sql;
                    command.CommandTimeout = 0;   // run until finished or Stop is pressed

                    if (command is SqlCommand sqlCommand)
                    {
                        sqlCommand.StatementCompleted += (_, e) => Message($"({e.RecordCount:N0} row(s) affected)");
                    }

                    try
                    {
                        await using var reader = await command.ExecuteReaderAsync(token);

                        do
                        {
                            if (reader.FieldCount == 0) continue;

                            var result = await ReadTableAsync(reader, token);
                            outcome.Results.Add(result);

                            Message($"Result {outcome.Results.Count}: {result.Table.Rows.Count:N0} row(s)"
                                + (result.Truncated ? $" (showing the first {MaxRowsPerResult:N0})" : ""));
                        }
                        while (await reader.NextResultAsync(token));

                        await reader.CloseAsync();

                        if (engine == SqlEngine.MySql && reader.RecordsAffected >= 0)
                        {
                            Message($"({reader.RecordsAffected:N0} row(s) affected)");
                        }
                    }
                    catch (Exception) when (token.IsCancellationRequested)
                    {
                        throw new OperationCanceledException(token);
                    }
                    catch (SqlException ex)
                    {
                        Message($"Error (line {line + ex.LineNumber - 1}): {ex.Message}");
                        outcome.Failed = true;
                        return;
                    }
                    catch (DbException ex)
                    {
                        Message($"Error: {ex.Message}");
                        outcome.Failed = true;
                        return;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Message("Query cancelled.");
                outcome.Failed = true;
            }
            catch (Exception ex)
            {
                Message($"Error: {ex.Message}");
                outcome.Failed = true;
            }
            finally
            {
                if (connection is SqlConnection sc && sqlInfo != null) sc.InfoMessage -= sqlInfo;
                if (connection is MySqlConnection mc && mySqlInfo != null) mc.InfoMessage -= mySqlInfo;
            }
        }

        private static async Task<QueryResult> ReadTableAsync(DbDataReader reader, CancellationToken token)
        {
            var table = new DataTable();
            var binary = new bool[reader.FieldCount];
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (var i = 0; i < reader.FieldCount; i++)
            {
                var name = reader.GetName(i);
                if (string.IsNullOrWhiteSpace(name)) name = $"(No column name {i + 1})";

                var unique = name;
                for (var n = 2; !names.Add(unique); n++) unique = $"{name}_{n}";

                Type type;
                try
                {
                    type = reader.GetFieldType(i) ?? typeof(object);
                }
                catch
                {
                    type = typeof(object);
                }

                if (type == typeof(byte[]))
                {
                    binary[i] = true;
                    type = typeof(string);
                }

                table.Columns.Add(unique, type);
            }

            var values = new object[reader.FieldCount];
            var truncated = false;

            table.BeginLoadData();

            while (await reader.ReadAsync(token))
            {
                if (table.Rows.Count >= MaxRowsPerResult)
                {
                    truncated = true;
                    break;
                }

                for (var i = 0; i < values.Length; i++) values[i] = ReadValue(reader, i, binary[i]);

                table.Rows.Add(values);
            }

            table.EndLoadData();

            return new QueryResult(table, truncated);
        }

        private static object ReadValue(DbDataReader reader, int ordinal, bool binary)
        {
            try
            {
                if (reader.IsDBNull(ordinal)) return DBNull.Value;

                var value = reader.GetValue(ordinal);
                return binary && value is byte[] bytes ? ToHex(bytes) : value;
            }
            catch
            {
                // Values the provider cannot convert (e.g. an invalid date) show as NULL.
                return DBNull.Value;
            }
        }

        private static string ToHex(byte[] bytes)
        {
            const int max = 64;
            var hex = Convert.ToHexString(bytes, 0, Math.Min(bytes.Length, max));
            return bytes.Length > max ? $"0x{hex}... ({bytes.Length:N0} bytes)" : $"0x{hex}";
        }

        // SQL Server scripts use "GO" lines between batches; MySQL takes the whole text
        // (several statements separated by ';' run in one go). Each batch keeps its first
        // line number so SQL Server errors point at the right editor line.
        private static List<(string Sql, int Line)> SplitBatches(string sql, SqlEngine engine)
        {
            var batches = new List<(string, int)>();

            void Add(string text, int line)
            {
                if (!string.IsNullOrWhiteSpace(text)) batches.Add((text, line));
            }

            if (engine == SqlEngine.MySql)
            {
                Add(sql, 1);
                return batches;
            }

            var start = 0;
            foreach (Match go in GoSeparator.Matches(sql))
            {
                Add(sql[start..go.Index], LineOf(sql, start));
                start = go.Index + go.Length;
            }

            Add(sql[start..], LineOf(sql, start));
            return batches;
        }

        private static int LineOf(string text, int index) => 1 + text.AsSpan(0, Math.Min(index, text.Length)).Count('\n');

        // ---------------------------------------------------------------- import / export

        private async Task ExportDatabaseAsync()
        {
            if (session == null || connection == null || running != null) return;

            if (cmbDatabase.SelectedItem is not string database)
            {
                MessageBox.Show(this, "Choose the database to export in the Database list at the top first.", "Export database",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            List<string> tables;

            try
            {
                tables = (await QueryMetadataAsync(
                    "SELECT TABLE_NAME FROM information_schema.TABLES WHERE TABLE_SCHEMA = @db AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME",
                    null, ("@db", database))).Select(r => r[0]).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not list the tables of {database}:{Environment.NewLine}{ex.Message}", "Export database",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var optionsForm = new SqlExportForm(database, tables);
            if (optionsForm.ShowDialog(this) != DialogResult.OK || optionsForm.Options == null) return;

            using var save = new SaveFileDialog
            {
                Filter = "SQL file (*.sql)|*.sql|All files (*.*)|*.*",
                FileName = $"{database}_{DateTime.Now:yyyyMMdd_HHmm}.sql",
            };
            if (save.ShowDialog(this) != DialogResult.OK) return;

            var path = save.FileName;
            var target = session;
            var options = optionsForm.Options;

            await RunToolAsync(
                $"Exporting database {database} to {path}...",
                token => SqlImportExport.ExportDatabase(target, database, path, options, ReportProgress, token),
                () => $"Exported database {database} to {path} ({new FileInfo(path).Length / 1024.0 / 1024.0:N1} MB)",
                () => DeleteQuietly(path));
        }

        private async Task ImportFileAsync()
        {
            if (session == null || connection == null || running != null) return;

            using var open = new OpenFileDialog { Filter = "SQL files (*.sql)|*.sql|All files (*.*)|*.*" };
            if (open.ShowDialog(this) != DialogResult.OK) return;

            var path = open.FileName;
            var name = Path.GetFileName(path);
            var megabytes = new FileInfo(path).Length / 1024.0 / 1024.0;
            var database = cmbDatabase.SelectedItem as string ?? connection.Database;

            if (string.IsNullOrEmpty(database))
            {
                MessageBox.Show(this, "Choose the database to import into in the Database list at the top first.", "Import SQL file",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var answer = MessageBox.Show(this,
                $"Import {name} ({megabytes:N1} MB) into database \"{database}\" on {session.DescribeTarget()}?{Environment.NewLine}{Environment.NewLine}" +
                "Every statement in the file will run. SQL dumps usually DROP and re-create their tables, " +
                "which replaces the data in those tables. This cannot be undone.",
                "Import SQL file", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes) return;

            if (session.Engine == SqlEngine.MySql)
            {
                var target = session;
                var imported = await RunToolAsync(
                    $"Importing {name} into {database}...",
                    token => SqlImportExport.ImportMySqlFile(target, database, path, ReportProgress, token),
                    () => $"Imported {name} into {database}");

                if (imported) await LoadDatabasesAsync();
            }
            else
            {
                // SQL Server scripts run through the normal query runner (GO batches, messages).
                if (megabytes > 200)
                {
                    MessageBox.Show(this, $"{name} is {megabytes:N0} MB, too large to run here. Use sqlcmd or SSMS for files this size.",
                        "Import SQL file", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                await RunAsync(await File.ReadAllTextAsync(path), name);
                await LoadDatabasesAsync();
            }
        }

        private async Task ExportResultAsync(bool excel)
        {
            var table = CurrentResultTable();
            if (table == null || tabs.SelectedTab == null) return;

            // Tab text is "<name> (<rows>)"; keep the name for the file and sheet.
            var name = Regex.Replace(tabs.SelectedTab.Text, @"\s*\([\d,+]+\)$", "");
            var safeName = string.Concat(name.Split(Path.GetInvalidFileNameChars()));

            using var save = new SaveFileDialog
            {
                Filter = excel ? "Excel workbook (*.xlsx)|*.xlsx" : "CSV file (*.csv)|*.csv",
                FileName = $"{safeName}_{DateTime.Now:yyyyMMdd_HHmm}{(excel ? ".xlsx" : ".csv")}",
            };
            if (save.ShowDialog(this) != DialogResult.OK) return;

            var path = save.FileName;
            var rows = table.DefaultView.ToTable();   // keeps the grid's current sort order

            UseWaitCursor = true;
            lblResult.Text = $"Saving {Path.GetFileName(path)}...";

            try
            {
                await Task.Run(() =>
                {
                    if (excel) SqlImportExport.ExportExcel(rows, path, name);
                    else SqlImportExport.ExportCsv(rows, path);
                });

                AppendMessage($"-- {DateTime.Now:HH:mm:ss}  Saved {rows.Rows.Count:N0} row(s) from \"{name}\" to {path}");
                lblResult.Text = $"Saved {rows.Rows.Count:N0} row(s) to {Path.GetFileName(path)}";
            }
            catch (Exception ex)
            {
                lblResult.Text = "Save failed";
                MessageBox.Show(this, $"Could not save {path}:{Environment.NewLine}{ex.Message}", "Export result",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        // Runs a long blocking job (export/import) off the UI thread with the Stop
        // button enabled; progress goes to the status bar, the outcome to Messages.
        private async Task<bool> RunToolAsync(string starting, Action<CancellationToken> work, Func<string> succeeded, Action? cancelled = null)
        {
            running = new CancellationTokenSource();
            UpdateButtons();
            lblResult.Text = starting;
            AppendMessage($"-- {DateTime.Now:HH:mm:ss}  {starting}");

            var token = running.Token;
            var stopwatch = Stopwatch.StartNew();
            var ok = false;

            try
            {
                await Task.Run(() => work(token));
                ok = true;
                AppendMessage($"{succeeded()} in {stopwatch.Elapsed.TotalSeconds:0.0} s.");
                lblResult.Text = $"Done in {stopwatch.Elapsed.TotalSeconds:0.0} s";
            }
            catch (Exception ex)
            {
                if (token.IsCancellationRequested)
                {
                    cancelled?.Invoke();
                    AppendMessage("Cancelled.");
                    lblResult.Text = "Cancelled";
                }
                else
                {
                    AppendMessage($"Error: {ex.Message}");
                    lblResult.Text = "Error";
                }
            }
            finally
            {
                running.Dispose();
                running = null;
                if (!IsDisposed) UpdateButtons();
            }

            if (!IsDisposed) tabs.SelectedTab = messagesPage;
            return ok;
        }

        private void ReportProgress(string text)
        {
            try
            {
                if (!IsDisposed && IsHandleCreated) BeginInvoke(() => lblResult.Text = text);
            }
            catch (InvalidOperationException)
            {
                // The window closed while the job was finishing.
            }
        }

        private DataTable? CurrentResultTable() =>
            tabs.SelectedTab?.Controls.OfType<DataGridView>().FirstOrDefault()?.DataSource as DataTable;

        private static void DeleteQuietly(string path)
        {
            try
            {
                File.Delete(path);
            }
            catch
            {
                // A partial file left behind is harmless; the message says it was cancelled.
            }
        }

        // ---------------------------------------------------------------- output

        private void ShowOutcome(QueryOutcome outcome, string? title, TimeSpan elapsed)
        {
            tabs.SuspendLayout();

            foreach (var page in tabs.TabPages.Cast<TabPage>().Where(p => p != messagesPage).ToList())
            {
                tabs.TabPages.Remove(page);
                page.Dispose();
            }

            TabPage? first = null;

            for (var i = 0; i < outcome.Results.Count; i++)
            {
                var result = outcome.Results[i];
                var name = title != null && outcome.Results.Count == 1 ? title : $"Result {i + 1}";
                var page = new TabPage($"{name} ({result.Table.Rows.Count:N0}{(result.Truncated ? "+" : "")})");

                var grid = CreateGrid();
                page.Controls.Add(grid);
                tabs.TabPages.Insert(tabs.TabPages.IndexOf(messagesPage), page);
                grid.DataSource = result.Table;

                first ??= page;
            }

            var header = $"-- {DateTime.Now:HH:mm:ss}  {(outcome.Failed ? "finished with errors" : "done")} in {elapsed.TotalSeconds:0.000} s";
            AppendMessage(header + Environment.NewLine + string.Join(Environment.NewLine, outcome.Messages));

            tabs.SelectedTab = outcome.Failed || first == null ? messagesPage : first;
            tabs.ResumeLayout();

            var rows = outcome.Results.Sum(r => r.Table.Rows.Count);
            lblResult.Text = outcome.Failed
                ? $"Error  ·  {elapsed.TotalSeconds:0.000} s"
                : $"{outcome.Results.Count} result(s)  ·  {rows:N0} row(s)  ·  {elapsed.TotalSeconds:0.000} s";
        }

        private static DataGridView CreateGrid()
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                BackgroundColor = SystemColors.Window,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            };

            grid.DefaultCellStyle.NullValue = "(NULL)";

            // Smoother scrolling over large results.
            typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(grid, true);

            grid.DataError += (_, e) => e.ThrowException = false;

            grid.CellFormatting += (_, e) =>
            {
                if (e.Value is DBNull or null && e.CellStyle != null) e.CellStyle.ForeColor = SystemColors.GrayText;
            };

            grid.DataBindingComplete += (_, _) =>
            {
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    if (column.ValueType == typeof(DateTime)) column.DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";
                    column.SortMode = DataGridViewColumnSortMode.Automatic;
                }

                grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

                foreach (DataGridViewColumn column in grid.Columns)
                {
                    column.Width = Math.Clamp(column.Width, 60, 360);
                }
            };

            return grid;
        }

        private void AppendMessage(string text)
        {
            var builder = new StringBuilder();
            if (txtMessages.TextLength > 0) builder.AppendLine();
            builder.Append(text.Replace("\r\n", "\n").Replace("\n", Environment.NewLine));

            txtMessages.AppendText(builder.ToString());
        }
    }
}
