using ExtractStatementPDF.Consolidation;
using System.Diagnostics;

namespace ExtractStatementPDF
{
    public sealed class SanitizeResultDialog : Form
    {
        public SanitizeResultDialog(SanitizeResult result)
        {
            Text = "Sanitize Result";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(760, 560);
            Size = new Size(860, 640);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 7
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 65));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var summaryLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Font = new Font(Font, FontStyle.Bold),
                Text = BuildSummary(result)
            };

            var scannedFolderLink = CreatePathLinkLabel(
                $"Scanned folder: {result.ScannedDirectory}",
                result.ScannedDirectory,
                isDirectory: true);

            var archiveHeader = new Label
            {
                AutoSize = true,
                Text = "Archive folder(s): Click a path to open it."
            };

            var archivePanel = new FlowLayoutPanel
            {
                AutoScroll = true,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 4)
            };

            if (result.ArchiveDirectories.Count == 0)
            {
                archivePanel.Controls.Add(new Label
                {
                    AutoSize = true,
                    Text = "No archive folders were created."
                });
            }
            else
            {
                foreach (var archiveDirectory in result.ArchiveDirectories)
                {
                    archivePanel.Controls.Add(CreatePathLinkLabel(archiveDirectory, archiveDirectory, isDirectory: true));
                }
            }

            var movedFilesHeader = new Label
            {
                AutoSize = true,
                Text = "Moved file destination(s):"
            };

            var movedFilesTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Text = result.MovedFiles.Count == 0
                    ? "No files were moved."
                    : string.Join(Environment.NewLine, result.MovedFiles)
            };

            var closeButton = new Button
            {
                AutoSize = true,
                Text = "Close",
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Right
            };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true
            };
            buttonPanel.Controls.Add(closeButton);

            layout.Controls.Add(summaryLabel, 0, 0);
            layout.Controls.Add(scannedFolderLink, 0, 1);
            layout.Controls.Add(archiveHeader, 0, 2);
            layout.Controls.Add(archivePanel, 0, 3);
            layout.Controls.Add(movedFilesHeader, 0, 4);
            layout.Controls.Add(movedFilesTextBox, 0, 5);
            layout.Controls.Add(buttonPanel, 0, 6);

            Controls.Add(layout);
            AcceptButton = closeButton;
        }

        private static string BuildSummary(SanitizeResult result)
        {
            if (result.MovedFileCount == 0)
            {
                return "Sanitize completed. No manual-order matches were archived.";
            }

            return $"Sanitize completed. Archived {result.MovedFileCount} file(s) from {result.ArchivedMatchCount} matched set(s).";
        }

        private static LinkLabel CreatePathLinkLabel(string labelText, string targetPath, bool isDirectory)
        {
            var linkLabel = new LinkLabel
            {
                AutoSize = true,
                MaximumSize = new Size(780, 0),
                Text = labelText,
                Tag = (targetPath, isDirectory)
            };

            linkLabel.Links.Add(0, labelText.Length, linkLabel.Tag);
            linkLabel.LinkClicked += (_, _) => OpenPath(targetPath, isDirectory);

            return linkLabel;
        }

        private static void OpenPath(string targetPath, bool isDirectory)
        {
            if (isDirectory)
            {
                if (!Directory.Exists(targetPath))
                {
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = targetPath,
                    UseShellExecute = true
                });
                return;
            }

            if (!File.Exists(targetPath))
            {
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{targetPath}\"",
                UseShellExecute = true
            });
        }
    }
}
