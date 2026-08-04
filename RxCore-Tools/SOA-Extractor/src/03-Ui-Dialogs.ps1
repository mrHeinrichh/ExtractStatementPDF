# ---- Simple GUI dialogs (message box + file browser) ----------------------

# Show-Message : pop a message box (used for the login warning).
function Show-Message([string]$text, [string]$title = 'SOA Extractor') {
  [System.Windows.Forms.MessageBox]::Show($text, $title,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}

# Browse-File : open-file dialog for the reference CSV. Starts in Documents.
# Returns the chosen path, or $null if cancelled.
function Browse-File([string]$title) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $title
  $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
  $docs = [Environment]::GetFolderPath('MyDocuments')
  if ($docs -and (Test-Path $docs)) { $dlg.InitialDirectory = $docs }
  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  return $dlg.FileName
}
