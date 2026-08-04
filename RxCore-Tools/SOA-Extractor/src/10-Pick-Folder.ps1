# Pick-Folder : show the paste-a-path dialog (with a Browse... button).
# Returns the chosen folder path, or $null if cancelled.
function Pick-Folder {
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "SOA Extractor"
  $form.Size = New-Object System.Drawing.Size(620, 190)
  $form.StartPosition = "CenterScreen"
  $form.FormBorderStyle = "FixedDialog"
  $form.MaximizeBox = $false; $form.MinimizeBox = $false
  $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.Text = "Paste the folder path that contains the .xls files:"
  $label.Location = New-Object System.Drawing.Point(15, 15)
  $label.AutoSize = $true
  $form.Controls.Add($label)

  $box = New-Object System.Windows.Forms.TextBox
  $box.Location = New-Object System.Drawing.Point(15, 45)
  $box.Size = New-Object System.Drawing.Size(490, 25)
  $box.Anchor = "Top,Left,Right"
  $box.Add_KeyDown({ if ($_.Control -and $_.KeyCode -eq 'A') { $box.SelectAll(); $_.SuppressKeyPress = $true } })
  $form.Controls.Add($box)

  $browse = New-Object System.Windows.Forms.Button
  $browse.Text = "Browse..."
  $browse.Location = New-Object System.Drawing.Point(510, 44)
  $browse.Size = New-Object System.Drawing.Size(75, 25)
  $browse.Anchor = "Top,Right"
  $browse.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "Select the folder that contains the .xls files"
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $box.Text = $fb.SelectedPath }
  })
  $form.Controls.Add($browse)

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "OK"; $ok.Location = New-Object System.Drawing.Point(410, 100)
  $ok.Size = New-Object System.Drawing.Size(80, 30)
  $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
  $ok.Anchor = "Bottom,Right"
  $form.Controls.Add($ok); $form.AcceptButton = $ok

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"; $cancel.Location = New-Object System.Drawing.Point(500, 100)
  $cancel.Size = New-Object System.Drawing.Size(80, 30)
  $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $cancel.Anchor = "Bottom,Right"
  $form.Controls.Add($cancel); $form.CancelButton = $cancel

  $form.Add_Shown({ $form.Activate(); $box.Focus() })
  $res = $form.ShowDialog()
  if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  return $box.Text.Trim().Trim('"')
}
