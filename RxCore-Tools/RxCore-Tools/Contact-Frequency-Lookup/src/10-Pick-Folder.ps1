# Pick-Folder : paste-a-path dialog (with Browse...). Returns the path or $null.
function Pick-Folder {
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Contact Frequency Lookup"; $form.Size = New-Object System.Drawing.Size(620,190)
  $form.StartPosition = "CenterScreen"; $form.FormBorderStyle = "FixedDialog"
  $form.MaximizeBox=$false; $form.MinimizeBox=$false; $form.TopMost=$true
  $l = New-Object System.Windows.Forms.Label; $l.Text="Paste the folder path that contains the .xls files:"; $l.Location='15,15'; $l.AutoSize=$true; $form.Controls.Add($l)
  $box = New-Object System.Windows.Forms.TextBox; $box.Location='15,45'; $box.Size='490,25'; $box.Anchor='Top,Left,Right'
  $box.Add_KeyDown({ if ($_.Control -and $_.KeyCode -eq 'A') { $box.SelectAll(); $_.SuppressKeyPress=$true } }); $form.Controls.Add($box)
  $br = New-Object System.Windows.Forms.Button; $br.Text="Browse..."; $br.Location='510,44'; $br.Size='75,25'; $br.Anchor='Top,Right'
  $br.Add_Click({ $fb=New-Object System.Windows.Forms.FolderBrowserDialog; if($fb.ShowDialog() -eq 'OK'){ $box.Text=$fb.SelectedPath } }); $form.Controls.Add($br)
  $ok = New-Object System.Windows.Forms.Button; $ok.Text="OK"; $ok.Location='410,100'; $ok.Size='80,30'; $ok.DialogResult='OK'; $ok.Anchor='Bottom,Right'; $form.Controls.Add($ok); $form.AcceptButton=$ok
  $cx = New-Object System.Windows.Forms.Button; $cx.Text="Cancel"; $cx.Location='500,100'; $cx.Size='80,30'; $cx.DialogResult='Cancel'; $cx.Anchor='Bottom,Right'; $form.Controls.Add($cx); $form.CancelButton=$cx
  $form.Add_Shown({ $form.Activate(); $box.Focus() })
  if ($form.ShowDialog() -ne 'OK') { return $null }
  return $box.Text.Trim().Trim('"')
}
