# ---- Choose which Accounting instance to drive ----------------------------
# Lets you pick from the running Accounting windows (PID + title + exe path) when
# more than one is open, and remembers the choice ($script:AcctPid) so every other
# function targets that exact window. This is what lets it work on a PC that has a
# differently-located window or more than one Accounting open.

$script:AcctPid = $null

# Return the running Accounting windows as @{ Pid; Path; Title; Handle }.
function Get-AcctCandidates {
  $list = New-Object System.Collections.Generic.List[object]
  $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object {
    $_.MainWindowHandle -ne 0 -and (
      $_.ProcessName -like 'Accounting*' -or ($_.MainWindowTitle -like '*Accounting( Version*')
    )
  }
  foreach ($p in $procs) {
    $path = ''
    try { $path = $p.Path } catch {}
    $list.Add([pscustomobject]@{ Pid=$p.Id; Path=$path; Title=$p.MainWindowTitle; Handle=$p.MainWindowHandle })
  }
  return $list
}

# Show the picker (only when >1). Sets $script:AcctPid. Returns the chosen
# candidate, or $null if none running / cancelled.
function Choose-Acct {
  $cands = @(Get-AcctCandidates)
  if ($cands.Count -eq 0) { return $null }
  if ($cands.Count -eq 1) { $script:AcctPid = $cands[0].Pid; return $cands[0] }

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Choose the Accounting window to use"
  $form.Size = New-Object System.Drawing.Size(760, 300)
  $form.StartPosition = "CenterScreen"; $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.Text = "More than one Accounting is open - pick the one to extract from:"
  $label.Location = '12,10'; $label.AutoSize = $true; $form.Controls.Add($label)

  $lb = New-Object System.Windows.Forms.ListBox
  $lb.Location = '12,35'; $lb.Size = New-Object System.Drawing.Size(720, 180)
  $lb.Anchor = "Top,Left,Right,Bottom"; $lb.HorizontalScrollbar = $true
  foreach ($c in $cands) { [void]$lb.Items.Add(("PID {0}   |   {1}   |   {2}" -f $c.Pid, $c.Title, $c.Path)) }
  $lb.SelectedIndex = 0
  $form.Controls.Add($lb)

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "Use this one"; $ok.Location = '556,222'; $ok.Size = '95,30'
  $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK; $ok.Anchor = "Bottom,Right"
  $form.Controls.Add($ok); $form.AcceptButton = $ok

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"; $cancel.Location = '657,222'; $cancel.Size = '75,30'
  $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $cancel.Anchor = "Bottom,Right"
  $form.Controls.Add($cancel); $form.CancelButton = $cancel

  if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  $chosen = $cands[$lb.SelectedIndex]
  $script:AcctPid = $chosen.Pid
  return $chosen
}
