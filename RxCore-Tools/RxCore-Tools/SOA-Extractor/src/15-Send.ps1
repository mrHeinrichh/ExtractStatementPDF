# Send : send raw keystrokes to the focused control (SendKeys syntax, e.g. '{ENTER}').
function Send([string]$k) { [System.Windows.Forms.SendKeys]::SendWait($k) }
