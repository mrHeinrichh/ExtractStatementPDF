# Send : send raw keystrokes to the focused control (SendKeys syntax).
function Send([string]$k){ [System.Windows.Forms.SendKeys]::SendWait($k) }
