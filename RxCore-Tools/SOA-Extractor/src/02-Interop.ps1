# ---- Win32 interop + UI Automation shortcuts ------------------------------
# Native window/mouse functions used to focus the app, click at screen
# coordinates, and capture screenshots.
$sig = @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool BringWindowToTop(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
[DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr hWnd, out RECT lpRect);
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
'@
if (-not ([System.Management.Automation.PSTypeName]'SOA.Win').Type) {
  Add-Type -MemberDefinition $sig -Name Win -Namespace SOA -PassThru | Out-Null
}

# Shorthand for the UI Automation entry type used throughout.
$AE = [System.Windows.Automation.AutomationElement]
