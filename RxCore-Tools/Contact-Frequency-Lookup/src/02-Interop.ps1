# ---- Win32 interop + UI Automation shortcuts ------------------------------
$sig = @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h,int n);
[DllImport("user32.dll")] public static extern bool SetWindowPos(System.IntPtr h,System.IntPtr ins,int x,int y,int cx,int cy,uint f);
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint f,uint x,uint y,uint d,int e);
[DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr h,out RECT r);
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)] public struct RECT{public int Left,Top,Right,Bottom;}
'@
if (-not ([System.Management.Automation.PSTypeName]'CF.Win').Type) {
  Add-Type -MemberDefinition $sig -Name Win -Namespace CF -PassThru | Out-Null
}
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]::Descendants
