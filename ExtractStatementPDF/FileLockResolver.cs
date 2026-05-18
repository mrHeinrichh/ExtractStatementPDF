using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExtractStatementPDF
{
    internal static class FileLockResolver
    {
        private const int RmRebootReasonNone = 0;
        private const int CchRmSessionKey = 32;
        private const int CchRmMaxAppName = 255;
        private const int CchRmMaxSvcName = 63;
        private const int ErrorMoreData = 234;

        public static void KillLockingProcesses(string path)
        {
            foreach (var process in GetLockingProcesses(path))
            {
                try
                {
                    if (process.Id == Environment.ProcessId || process.HasExited)
                    {
                        continue;
                    }

                    process.Kill(true);
                    process.WaitForExit(5000);
                }
                catch
                {
                }
            }
        }

        private static IReadOnlyList<Process> GetLockingProcesses(string path)
        {
            var processes = new List<Process>();
            var sessionKey = Guid.NewGuid().ToString("N");
            var handle = 0u;

            var startSessionResult = RmStartSession(out handle, 0, sessionKey);
            if (startSessionResult != 0)
            {
                return processes;
            }

            try
            {
                var resources = new[] { path };
                var registerResult = RmRegisterResources(handle, (uint)resources.Length, resources, 0, null, 0, null);
                if (registerResult != 0)
                {
                    return processes;
                }

                uint processInfoNeeded = 0;
                uint processInfoCount = 0;
                uint rebootReasons = RmRebootReasonNone;
                var getListResult = RmGetList(handle, out processInfoNeeded, ref processInfoCount, null, ref rebootReasons);
                if (getListResult == ErrorMoreData)
                {
                    var processInfos = new RmProcessInfo[processInfoNeeded];
                    processInfoCount = processInfoNeeded;
                    getListResult = RmGetList(handle, out processInfoNeeded, ref processInfoCount, processInfos, ref rebootReasons);

                    if (getListResult == 0)
                    {
                        for (var i = 0; i < processInfoCount; i++)
                        {
                            var processInfo = processInfos[i];

                            try
                            {
                                var process = Process.GetProcessById(processInfo.Process.dwProcessId);
                                if (HasMatchingStartTime(process, processInfo.Process.ProcessStartTime))
                                {
                                    processes.Add(process);
                                }
                            }
                            catch
                            {
                            }
                        }
                    }
                }
            }
            finally
            {
                RmEndSession(handle);
            }

            return processes;
        }

        private static bool HasMatchingStartTime(Process process, FileTime processStartTime)
        {
            try
            {
                return process.StartTime.ToFileTime() == processStartTime.ToLong();
            }
            catch
            {
                return false;
            }
        }

        [DllImport("rstrtmgr", CharSet = CharSet.Unicode)]
        private static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, string strSessionKey);

        [DllImport("rstrtmgr", CharSet = CharSet.Unicode)]
        private static extern int RmRegisterResources(
            uint dwSessionHandle,
            uint nFiles,
            string[]? rgsFilenames,
            uint nApplications,
            [In] RmUniqueProcess[]? rgApplications,
            uint nServices,
            string[]? rgsServiceNames);

        [DllImport("rstrtmgr", CharSet = CharSet.Unicode)]
        private static extern int RmGetList(
            uint dwSessionHandle,
            out uint pnProcInfoNeeded,
            ref uint pnProcInfo,
            [In, Out] RmProcessInfo[]? rgAffectedApps,
            ref uint lpdwRebootReasons);

        [DllImport("rstrtmgr")]
        private static extern int RmEndSession(uint pSessionHandle);

        [StructLayout(LayoutKind.Sequential)]
        private struct FileTime
        {
            public uint DwLowDateTime;
            public uint DwHighDateTime;

            public long ToLong()
            {
                return ((long)DwHighDateTime << 32) | DwLowDateTime;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RmUniqueProcess
        {
            public int dwProcessId;
            public FileTime ProcessStartTime;
        }

        private enum RmAppType
        {
            RmUnknownApp = 0,
            RmMainWindow = 1,
            RmOtherWindow = 2,
            RmService = 3,
            RmExplorer = 4,
            RmConsole = 5,
            RmCritical = 1000
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct RmProcessInfo
        {
            public RmUniqueProcess Process;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CchRmMaxAppName + 1)]
            public string strAppName;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CchRmMaxSvcName + 1)]
            public string strServiceShortName;

            public RmAppType ApplicationType;
            public uint AppStatus;
            public uint TSSessionId;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bRestartable;
        }
    }
}
