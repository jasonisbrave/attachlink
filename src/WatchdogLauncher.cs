// =====================================================================
// AttachLink 看门狗启动器（AttachLink.WatchdogLauncher.exe）
// ---------------------------------------------------------------------
// 计划任务的直接动作：无窗口地启动常驻看门狗，看着它运行，异常退出时把它拉起来；自己结束时把看门狗的退出码返回给计划任务。
// 看门狗有两种引擎，本程序旁边有 AttachLink.Watchdog.exe 就用它（v1.7.0 起的默认，预编译的 .NET 程序），
// 否则退回 Windows PowerShell 5.1 + watchdog.ps1（部署时带 -LegacyPowerShell）。两种引擎的监管方式相同：
//   - 无控制台窗口：CreateNoWindow（不是 -WindowStyle Hidden 那种事后隐藏）。Windows 11 的
//     默认终端是 Windows Terminal 时，直接拿 powershell.exe 当任务动作会被交给终端窗口托管、
//     留在桌面上（2026-11-27 实测）；本程序是 GUI 子系统程序，自己也没有窗口。
//   - 可被计划任务监管：计划任务跟踪的就是本进程，看门狗运行期间任务状态是 Running；停止任务、
//     注销时计划任务结束本进程。计划任务的「失败后重启」只在「动作没能启动」时生效，进程以非 0
//     退出码结束时不会触发（2026-11-27 实测，按触发启动的任务也一样），所以拉起看门狗由本程序自己做。
//   - 拉起策略：看门狗非 0 退出后等一会儿再启动，连续 3 次没能稳定运行（每次不足 60 秒）就放弃，
//     以它的退出码结束本进程。稳定运行过 60 秒之后再出事，重新计数。
//   - 不留孤儿：看门狗进程放进「关闭句柄即结束」的作业对象（JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE），
//     本进程一死（被结束、被停止任务），看门狗随之结束。
//   - 记录：启动器的动作写进看门狗的日志（默认 %TEMP%\AttachLinkWatchdog.log，与看门狗同一格式）。
// 可选参数（只供回归使用；计划任务不带任何参数）：
//   -LogFile <路径>  -ChildLog <路径>  -MaxFailures <n>  -RestartDelaySeconds <s>  -StableSeconds <s>
// 由 build\make-launcher.ps1 用系统自带的 csc.exe 编译（.NET Framework 4.x）；只用 C# 5 语法。
// 退出码：看门狗的退出码（放弃时为最后一次的退出码）；2 = 两种引擎的看门狗都不在本程序旁边；
//         3 = 找不到 Windows PowerShell（只在回退到 PowerShell 版时才会检查）；4 = 启动看门狗失败；5 = 参数不认识。
// =====================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace AttachLink
{
    internal static class WatchdogLauncher
    {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string lpName);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(IntPtr hJob, int infoClass, ref JOBOBJECT_EXTENDED_LIMIT_INFORMATION lpInfo, int cbInfo);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential)]
        private struct JOBOBJECT_BASIC_LIMIT_INFORMATION
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct IO_COUNTERS
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
            public IO_COUNTERS IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        private const int JobObjectExtendedLimitInformation = 9;
        private const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000;

        private const int ExitDeployIncomplete = 2;
        private const int ExitNoPowerShell = 3;
        private const int ExitStartFailed = 4;
        private const int ExitBadArgument = 5;

        // 产品默认值（计划任务不带参数）
        private static int maxFailures = 3;
        private static int restartDelaySeconds = 10;
        private static int stableSeconds = 60;
        private static string logFile = Path.Combine(Path.GetTempPath(), "AttachLinkWatchdog.log");
        private static string childLogFile = "";       // 仅回归：把看门狗自己的日志改到指定文件（-ChildLog）

        // 日志行格式与 watchdog.ps1 的 WdLog 一致（UTF-8，无 BOM，CRLF）
        private static void Log(string message)
        {
            try
            {
                File.AppendAllText(logFile, DateTime.Now.ToString("MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                    + " " + message + "\r\n", new UTF8Encoding(false));
            }
            catch (Exception) { /* 日志写不进去不能影响看门狗 */ }
        }

        private static bool TryParseInt(string s, out int value)
        {
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        private static int Main(string[] args)
        {
            // ---- 可选参数（只供回归使用）----
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i].ToLowerInvariant();
                if (a == "-logfile" && i + 1 < args.Length) { logFile = args[++i]; }
                else if (a == "-childlog" && i + 1 < args.Length) { childLogFile = args[++i]; }
                else if (a == "-maxfailures" && i + 1 < args.Length)
                {
                    int v; if (!TryParseInt(args[++i], out v) || v < 1) { return ExitBadArgument; } maxFailures = v;
                }
                else if (a == "-restartdelayseconds" && i + 1 < args.Length)
                {
                    int v; if (!TryParseInt(args[++i], out v) || v < 0) { return ExitBadArgument; } restartDelaySeconds = v;
                }
                else if (a == "-stableseconds" && i + 1 < args.Length)
                {
                    int v; if (!TryParseInt(args[++i], out v) || v < 0) { return ExitBadArgument; } stableSeconds = v;
                }
                else { return ExitBadArgument; }
            }

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string netWatchdog = Path.Combine(baseDir, "AttachLink.Watchdog.exe");   // v1.7.0：预编译的 .NET 看门狗，在它就优先
            string script = Path.Combine(baseDir, "watchdog.ps1");                   // 回退：PowerShell 版（部署时带 -LegacyPowerShell）
            string xlam = Path.Combine(baseDir, "AttachLink.xlam");
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\powershell.exe");
            bool useNet = File.Exists(netWatchdog);
            if (!useNet)
            {
                if (!File.Exists(script)) { Log("启动器: 找不到 watchdog.ps1，启动器退出（退出码=2）"); return ExitDeployIncomplete; }
                if (!File.Exists(ps)) { Log("启动器: 找不到 Windows PowerShell，启动器退出（退出码=3）"); return ExitNoPowerShell; }
            }
            Log("启动器: 启动看门狗（引擎=" + (useNet ? ".NET：AttachLink.Watchdog.exe" : "PowerShell：watchdog.ps1") + "）");

            // 作业对象：句柄关闭（本进程结束）时结束作业内的一切进程。
            // 设不上就不挂——看门狗照常启动，只是失去「本进程一死它也结束」的保障。
            IntPtr job = CreateJobObject(IntPtr.Zero, null);
            if (job != IntPtr.Zero)
            {
                JOBOBJECT_EXTENDED_LIMIT_INFORMATION info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION();
                info.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
                if (!SetInformationJobObject(job, JobObjectExtendedLimitInformation, ref info,
                        Marshal.SizeOf(typeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION))))
                {
                    CloseHandle(job);
                    job = IntPtr.Zero;
                }
            }

            // 两种引擎收到的 -XlamPath 相同；-ChildLog（仅回归）把它们各自的日志改到指定文件。
            // PowerShell 那一路的参数与旧 VBS 启动器传给 powershell.exe 的一致（窗口风格那一项在无窗口控制台下不起作用，留作双保险）
            string childLog = childLogFile.Length > 0 ? " -LogFile \"" + childLogFile + "\"" : "";
            ProcessStartInfo psi;
            if (useNet)
            {
                psi = new ProcessStartInfo(netWatchdog, "-XlamPath \"" + xlam + "\"" + childLog);
            }
            else
            {
                psi = new ProcessStartInfo(ps,
                    "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + script +
                    "\" -XlamPath \"" + xlam + "\"" + childLog);
            }
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.WorkingDirectory = baseDir;

            int failures = 0;
            int lastCode = 0;
            while (true)
            {
                Process child;
                try { child = Process.Start(psi); }
                catch
                {
                    Log("启动器: 看门狗没能启动，启动器退出（退出码=4）");
                    if (job != IntPtr.Zero) { CloseHandle(job); }
                    return ExitStartFailed;
                }
                if (job != IntPtr.Zero)
                {
                    try { AssignProcessToJobObject(job, child.Handle); }
                    catch
                    {
                        CloseHandle(job);
                        job = IntPtr.Zero;
                    }
                }

                DateTime startedAt = DateTime.Now;
                child.WaitForExit();
                int code = child.ExitCode;
                int ranSeconds = (int)(DateTime.Now - startedAt).TotalSeconds;

                if (code == 0)
                {
                    // 正常结束（只在带 -MaxCycles 运行时发生）：不留任何东西，原样返回
                    if (job != IntPtr.Zero) { CloseHandle(job); }
                    return 0;
                }

                lastCode = code;
                // 稳定运行过一段时间之后再出事，重新计数
                if (ranSeconds >= stableSeconds) { failures = 1; } else { failures++; }
                Log("启动器: 看门狗已结束（退出码=" + code.ToString(CultureInfo.InvariantCulture)
                    + "，运行 " + ranSeconds.ToString(CultureInfo.InvariantCulture) + " 秒）");
                if (failures >= maxFailures)
                {
                    Log("启动器: 看门狗连续 " + failures.ToString(CultureInfo.InvariantCulture)
                        + " 次没能稳定运行，启动器退出（退出码=" + code.ToString(CultureInfo.InvariantCulture)
                        + "）；计划任务的下一次触发会再启动");
                    if (job != IntPtr.Zero) { CloseHandle(job); }
                    return code;
                }
                Log("启动器: " + restartDelaySeconds.ToString(CultureInfo.InvariantCulture) + " 秒后重启看门狗（第 "
                    + failures.ToString(CultureInfo.InvariantCulture) + "/" + maxFailures.ToString(CultureInfo.InvariantCulture) + " 次）");
                if (restartDelaySeconds > 0) { Thread.Sleep(TimeSpan.FromSeconds(restartDelaySeconds)); }
            }
        }
    }
}
