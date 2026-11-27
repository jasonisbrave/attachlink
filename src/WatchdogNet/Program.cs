// =====================================================================
// AttachLink.Watchdog：命令行与入口
// ---------------------------------------------------------------------
// 参数沿用 PowerShell 版的名字与语义（-XlamPath / -IntervalSeconds / -MaxCycles / -ParentAllowlist /
// -HeartbeatSeconds / -LogFile / -AttachHiddenInstances），这样现有回归加一个开关就能换成 .NET 版来跑。
// 切片 A 只有 --selftest（内置单元回归）与 --diag（只读诊断：跑一个周期后退出；--watch N 持续观察 N 秒）。
// 退出码：0 = 正常；1 = --selftest 有失败；2 = 参数错误或请求了尚未启用的模式。
// =====================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace AttachLink.Watchdog
{
    internal sealed class Options
    {
        public string XlamPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "AddIns", Decisions.AddInName);
        public int IntervalSeconds = 2;
        public int MaxCycles = 0;
        public string ParentAllowlist = "";
        public int HeartbeatSeconds = 30;
        public string LogFile = "";
        public bool AttachHiddenInstances;
        public long LogMaxBytes = Logger.DefaultMaxBytes;
        public bool SelfTest;
        public bool Diag;
        public bool Once;                                   // 手动装载一次（切片 B）
        public int OnlyPid;                                 // 配合 --once：只处理这个 Excel 进程（-Pid）
        public bool DebugLog;                               // -DebugLog：记「[调试]」细节行（排障用）
        public int WatchSeconds;

        public static Options Parse(string[] args)
        {
            Options o = new Options();
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                if (a.Length < 2 || (a[0] != '-' && a[0] != '/')) throw new ArgumentException("不认识的参数: " + a);
                int c = a.IndexOf(':');
                string name = (c >= 0 ? a.Substring(0, c) : a).TrimStart('-', '/').ToLowerInvariant();
                string attached = c >= 0 ? a.Substring(c + 1) : null;
                Func<string> value = delegate
                {
                    if (attached != null) return attached;
                    if (i + 1 >= args.Length) throw new ArgumentException("参数 " + a + " 缺少值");
                    return args[++i];
                };
                switch (name)
                {
                    case "xlampath": o.XlamPath = value(); break;
                    case "intervalseconds": o.IntervalSeconds = ParseInt(value(), name); break;
                    case "maxcycles": o.MaxCycles = ParseInt(value(), name); break;
                    case "parentallowlist": o.ParentAllowlist = value(); break;
                    case "heartbeatseconds": o.HeartbeatSeconds = ParseInt(value(), name); break;
                    case "logfile": o.LogFile = value(); break;
                    case "logmaxbytes": o.LogMaxBytes = ParseInt(value(), name); break;
                    case "attachhiddeninstances": o.AttachHiddenInstances = attached == null ? true : ParseBool(attached, name); break;
                    case "selftest": o.SelfTest = true; break;
                    case "diag": o.Diag = true; break;
                    case "once": o.Once = true; break;
                    case "pid": o.OnlyPid = ParseInt(value(), name); break;
                    case "watch": o.WatchSeconds = ParseInt(value(), name); break;
                    case "debuglog": o.DebugLog = attached == null ? true : ParseBool(attached, name); break;
                    default: throw new ArgumentException("不认识的参数: " + a);
                }
            }
            return o;
        }

        private static int ParseInt(string s, string name)
        {
            int v;
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out v)) throw new ArgumentException("参数 " + name + " 不是整数: " + s);
            return v;
        }

        private static bool ParseBool(string s, string name)
        {
            if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase)) return true;
            if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase)) return false;
            throw new ArgumentException("参数 " + name + " 只接受 true / false: " + s);
        }
    }

    internal static class Program
    {
        public const string Usage =
            "AttachLink.Watchdog（v1.7.0：预编译的 .NET 看门狗）\n" +
            "用法：\n" +
            "  AttachLink.Watchdog.exe [选项]\n" +
            "      常驻：每个周期枚举 Excel 文档窗口，按规则装载并初始化 AttachLink.xlam，直到被结束\n" +
            "      （计划任务经 AttachLink.WatchdogLauncher.exe 以无窗口方式启动；-MaxCycles 次 仅供回归）\n" +
            "  AttachLink.Watchdog.exe --selftest\n" +
            "      内置单元回归（判定纯函数、日志轮转、参数解析），末行 RESULT: PASS / FAIL\n" +
            "  AttachLink.Watchdog.exe --diag [-watch 秒] [-MaxCycles 次] [选项]\n" +
            "      只读诊断：给出每个 Excel 文档窗口的判定并写日志，不做任何写操作\n" +
            "  AttachLink.Watchdog.exe --once [-Pid 进程号] [-watch 秒] [选项]\n" +
            "      手动装载一次：对可见的 Excel 工作窗口执行一次完整装载，然后退出\n" +
            "  选项（名字与 PowerShell 版相同）：-XlamPath 路径 -LogFile 路径 -IntervalSeconds 秒 -ParentAllowlist 名单\n" +
            "      -HeartbeatSeconds 秒 -AttachHiddenInstances（仅回归用）；排障用 -DebugLog\n" +
            "退出码：0 = 正常；1 = --selftest 有失败；2 = 参数错误。";

        [STAThread]
        private static int Main(string[] args)
        {
            try { Console.OutputEncoding = new UTF8Encoding(false); } catch { }
            Options o;
            try { o = Options.Parse(args); }
            catch (ArgumentException ex)
            {
                Console.WriteLine("参数错误: " + ex.Message);
                Console.WriteLine(Usage);
                return 2;
            }
            if (o.SelfTest) return SelfTest.Run();
            if (o.Diag && o.Once)
            {
                Console.WriteLine("参数错误: --diag 与 --once 不能同时给");
                return 2;
            }
            if (o.Diag || o.Once)
            {
                string logPath = o.LogFile.Length > 0 ? o.LogFile : Path.Combine(Path.GetTempPath(), "AttachLinkWatchdog.log");
                Logger logger = new Logger(logPath, o.LogMaxBytes, true);
                if (o.Diag) return new Engine(o, logger, true).Run("只读诊断");
                return new Engine(o, logger, false).Run("手动装载一次");
            }
            // 没给模式 = 常驻（由启动器以无窗口方式启动；不回显到控制台）
            string residentLog = o.LogFile.Length > 0 ? o.LogFile : Path.Combine(Path.GetTempPath(), "AttachLinkWatchdog.log");
            return new Engine(o, new Logger(residentLog, o.LogMaxBytes, false), false).RunResident();
        }
    }
}
