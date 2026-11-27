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
                    case "watch": o.WatchSeconds = ParseInt(value(), name); break;
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
            "AttachLink.Watchdog（预研基础第 5 步 · 切片 A：只读诊断）\n" +
            "用法：\n" +
            "  AttachLink.Watchdog.exe --selftest\n" +
            "      内置单元回归（判定纯函数、日志轮转、参数解析），末行 RESULT: PASS / FAIL\n" +
            "  AttachLink.Watchdog.exe --diag [-watch 秒] [-XlamPath 路径] [-LogFile 路径] [-IntervalSeconds 秒]\n" +
            "      [-MaxCycles 次] [-ParentAllowlist 名单] [-HeartbeatSeconds 秒] [-AttachHiddenInstances]\n" +
            "      只读诊断：按与 PowerShell 版相同的规则给出每个 Excel 文档窗口的判定并写日志，不做任何写操作\n" +
            "常驻模式在切片 C 才启用。";

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
            if (o.Diag)
            {
                string logPath = o.LogFile.Length > 0 ? o.LogFile : Path.Combine(Path.GetTempPath(), "AttachLinkWatchdog.log");
                Logger logger = new Logger(logPath, o.LogMaxBytes, true);
                return new Engine(o, logger, true).RunDiag();
            }
            Console.WriteLine(Usage);
            return 2;
        }
    }
}
