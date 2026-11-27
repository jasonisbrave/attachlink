// =====================================================================
// AttachLink.Watchdog：内置单元回归（--selftest）
// ---------------------------------------------------------------------
// 不引入第三方测试框架：一个控制台入口，把 scripts\test-watchdog-units.ps1 里纯函数的用例搬过来
// （G 段 Get-BackoffDelay、H 段 Get-InstanceDecision、A2 段隐藏文档窗口状态机、在位判定假对象用例），
// 再加上日志轮转与参数解析。末行 RESULT: PASS / FAIL；退出码 0 / 1。
// =====================================================================
using System;
using System.Collections.Generic;
using System.IO;

namespace AttachLink.Watchdog
{
    internal static class SelfTest
    {
        private static int _pass, _fail;

        private static void Check(bool cond, string label)
        {
            if (cond) { _pass++; Console.WriteLine("pass: " + label); }
            else { _fail++; Console.WriteLine("FAIL: " + label); }
        }

        private static WindowState NewState() { return new WindowState(); }

        public static int Run()
        {
            DateTime t0 = new DateTime(2026, 11, 27, 12, 0, 0);

            // ---- Get-BackoffDelay ----
            Check(Decisions.GetBackoffDelay(0, 2, 30) == 2, "退避：首次取 base（2 秒）");
            Check(Decisions.GetBackoffDelay(-1, 2, 30) == 2, "退避：prev ≤ 0 视为首次");
            Check(Decisions.GetBackoffDelay(2, 2, 30) == 4 && Decisions.GetBackoffDelay(4, 2, 30) == 8 && Decisions.GetBackoffDelay(8, 2, 30) == 16, "退避：逐次翻倍 2 → 4 → 8 → 16");
            Check(Decisions.GetBackoffDelay(16, 2, 30) == 30 && Decisions.GetBackoffDelay(30, 2, 30) == 30, "退避：封顶 max（30 秒）");
            Check(Decisions.GetBackoffDelay(2, 5, 3) == 3, "退避：base 大于 max 时以 max 为准");
            Check(Decisions.GetBackoffDelay(0.5, 2, 30) == 2, "退避：结果低于 base 时取 base");

            // ---- Get-PresenceDecision ----
            Check(Decisions.GetPresenceDecision(false, LookupState.Error, false) == Presence.NotLoaded, "在位判定：菜单不在位 → notloaded（不看三态）");
            Check(Decisions.GetPresenceDecision(true, LookupState.Found, true) == Presence.InPlace, "在位判定：菜单在位、在位的就是配置路径 → inplace");
            Check(Decisions.GetPresenceDecision(true, LookupState.Found, false) == Presence.Foreign, "在位判定：菜单在位、在位的是异路径副本 → foreign");
            Check(Decisions.GetPresenceDecision(true, LookupState.NotFound, false) == Presence.DeadMenu, "在位判定：菜单在位、工作簿明确 notfound → deadmenu");
            Check(Decisions.GetPresenceDecision(true, LookupState.Error, false) == Presence.Unknown, "在位判定：菜单在位、查询出错 → unknown");

            // ---- Get-InstanceDecision ----
            Check(Decisions.GetInstanceDecision(1, false, null, t0, 2), "实例判定：可见 → attach");
            Check(!Decisions.GetInstanceDecision(0, false, null, t0, 2), "实例判定：产品规则下隐藏 → defer");
            Check(!Decisions.GetInstanceDecision(-1, false, null, t0, 2), "实例判定：产品规则下读不到 → defer");
            Check(!Decisions.GetInstanceDecision(0, true, null, t0, 2), "实例判定：回归开关下隐藏、首次 → defer");
            Check(!Decisions.GetInstanceDecision(0, true, t0.AddSeconds(-1), t0, 2), "实例判定：回归开关下隐藏、宽限未满 → defer");
            Check(Decisions.GetInstanceDecision(0, true, t0.AddSeconds(-2), t0, 2), "实例判定：回归开关下隐藏、宽限已满 → attach");
            Check(Decisions.GetInstanceDecision(-1, true, null, t0, 2), "实例判定：回归开关下读不到 → attach");

            // ---- Get-HiddenDecision ----
            Check(Decisions.GetHiddenDecision(NewState(), -1, true, t0, 2) == HiddenVerdict.Wait, "隐藏窗口：Window.Visible 读不到 → wait");
            Check(Decisions.GetHiddenDecision(NewState(), 1, true, t0, 2) == HiddenVerdict.Process, "隐藏窗口：Excel 报告可见 → process");
            WindowState h1 = NewState();
            Check(Decisions.GetHiddenDecision(h1, 0, false, t0, 2) == HiddenVerdict.Wait && h1.Since == t0, "隐藏窗口：首次报告隐藏 → wait，记下时刻");
            Check(Decisions.GetHiddenDecision(h1, 0, false, t0.AddSeconds(1), 2) == HiddenVerdict.Wait && !h1.Done, "隐藏窗口：宽限未满 → wait，不标 done");
            Check(Decisions.GetHiddenDecision(h1, 0, false, t0.AddSeconds(3), 2) == HiddenVerdict.Skip && h1.Done && h1.Shown == false && h1.LastCheck == t0.AddSeconds(3), "隐藏窗口：宽限已满 → skip");
            WindowState h2 = NewState(); h2.UnknownSince = t0;
            Decisions.GetHiddenDecision(h2, 1, true, t0, 2);
            Check(h2.UnknownSince == null && !h2.UnknownEscalated, "隐藏窗口：重新读到可见后清掉「读不到」的计时");

            // ---- Test-HiddenReevaluate（maxReeval=3，recheck=30）----
            WindowState r1 = NewState(); r1.Done = true; r1.Shown = false;
            Check(!Decisions.TestHiddenReevaluate(r1, false, t0, 3, 30) && r1.Done && !r1.Shown, "重评估：Win32 持续不可见 → 不重查");
            Check(Decisions.TestHiddenReevaluate(r1, true, t0, 3, 30) && !r1.Done && r1.Rev == 1 && r1.Since == t0 && r1.LastCheck == t0, "重评估：不可见 → 可见的跳变立即触发，计数 1");
            WindowState r2 = NewState(); r2.Done = true; r2.Shown = true; r2.Rev = 1; r2.LastCheck = t0.AddSeconds(-10);
            Check(!Decisions.TestHiddenReevaluate(r2, true, t0, 3, 30), "重评估：Win32 持续可见、没有新跳变、低频重查未到 → 不重查");
            WindowState r3 = NewState(); r3.Done = true; r3.Shown = true; r3.Rev = 1; r3.LastCheck = t0.AddSeconds(-31);
            Check(Decisions.TestHiddenReevaluate(r3, true, t0, 3, 30) && r3.Rev == 1, "重评估：Win32 持续可见、距上次评估超过 30 秒 → 低频重查，不消耗跳变计数");
            WindowState r4 = NewState(); r4.Done = true; r4.Shown = false; r4.Rev = 3; r4.LastCheck = t0.AddSeconds(-1);
            Check(!Decisions.TestHiddenReevaluate(r4, true, t0, 3, 30) && r4.Rev == 3 && r4.LastCheck == t0.AddSeconds(-1), "重评估：跳变计数用完后，新跳变本身不再触发");
            WindowState r5 = NewState(); r5.Done = true; r5.Shown = false; r5.Rev = 3; r5.LastCheck = t0.AddSeconds(-31);
            Check(Decisions.TestHiddenReevaluate(r5, true, t0, 3, 30) && r5.Rev == 3, "重评估：计数用完仍可经低频重查触发（不会被永久跳过）");
            WindowState r6 = NewState(); r6.Done = true; r6.Shown = true; r6.Rev = 3; r6.LastCheck = t0.AddSeconds(-31);
            bool re6 = Decisions.TestHiddenReevaluate(r6, true, t0, 3, 30);
            bool again6 = Decisions.GetHiddenDecision(r6, 0, true, t0, 2) == HiddenVerdict.Wait
                        && Decisions.GetHiddenDecision(r6, 0, true, t0.AddSeconds(2), 2) == HiddenVerdict.Skip && r6.Done;
            Check(re6 && again6, "重评估：重查时 Excel 仍报告隐藏 → 宽限一个周期后再次 skip（可再次被重查）");

            // ---- Test-SamePath ----
            Check(Decisions.SamePath(@"C:\A\b.xlam", @"c:\a\B.xlam"), "路径比较：不区分大小写");
            Check(Decisions.SamePath(@"C:\a\..\a\b.xlam", @"C:\a\b.xlam"), "路径比较：先规范化再比较");
            Check(!Decisions.SamePath(@"C:\A\b.xlam", @"C:\A\c.xlam") && !Decisions.SamePath("", @"C:\a"), "路径比较：不同路径 / 无效路径 → false");

            // ---- 日志轮转 ----
            string dir = Path.Combine(Path.GetTempPath(), "AttachLinkWatchdogSelfTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string main = Path.Combine(dir, "AttachLinkWatchdog.log");
                Logger logger = new Logger(main, 300, false);
                Check(logger.PreviousPath == Path.Combine(dir, "AttachLinkWatchdog.prev.log"), "日志：旧文件名 = AttachLinkWatchdog.prev.log");
                int total = 40;
                for (int i = 1; i <= total; i++) logger.Log("selftest-line-" + i + "-AttachLinkWatchdog-padding-padding");
                string[] pl = File.ReadAllLines(logger.PreviousPath);
                string[] ml = File.ReadAllLines(main);
                // 只保留一份旧文件：40 行、每 5 行转一次，最后旧文件是第 31～35 行、新文件是第 36～40 行
                int LineNo(string line) { int a = line.IndexOf("selftest-line-", StringComparison.Ordinal) + 14; int b = line.IndexOf("-AttachLinkWatchdog", a); return int.Parse(line.Substring(a, b - a), System.Globalization.CultureInfo.InvariantCulture); }
                Check(pl.Length > 0 && ml.Length > 0 && LineNo(ml[ml.Length - 1]) == total && LineNo(pl[pl.Length - 1]) + 1 == LineNo(ml[0])
                      && pl[pl.Length - 1].EndsWith("-padding-padding", StringComparison.Ordinal), "日志：超过上限后转成一份旧文件（只保留一份），两份首尾相接（旧 " + LineNo(pl[0]) + "～" + LineNo(pl[pl.Length - 1]) + "，新 " + LineNo(ml[0]) + "～" + LineNo(ml[ml.Length - 1]) + "）");
                Check(new FileInfo(main).Length <= 300 + 80, "日志：新文件不超过「上限 + 一行」");
                Check(ml[0].Length > 15 && char.IsDigit(ml[0][0]) && ml[0][2] == '-' && ml[0][5] == ' ' && ml[0][8] == ':' && ml[0][11] == ':' && ml[0][14] == ' ', "日志：行格式「MM-dd HH:mm:ss 内容」");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }

            // ---- 参数解析 ----
            Options p1 = Options.Parse(new[] { "-XlamPath", @"C:\a.xlam", "-IntervalSeconds", "3", "-AttachHiddenInstances" });
            Check(p1.XlamPath == @"C:\a.xlam" && p1.IntervalSeconds == 3 && p1.AttachHiddenInstances && !p1.Diag, "参数：-Name 值 与开关");
            Options p2 = Options.Parse(new[] { "/LogFile:C:\\x.log", "--diag", "-watch:5", "-AttachHiddenInstances:false" });
            Check(p2.LogFile == "C:\\x.log" && p2.Diag && p2.WatchSeconds == 5 && !p2.AttachHiddenInstances, "参数：/Name:值、--diag、-Name:false");
            bool threw1 = false, threw2 = false;
            try { Options.Parse(new[] { "-Bogus" }); } catch (ArgumentException) { threw1 = true; }
            try { Options.Parse(new[] { "-IntervalSeconds", "abc" }); } catch (ArgumentException) { threw2 = true; }
            Check(threw1 && threw2, "参数：不认识的参数 / 非整数的值 → 报错");

            Console.WriteLine("RESULT: " + (_fail == 0 ? "PASS" : "FAIL") + "（pass=" + _pass + " FAIL=" + _fail + "）");
            return _fail == 0 ? 0 : 1;
        }
    }
}
