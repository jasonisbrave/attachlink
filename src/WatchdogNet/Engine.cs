// =====================================================================
// AttachLink.Watchdog：主循环（切片 A：只读诊断）
// ---------------------------------------------------------------------
// 逐行移植 scripts\watchdog.ps1 的主循环：周期行、宿主身份校验、附着退避、隐藏文档窗口状态机、
// 实例可见性、父进程白名单、在位判定——判定与日志措辞与 PowerShell 版相同。
// 切片 A 里「写」的路径（删除死菜单控件、打开加载项、执行宏、写回宿主设置）全部不执行：
// 走到该装载的地方只记一行「[只读诊断]」，由切片 B / C 接上。
// =====================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace AttachLink.Watchdog
{
    internal sealed class Engine
    {
        private sealed class BackoffEntry { public DateTime Until; public double Delay; }

        private const int MaxReeval = 3;                       // 已判定为隐藏的窗口因「Win32 由不可见变为可见」而立即重新评估的次数上限
        private const int HiddenRecheckSeconds = 30;           // Win32 可见却没有新跳变时的低频只读重查间隔（秒）；不设次数上限
        private const double BackoffMaxSeconds = 30;           // 退避间隔上限（秒）
        private const int UnknownLogSeconds = 60;

        private readonly Options _opt;
        private readonly Logger _log;
        private readonly bool _readOnly;
        private readonly HashSet<string> _seen = new HashSet<string>(StringComparer.Ordinal);                       // 窗口键 hwnd:pid
        private readonly Dictionary<string, WindowState> _hiddenWin = new Dictionary<string, WindowState>(StringComparer.Ordinal);
        private readonly Dictionary<uint, DateTime> _hiddenSince = new Dictionary<uint, DateTime>();               // pid → 首次见到（隐藏）的时刻
        private readonly Dictionary<string, BackoffEntry> _backoff = new Dictionary<string, BackoffEntry>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _lastReadOnlyNote = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly HashSet<string> _deadMenuLogged = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _filter = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private DateTime? _lastPresenceUnknownLog;
        private int _cycle;
        private int _lastCount = -1;
        private DateTime _lastBeat;
        private bool _xlamMissing;

        public Engine(Options opt, Logger log, bool readOnly)
        {
            _opt = opt; _log = log; _readOnly = readOnly;
            string filter = _opt.ParentAllowlist;
            if (string.IsNullOrEmpty(filter))
            {
                // 未指定参数时读取程序同目录 watchdog.conf 的 ParentAllowlist= 行
                try
                {
                    string conf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "watchdog.conf");
                    if (File.Exists(conf))
                    {
                        foreach (string raw in File.ReadAllLines(conf))
                        {
                            string line = raw.Trim();
                            if (line.StartsWith("ParentAllowlist", StringComparison.OrdinalIgnoreCase))
                            {
                                int eq = line.IndexOf('=');
                                if (eq >= 0) { filter = line.Substring(eq + 1).Trim(); break; }
                            }
                        }
                    }
                }
                catch { }
            }
            if (!string.IsNullOrEmpty(filter))
            {
                foreach (string part in filter.Split(',')) { string p = part.Trim().ToLowerInvariant(); if (p.Length > 0) _filter.Add(p); }
                List<string> ordered = new List<string>(_filter);
                _log.Log("父进程白名单生效: " + string.Join(", ", ordered.ToArray()));
            }
        }

        public int RunDiag()
        {
            _log.Log((_readOnly ? "看门狗启动（只读诊断）" : "看门狗启动") + ", pid=" + Process.GetCurrentProcess().Id + ", xlam=" + _opt.XlamPath);
            DateTime start = DateTime.Now;
            while (true)
            {
                _cycle++;
                RunCycle();
                if (_opt.MaxCycles > 0 && _cycle >= _opt.MaxCycles) break;
                if (_opt.WatchSeconds > 0) { if ((DateTime.Now - start).TotalSeconds >= _opt.WatchSeconds) break; }
                else if (_opt.MaxCycles <= 0) break;            // 没给 --watch 也没给 -MaxCycles：只跑一个周期
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(_opt.IntervalSeconds));
            }
            return 0;
        }

        private void RunCycle()
        {
            try
            {
                DateTime now = DateTime.Now;
                if (!File.Exists(_opt.XlamPath))
                {
                    // xlam 缺失（升级 / 卸载进行中、文件被移走）：不向任何宿主发 COM 调用，静默等待
                    if (!_xlamMissing)
                    {
                        _xlamMissing = true; _lastBeat = now;
                        _log.Log("xlam 不存在，暂停附着直至文件就位: " + _opt.XlamPath);
                    }
                    else if ((now - _lastBeat).TotalSeconds >= _opt.HeartbeatSeconds)
                    {
                        _lastBeat = now;
                        _log.Log("cycle " + _cycle + ": xlam 仍不存在，继续等待");
                    }
                }
                else
                {
                    if (_xlamMissing) { _xlamMissing = false; _lastCount = -1; _log.Log("xlam 已就位，恢复附着"); }
                    List<long> wins = Native.FindExcel7Windows();
                    if (wins.Count != _lastCount || (now - _lastBeat).TotalSeconds >= _opt.HeartbeatSeconds)
                    {
                        _lastCount = wins.Count; _lastBeat = now;
                        _log.Log("cycle " + _cycle + ": EXCEL7 窗口数=" + wins.Count);
                    }
                    foreach (long h in wins) ProcessWindow(h);
                    // 清理已消失进程 / 窗口的记录
                    HashSet<uint> livePids = new HashSet<uint>();
                    HashSet<string> liveKeys = new HashSet<string>(StringComparer.Ordinal);
                    foreach (long h in wins)
                    {
                        uint p = Native.GetWindowProcessId(h);
                        livePids.Add(p); liveKeys.Add(h + ":" + p);
                    }
                    RemoveWhere(_hiddenSince, kv => !livePids.Contains(kv.Key));
                    RemoveWhere(_hiddenWin, kv => !liveKeys.Contains(kv.Key));
                    RemoveWhere(_backoff, kv => !liveKeys.Contains(kv.Key));
                    RemoveWhere(_lastReadOnlyNote, kv => !liveKeys.Contains(kv.Key));
                    _deadMenuLogged.RemoveWhere(k => !liveKeys.Contains(k));
                }
            }
            catch (Exception ex)
            {
                _log.Log("cycle " + _cycle + " 异常: " + ex.Message);
            }
        }

        private static void RemoveWhere<TKey, TValue>(Dictionary<TKey, TValue> dict, Func<KeyValuePair<TKey, TValue>, bool> dead)
        {
            List<TKey> gone = null;
            foreach (KeyValuePair<TKey, TValue> kv in dict) { if (dead(kv)) { if (gone == null) gone = new List<TKey>(); gone.Add(kv.Key); } }
            if (gone != null) foreach (TKey k in gone) dict.Remove(k);
        }

        private void ProcessWindow(long h)
        {
            uint winPid = Native.GetWindowProcessId(h);
            string key = h + ":" + winPid;
            DateTime now = DateTime.Now;
            if (_seen.Contains(key))
            {
                // 已判定为「隐藏窗口、不作处理对象」的窗口：每个周期只做一次 Win32 可见性查询来决定是否重新评估
                WindowState st0;
                if (_hiddenWin.TryGetValue(key, out st0) && st0.Done
                    && Decisions.TestHiddenReevaluate(st0, Native.IsWindowShown(h), now, MaxReeval, HiddenRecheckSeconds))
                {
                    _seen.Remove(key); _lastReadOnlyNote.Remove(key);
                    _log.Log("先前隐藏的窗口已显示，重新评估: hwnd=" + h + " pid=" + winPid);
                }
                else return;
            }

            // ---- 宿主身份校验：只附着 Microsoft Excel（EXCEL.EXE）。必须在取得 Application 对象之前完成（fail-closed）----
            string hostName = "";
            try { hostName = Process.GetProcessById((int)winPid).ProcessName; } catch { }
            if (!string.Equals(hostName, "EXCEL", StringComparison.OrdinalIgnoreCase))
            {
                _seen.Add(key);
                hostName = hostName.Length > 0 ? hostName + ".exe" : "(进程已退出/不可识别)";
                _log.Log("跳过 hwnd=" + h + "：宿主进程=" + hostName + "，仅附着 EXCEL.EXE");
                return;
            }

            // ---- 附着退避（v1.6.1）：这个窗口上一次附着失败后，退避期内不向宿主发任何调用 ----
            BackoffEntry bo;
            if (_backoff.TryGetValue(key, out bo) && now < bo.Until) return;

            bool attachFailed = false;
            ComScope scope = new ComScope();
            try
            {
                int winVisible = -1;
                object win = scope.Track(Native.GetExcelWindow(h));
                object app = null;
                if (win != null)
                {
                    try { winVisible = ToVisible(Com.Get(win, "Visible")); } catch { }
                    app = scope.Track(Com.Get(win, "Application"));
                    scope.Release(win);      // Window 对象用完即释放，不留给 GC
                }
                if (app == null) { _log.Log("hwnd=" + h + " 未取得 Application"); return; }
                AttachWindow(h, winPid, key, app, winVisible, scope, now);
            }
            catch (Exception ex)
            {
                attachFailed = true;
                _log.Log("附着失败 hwnd=" + h + " : " + ex.Message + "（退避重试）");
            }
            finally
            {
                // 用完立即释放对宿主的对象引用；GC 只作兜底
                scope.Dispose();
                GC.Collect(); GC.WaitForPendingFinalizers();
                // 退避记账：失败的窗口按 GetBackoffDelay 排下次可试的时刻；得出任何非失败的结论即清零
                if (attachFailed)
                {
                    double prev = 0;
                    BackoffEntry old; if (_backoff.TryGetValue(key, out old)) prev = old.Delay;
                    double next = Decisions.GetBackoffDelay(prev, _opt.IntervalSeconds, BackoffMaxSeconds);
                    _backoff[key] = new BackoffEntry { Until = DateTime.Now.AddSeconds(next), Delay = next };
                }
                else _backoff.Remove(key);
            }
        }

        private static int ToVisible(object v)
        {
            if (v == null) return -1;
            if (v is bool) return (bool)v ? 1 : 0;
            try { return Convert.ToInt32(v, CultureInfo.InvariantCulture) != 0 ? 1 : 0; } catch { return -1; }
        }

        private void AttachWindow(long h, uint winPid, string key, object app, int winVisible, ComScope scope, DateTime now)
        {
            // ---- 隐藏 / 状态未知的文档窗口（判定在 Decisions.GetHiddenDecision 里）----
            WindowState st;
            if (!_hiddenWin.TryGetValue(key, out st)) { st = new WindowState(); _hiddenWin[key] = st; }
            HiddenVerdict decision = Decisions.GetHiddenDecision(st, winVisible, Native.IsWindowShown(h), now, _opt.IntervalSeconds);
            if (decision == HiddenVerdict.Wait)
            {
                if (winVisible < 0)
                {
                    if (!st.UnknownSince.HasValue)
                    {
                        st.UnknownSince = now;
                        _log.Log("Window.Visible 读不到，本周期不处理（下周期重试）: hwnd=" + h + " pid=" + winPid);
                    }
                    else if (!st.UnknownEscalated && (now - st.UnknownSince.Value).TotalSeconds >= UnknownLogSeconds)
                    {
                        st.UnknownEscalated = true;
                        _log.Log("Window.Visible 已持续 60 秒读不到，该窗口一直没有被处理: hwnd=" + h + " pid=" + winPid);
                    }
                }
                return;
            }
            if (decision == HiddenVerdict.Skip)
            {
                _seen.Add(key);
                if (!st.Logged) { st.Logged = true; _log.Log("隐藏窗口（加载项自身或被隐藏的工作簿），不作为处理对象: hwnd=" + h + " pid=" + winPid); }
                return;
            }
            _hiddenWin.Remove(key);      // Excel 报告为可见：按工作窗口处理

            // ---- 实例可见性（v1.6.2：只附着用户看得见的会话）----
            int appVisible = -1;
            try { appVisible = ToVisible(Com.Get(app, "Visible")); } catch { appVisible = -1; }
            DateTime firstSeen; bool hasFirst = _hiddenSince.TryGetValue(winPid, out firstSeen);
            if (!Decisions.GetInstanceDecision(appVisible, _opt.AttachHiddenInstances, hasFirst ? (DateTime?)firstSeen : null, now, _opt.IntervalSeconds))
            {
                if (!hasFirst)
                {
                    _hiddenSince[winPid] = now;
                    if (_opt.AttachHiddenInstances) _log.Log("隐藏实例，宽限一个周期再处理: pid=" + winPid);
                    else if (appVisible == 0) _log.Log("隐藏实例，暂不附着（等它可见）: pid=" + winPid);
                    else _log.Log("实例可见性暂时无法确认，本周期不处理（下周期重试）: pid=" + winPid);
                }
                return;
            }
            if (hasFirst)
            {
                _hiddenSince.Remove(winPid);
                if (!_opt.AttachHiddenInstances) _log.Log("先前隐藏的实例已可见，开始处理: pid=" + winPid);
            }

            // ---- 父进程白名单：仅附着白名单进程拉起的 Excel ----
            if (_filter.Count > 0)
            {
                string pname = Native.GetParentProcessName(winPid);
                if (string.IsNullOrEmpty(pname)) { _log.Log("hwnd=" + h + " 父进程未知，本周期跳过（不缓存，下轮重查）"); return; }
                if (!_filter.Contains(pname)) { _seen.Add(key); _log.Log("跳过 hwnd=" + h + "：父进程 " + pname + " 不在白名单"); return; }
            }

            // ---- 在位判定（v1.6.0 加固，R2）：菜单 Tag + Workbooks 三态联合确认 ----
            bool deadMenu = false;
            if (MenuExists(app, scope))
            {
                LookupState ls = LookupState.Error; string lookupPath = null; object wbsL = null; string lookupDetail = "";
                try
                {
                    wbsL = scope.Track(Com.Get(app, "Workbooks"));
                    lookupPath = GetLoadedAddInPath(wbsL, Decisions.AddInName, scope, out ls, out lookupDetail);
                }
                catch (Exception ex) { ls = LookupState.Error; lookupDetail = ex.GetType().Name + ": " + ex.Message; }   // 连 Workbooks 都取不到（宿主正忙）→ 按无法确认处理
                finally { scope.Release(wbsL); }
                if (ls == LookupState.Error && _readOnly) _log.Log("[只读诊断] 查询详情: hwnd=" + h + " pid=" + winPid + " " + lookupDetail);
                bool samePath = ls == LookupState.Found && Decisions.SamePath(lookupPath, _opt.XlamPath);
                Presence presence = Decisions.GetPresenceDecision(true, ls, samePath);
                if (presence == Presence.InPlace)
                {
                    _seen.Add(key);
                    _log.Log("已在位，无需装载: hwnd=" + h + " pid=" + winPid);
                    return;
                }
                if (presence == Presence.Foreign)
                {
                    _seen.Add(key);
                    _log.Log("已在位，无需装载: hwnd=" + h + " pid=" + winPid + " 在位副本=" + lookupPath);
                    return;
                }
                if (presence == Presence.Unknown)
                {
                    // 宿主繁忙（如用户正在编辑单元格）会每个周期都走到这里：限频 60 秒一行。不标 seen，下周期重试
                    if (!_lastPresenceUnknownLog.HasValue || (now - _lastPresenceUnknownLog.Value).TotalSeconds >= UnknownLogSeconds)
                    {
                        _lastPresenceUnknownLog = now;
                        _log.Log("菜单在位但无法确认加载项装载状态（查询出错或路径不可读），本周期不处理: hwnd=" + h + " pid=" + winPid);
                    }
                    return;
                }
                // deadmenu：菜单在位但加载项工作簿明确不在位（命令栏缓存残留），按未装载处理。
                // 只读模式下窗口不会被标 seen，同一个窗口的这一行只记一次（PowerShell 版记完就装载并标 seen，不会重复）
                if (!_deadMenuLogged.Contains(key)) { _deadMenuLogged.Add(key); _log.Log("菜单在位但加载项未装载（命令栏缓存残留），按未装载处理: hwnd=" + h + " pid=" + winPid); }
                deadMenu = true;
            }

            // ---- 从这里起是写操作（删除死菜单控件 / 打开加载项 / 执行宏）。写之前复核目标窗口仍在 ----
            if (!Native.IsWindowAlive(h)) return;
            if (_readOnly)
            {
                // 切片 A：不执行任何写操作。同一个窗口、同一种判定只记一行，判定变了才再记
                string note = deadMenu
                    ? "[只读诊断] 判定：未装载（死菜单，装载前应先从外部删除残留控件）；本版不执行装载: hwnd=" + h + " pid=" + winPid
                    : "[只读诊断] 判定：未装载；本版不执行装载: hwnd=" + h + " pid=" + winPid;
                string last;
                if (!_lastReadOnlyNote.TryGetValue(key, out last) || !string.Equals(last, note, StringComparison.Ordinal))
                {
                    _lastReadOnlyNote[key] = note;
                    _log.Log(note);
                }
                return;
            }
            _log.Log("常驻装载尚未实现（切片 C）: hwnd=" + h + " pid=" + winPid);
        }

        private static int ToCount(object v)
        {
            if (v == null) throw new HostException("命令栏控件数读不到（宿主繁忙）");
            return Convert.ToInt32(v, CultureInfo.InvariantCulture);
        }

        /// <summary>右键菜单 Tag 是否在位（移植 Menu-Exists）。读不到 ≠ 不在位：控件数读不到直接抛；控件的 Tag 读不到先记一笔、继续扫描；
        /// 找到带标签的控件就是「在位」；没找到且有读不到的才是「无法确认」，抛给主循环走「附着失败（退避重试）」。</summary>
        private static bool MenuExists(object app, ComScope scope)
        {
            bool found = false; int unknown = 0;
            object bars = null, bar = null, ctrls = null, ctrl = null;
            try
            {
                bars = scope.Track(Com.Get(app, "CommandBars"));
                bar = scope.Track(Com.GetOrInvoke(bars, "Item", "Cell"));
                ctrls = scope.Track(Com.Get(bar, "Controls"));
                int n = ToCount(Com.Get(ctrls, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    ctrl = null;
                    try
                    {
                        ctrl = scope.Track(Com.GetOrInvoke(ctrls, "Item", i));
                        object tag;
                        try { tag = Com.Get(ctrl, "Tag"); }
                        catch (HostException) { tag = null; }      // 与 PS 5.1 一致：Tag 读取失败按「读不到」计
                        if (tag == null) unknown++;
                        else if (string.Equals(tag.ToString(), Decisions.MenuTag, StringComparison.Ordinal)) found = true;
                    }
                    finally { scope.Release(ctrl); }
                    if (found) break;
                }
                if (unknown > 0 && !found) throw new HostException("有 " + unknown + " 个控件的 Tag 读不到（宿主繁忙），且没有找到带标签的菜单项");
            }
            finally
            {
                scope.Release(ctrl); scope.Release(ctrls); scope.Release(bar); scope.Release(bars);
            }
            return found;
        }

        /// <summary>取「当前已在位的、名为 name 的加载项工作簿」的完整路径（按名字取：Excel 的 Workbooks 集合不按序号枚举加载项工作簿）。
        /// 在位且路径可读返回完整路径，否则 null；三态结论经 state 带出。调用方只能把 NotFound 当作「可以装载」。</summary>
        private static string GetLoadedAddInPath(object wbs, string name, ComScope scope, out LookupState state, out string detail)
        {
            state = LookupState.Error; object w = null; string p = null; detail = "";
            try
            {
                w = scope.Track(Com.GetOrInvoke(wbs, "Item", name));
                object full = Com.Get(w, "FullName");
                p = full == null ? "" : full.ToString();
                detail = "FullName=[" + p + "]";
                if (p.Length > 0) state = LookupState.Found; else p = null;
            }
            catch (HostException ex)
            {
                detail = "HRESULT=0x" + ex.InnerHResult.ToString("X8") + " " + ex.Message;
                if (ex.InnerHResult == unchecked((int)0x8002000B)) state = LookupState.NotFound;   // DISP_E_BADINDEX：明确不在位
                p = null;
            }
            finally { scope.Release(w); }
            return p;
        }
    }
}
