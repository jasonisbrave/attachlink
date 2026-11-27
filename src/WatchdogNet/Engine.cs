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
        private bool _windowFailed;                       // AttachWindow 得出「没能确认初始化」的结论 → 本窗口按退避重试
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

        /// <summary>常驻（计划任务经启动器启动时的模式）：无限循环，直到被结束（停止任务 / 注销 / 启动器被结束）。
        /// 启动行与 PowerShell 版相同（「看门狗启动, pid=…, xlam=…」）。-MaxCycles 仅供回归。</summary>
        public int RunResident()
        {
            _log.Log("看门狗启动, pid=" + Process.GetCurrentProcess().Id + ", xlam=" + _opt.XlamPath);
            while (true)
            {
                _cycle++;
                RunCycle();
                if (_opt.MaxCycles > 0 && _cycle >= _opt.MaxCycles) break;
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(_opt.IntervalSeconds));
            }
            return 0;
        }

        /// <summary>跑若干周期后退出（--diag 只读诊断、--once 手动装载）。启动行标明模式。</summary>
        public int Run(string modeLabel)
        {
            _log.Log("看门狗启动" + (modeLabel.Length > 0 ? "（" + modeLabel + "）" : "") + ", pid=" + Process.GetCurrentProcess().Id + ", xlam=" + _opt.XlamPath);
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
            if (_opt.OnlyPid > 0 && winPid != (uint)_opt.OnlyPid) return;      // --once -Pid：只处理指定进程
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
                _windowFailed = false;
                AttachWindow(h, winPid, key, app, winVisible, scope, now);
                attachFailed = _windowFailed;
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

            // ---- 从这里起是写操作（删除死菜单控件 / 打开加载项 / 执行宏 / 写回宿主设置）。写之前复核目标窗口仍在：
            // 宿主若已开始退出，文档窗口首当其冲被销毁，此时再动命令栏或 Open / Run 会打断退出流程 ----
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

            if (deadMenu)
            {
                // 死菜单根除：装载之前先从外部删掉残留控件（为什么必须从外部删见 RemoveDeadMenuControls）。
                // 删不掉不影响本周期装载——装载后的初始化仍会在会话内清掉死菜单，缓存留待下个实例再试。
                int deadRemoved = RemoveDeadMenuControls(app, scope);
                if (deadRemoved > 0) _log.Log("已删除死菜单控件 " + deadRemoved + " 个（实例正常退出后不再写入命令栏缓存）: hwnd=" + h + " pid=" + winPid);
                else if (deadRemoved == 0) _log.Log("未找到可删除的死菜单控件（期间已被其它方清理），继续装载: hwnd=" + h + " pid=" + winPid);
                else _log.Log("死菜单控件暂未删除（宿主正忙，命令栏不可访问），本周期仍照常装载，缓存留待下个实例再试: hwnd=" + h + " pid=" + winPid);
            }

            string provState, foreignPath;
            if (Provision(app, scope, out provState, out foreignPath))
            {
                Debug("Provision 结论=" + provState + "；目标窗口仍在=" + Native.IsWindowAlive(h));
                if (provState == "trusted" && !Native.IsWindowAlive(h))
                {
                    // 装载期间目标窗口已消失（宿主开始退出）：不执行宏、不标记 seen
                    _log.Log("装载后目标窗口已消失（宿主正在退出），不执行初始化: hwnd=" + h + " pid=" + winPid);
                }
                else if (provState == "trusted")
                {
                    // 只有确认在位的那份就是 $XlamPath 本身，才执行初始化（显式正向判断，不靠排除法）。
                    // 成功判定从严（v1.6.1）：初始化宏执行成功、菜单 Tag 在位、版本宏返回非空，三者齐备才标 seen。
                    bool initOk = false;
                    try { Com.Call(app, "Run", InitMacro); initOk = true; }
                    catch (Exception ex) { _log.Log("Run AttachLink_Initialize 失败: " + ex.Message); }
                    bool menuOk = false;
                    try { menuOk = MenuExists(app, scope); }
                    catch (Exception ex) { _log.Log("装载后菜单读不到（退避重试）: " + ex.Message); }
                    string ver = "";
                    if (initOk && menuOk) { try { object v = Com.Call(app, "Run", VersionMacro); ver = v == null ? "" : v.ToString(); } catch { ver = ""; } }
                    if (initOk && menuOk && ver.Length > 0)
                    {
                        _seen.Add(key);
                        _log.Log("已装载并初始化: hwnd=" + h + " pid=" + winPid + " 版本=v" + ver);
                    }
                    else
                    {
                        // 没能确认就不标 seen，按退避重试
                        _windowFailed = true;
                        _log.Log("装载后未能确认初始化（初始化" + (initOk ? "成功" : "失败") + "、菜单" + (menuOk ? "在位" : "不在位")
                            + "、版本" + (ver.Length > 0 ? ver : "读不到") + "），下轮重试: hwnd=" + h + " pid=" + winPid + "（退避重试）");
                    }
                }
                else if (provState == "unknown")
                {
                    // 无法确认在位情况（宿主繁忙 / 路径不可读）：不打开、不执行宏，也不标记 seen，下个周期重试
                    _log.Log("无法确认在位的 AttachLink.xlam（查询出错或路径不可读），本周期不处理: hwnd=" + h + " pid=" + winPid);
                }
                else
                {
                    // foreign：宿主里是一份异路径的同名副本；mismatch：打开后复核不符。都不打开第二份、不执行宏，对该窗口不再重试
                    _seen.Add(key);
                    if (provState == "mismatch") _log.Log("打开后复核不符（在位的不是看门狗配置的那份），不执行宏: hwnd=" + h + " pid=" + winPid + " 在位=" + foreignPath);
                    else _log.Log("同名异路径副本已装载且未初始化，不执行其宏、不打开第二份，跳过: hwnd=" + h + " pid=" + winPid + " 副本=" + foreignPath);
                }
            }
        }

        private const string InitMacro = "'AttachLink.xlam'!AttachLink_Initialize";       // 宏名限定到工作簿（v1.6.1）
        private const string VersionMacro = "'AttachLink.xlam'!AttachLink_GetVersion";

        private sealed class HostUiState { public bool? Alerts; public bool? Screen; }

        /// <summary>宿主两项界面设置的原值快照（装载前调用）。读不到的项保持 null，恢复时跳过。</summary>
        private static HostUiState GetHostUiState(object app)
        {
            HostUiState s = new HostUiState();
            try { object v = Com.Get(app, "DisplayAlerts"); if (v != null) s.Alerts = Convert.ToBoolean(v, CultureInfo.InvariantCulture); } catch { }
            try { object v = Com.Get(app, "ScreenUpdating"); if (v != null) s.Screen = Convert.ToBoolean(v, CultureInfo.InvariantCulture); } catch { }
            return s;
        }

        /// <summary>把宿主的 DisplayAlerts / ScreenUpdating 恢复成装载前的值；某项没读到时不恢复该项。直接写回、不先比较。</summary>
        private static void RestoreHostUiState(object app, HostUiState state)
        {
            if (state == null) return;
            try { if (state.Alerts.HasValue) Com.Set(app, "DisplayAlerts", state.Alerts.Value); } catch { }
            try { if (state.Screen.HasValue) Com.Set(app, "ScreenUpdating", state.Screen.Value); } catch { }
        }

        /// <summary>只装载、不登记（v1.6.0，R1）——移植 Provision-App。返回 true = 已得出确定结论（provState：trusted / foreign / mismatch / unknown），
        /// false = 打开过程出错（已记日志，provState = failed）。只有 trusted 允许执行初始化宏。</summary>
        private bool Provision(object app, ComScope scope, out string provState, out string foreignPath)
        {
            object wbs = null, wb = null; HostUiState uiState = null;
            provState = "failed"; foreignPath = "";
            bool done = false;
            try
            {
                wbs = scope.Track(Com.Get(app, "Workbooks"));
                LookupState ls; string detail;
                string loaded = GetLoadedAddInPath(wbs, Decisions.AddInName, scope, out ls, out detail);
                if (ls == LookupState.Error) provState = "unknown";
                else if (ls == LookupState.Found)
                {
                    // 已有同名加载项在位：不打开第二份，只判定它是不是配置的那一份
                    if (Decisions.SamePath(loaded, _opt.XlamPath)) provState = "trusted";
                    else { provState = "foreign"; foreignPath = loaded; }
                }
                else
                {
                    if (_opt.DebugLog) Debug("三态查询=" + ls + " " + detail + "；宿主设置 " + DescribeHostUi(app));
                    // 明确不在位：从配置路径打开（只装载、不登记）
                    uiState = GetHostUiState(app);
                    System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
                    wb = scope.Track(Com.Call(wbs, "Open", _opt.XlamPath));
                    Debug("Workbooks.Open 返回，耗时 " + sw.ElapsedMilliseconds + " ms");
                    // 打开后立即复核：在位的那份必须就是配置的那一份（检查与打开之间宿主可能已装入同名副本）
                    string opened = "";
                    try { object f = Com.Get(wb, "FullName"); opened = f == null ? "" : f.ToString(); } catch (HostException ex) { Debug("FullName 读不到: " + ex.Message); opened = ""; }
                    Debug("打开后在位副本=[" + opened + "] 配置路径=[" + _opt.XlamPath + "] 同一份=" + Decisions.SamePath(opened, _opt.XlamPath));
                    if (Decisions.SamePath(opened, _opt.XlamPath)) provState = "trusted";
                    else { provState = "mismatch"; foreignPath = opened; }
                }
                done = true;
            }
            catch (Exception ex)
            {
                _log.Log("装载失败（Workbooks.Open）: " + ex.Message);
            }
            finally
            {
                // 恢复放在 finally：复核不符要恢复，Workbooks.Open 中途抛错（事件可能已经跑完）同样要恢复；没走到打开这一步则不碰
                if (uiState != null) { RestoreHostUiState(app, uiState); Debug("宿主设置已写回"); }
                scope.Release(wb); scope.Release(wbs);
            }
            return done;
        }

        private void Debug(string message)
        {
            if (_opt.DebugLog) _log.Log("[调试] " + message);
        }

        private static string DescribeHostUi(object app)
        {
            HostUiState s = GetHostUiState(app);
            return "DisplayAlerts=" + (s.Alerts.HasValue ? s.Alerts.Value.ToString() : "读不到") + " ScreenUpdating=" + (s.Screen.HasValue ? s.Screen.Value.ToString() : "读不到");
        }

        /// <summary>死菜单根除（v1.6.0 R2）：从外部删除 Cell 命令栏上带 AttachLinkCtx 标签的控件，返回删掉的个数；
        /// -1 = 一个都没删成、且存在读不到 / 删不掉的控件（调用方按「暂未删除、稍后重试」处理）。</summary>
        private static int RemoveDeadMenuControls(object app, ComScope scope)
        {
            int removed = 0, failed = 0;
            object bars = null, bar = null, ctrls = null, ctrl = null;
            try
            {
                bars = scope.Track(Com.Get(app, "CommandBars"));
                bar = scope.Track(Com.GetOrInvoke(bars, "Item", "Cell"));
                ctrls = scope.Track(Com.Get(bar, "Controls"));
                object nRaw = Com.Get(ctrls, "Count");
                if (nRaw == null) removed = -1;            // 控件数读不到：显式按「不可访问」处理，不能当成 0 个控件
                else
                {
                    int n = Convert.ToInt32(nRaw, CultureInfo.InvariantCulture);
                    for (int i = n; i >= 1; i--)           // 倒序：删除不改变尚未遍历部分的下标
                    {
                        ctrl = null;
                        try
                        {
                            ctrl = scope.Track(Com.GetOrInvoke(ctrls, "Item", i));
                            object tag = Com.Get(ctrl, "Tag");
                            if (tag == null) failed++;      // 读不到的那个可能就是死菜单，不能当成「无标签」放过
                            else if (string.Equals(tag.ToString(), Decisions.MenuTag, StringComparison.Ordinal)) { Com.Call(ctrl, "Delete"); removed++; }
                        }
                        catch (Exception) { failed++; }     // 取不到该控件，或读 Tag / 删除时抛异常（宿主繁忙）：跳过它
                        finally { scope.Release(ctrl); }
                    }
                }
            }
            catch (Exception) { removed = -1; }
            finally { scope.Release(ctrl); scope.Release(ctrls); scope.Release(bar); scope.Release(bars); }
            if (removed == 0 && failed > 0) return -1;
            return removed;
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
