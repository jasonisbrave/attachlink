# =====================================================================
# AttachLink 看门狗（Watchdog）v1.0
# ---------------------------------------------------------------------
# 背景：鼎信诺等审计软件通过 COM 自动化（CreateObject）拉起 Excel 时，
# Excel 跳过一切加载项初始化（.xlam / XLSTART / COM 加载项的 OnConnection
# 均不触发，本机逐项实测证实）。因此装载必须来自 Excel 进程外部。
#
# 机制：周期性枚举 Excel 的 EXCEL7 文档窗口（XLMAIN 的子窗口），通过
# oleacc!AccessibleObjectFromWindow(OBJID_NATIVEOM, IID_IDispatch) 取得
# 该窗口的 Excel Window 对象 → .Application → Workbooks.Open(xlam) 强制装载
# → 显式 Run 初始化 → 验证右键菜单 Tag 在位。该通道是微软文档化的
# Office Native OM 附着方式，不依赖 ROT、不注入。
#
# 时序说明：零工作簿实例没有 EXCEL7 窗口，自然不会附着；鼎信诺打开底稿
# （出现工作簿窗口）后下个周期自动完成装载——时机恰好。
#
# 宿主身份校验（v1.4.3）：WPS 表格为兼容自动化同样使用 XLMAIN/EXCEL7 窗口类名
# （COM 自动化实例与独立窗口模式下是顶层窗口，与 Excel 无法按类名区分）。仅凭
# 类名命中即附着，会把正在打开文档的 WPS 卡死，或反复弹「同名文件已打开」
# （2026-11-23 同事机日志 + 本机复现）。因此枚举到 EXCEL7 后先查其所属进程：
# 进程名不是 EXCEL 的窗口一律跳过，不向它发送任何 COM 调用（含
# AccessibleObjectFromWindow）。AttachLink 在 WPS 中的可用性走 WPS 自身的
# 加载项装载通道，与看门狗无关。
#
# 引用释放（v1.4.3）：每次附着结束立即 ReleaseComObject + 强制 GC。宿主（Excel / WPS）
# 在 Quit 之后要等全部外部 COM 引用释放才真正退出——旧版把取到的 Application 与枚举
# 过程产生的引用一直持有到下次附着或 GC，被附着过且已 Quit 的 Excel 实例因此被钉在
# 内存里不退（2026-11-23 实测；疑与「关了 Excel 仍提示在运行」「重开连回旧进程」
# 类现象同源，见 WORKLOG #27）。
#
# 已在位即不动（v1.4.3）：取得 Application 后先看右键菜单是否已在位——在位说明加载项
# 已随宿主正常启动装载（OPENn），直接标记完成，不再 Open / Run。安装器默认部署看门狗
# 之后，这是普通双击打开的 Excel 会话的常态。
#
# 对短命的隐藏自动化实例保持克制（v1.4.3）：不可见实例首次出现时宽限一个周期再处理；宿主里
# 已有同名加载项工作簿时不再打开第二份。各类工具 / 测试脚本常自己装载另一份同名加载项，
# 看门狗抢先装载会让两份副本相撞（2026-11-23 实测：后打开的被当作只读副本，宿主随后退不干净）。
#
# 只对确认可信的那一份执行宏（v1.4.3）：宿主里在位的副本是否就是本脚本配置的那一份（-XlamPath），
# 分 trusted / foreign / mismatch / unknown 四态判定，只有 trusted 才 Run 初始化；查询出错
# （宿主繁忙、路径读不到）一律按「无法确认」处理，不打开、不执行（见 Provision-App）。
#
# 隐藏的文档窗口不作为处理对象（v1.4.3）：每个加载项都有一个隐藏的文档窗口，被隐藏的工作簿也是。
# 它们不是用户要右键的工作窗口。旧逻辑把它们当新窗口处理：宿主开始退出后（外部客户端还持着引用、
# 进程迟迟不退），看门狗可能对加载项自身的窗口再执行一次「打开 + 初始化」——这是一条可疑的致残留路径
# （2026-11-24 排查，未能定论，见 WORKLOG #29），本版予以规避。是否隐藏以 Excel 报告的 Window.Visible
# 为准——不能用 Win32 可见性：实例整体不可见时，工作簿窗口在 Win32 层面同样「不可见」（2026-11-24 实测）。
# Window.Visible 读不到时只延迟重试（既不按可见、也不按隐藏处理）；隐藏窗口宽限一个周期后标记跳过（用户的
# 工作窗口在创建瞬间可能尚未显示，立刻判定会让这个实例永远不被装载）；已判定的窗口只在 Win32「由不可见变为
# 可见」时重新评估（次数有上限），Win32 持续可见却没有新跳变的按低频只读重查（不设上限）——任何路径都不会
# 让一个窗口被永久跳过。每次写操作（打开加载项 / 执行宏）之前都复核目标窗口仍然存在。
#
# 部署：install-elevated.ps1（管理员）将本脚本复制到
# C:\Program Files\AttachLink\ 并注册为登录计划任务（最高权限运行，
# 可同时服务普通 Excel 与管理员权限的鼎信诺 Excel）。
# =====================================================================
param(
    [string]$XlamPath = (Join-Path $env:APPDATA 'Microsoft\AddIns\AttachLink.xlam'),
    [int]$IntervalSeconds = 2,
    [int]$MaxCycles = 0,
    # 父进程白名单（逗号分隔进程名，如 "sjp7.exe"）：仅附着由这些进程拉起的 Excel。
    # 空 = 不过滤（默认，保持既有行为）。未指定参数时读取脚本同目录 watchdog.conf
    # 的 ParentAllowlist= 行（install-elevated.ps1 -WatchdogParentFilter 部署时写入）。
    [string]$ParentAllowlist = '',
    # 周期行心跳间隔（秒）：仅当 EXCEL7 窗口数变化、或距上次记周期行超过该间隔时，
    # 才记一行「cycle N: EXCEL7 窗口数=X」。旧版每周期必记，常驻一天 4 万余行
    # （2026-11-23 实测 3 天 3.3MB，排障时事件行被淹没）。0 = 恢复每周期必记。
    [int]$HeartbeatSeconds = 30,
    # 日志路径。默认 %TEMP%\AttachLinkWatchdog.log；回归测试传入独立路径，
    # 避免与机器级常驻看门狗的日志交错。
    [string]$LogFile = ''
)

$ErrorActionPreference = 'Continue'
$logFile = if ($LogFile) { $LogFile } else { Join-Path $env:TEMP 'AttachLinkWatchdog.log' }

function WdLog([string]$m) {
    # 时间戳带日期：日志长期追加，只有时分秒的话隔天同一时刻的行无法区分（排障时按时刻截取会混入前几天的记录）
    try { [IO.File]::AppendAllText($logFile, (Get-Date -Format 'MM-dd HH:mm:ss') + ' ' + $m + "`r`n") } catch {}
}

# 生效的父进程过滤：参数优先，其次配置文件
$effectiveFilter = $ParentAllowlist
if (-not $effectiveFilter) {
    $conf = Join-Path $PSScriptRoot 'watchdog.conf'
    if (Test-Path $conf) {
        $line = (Get-Content $conf -ErrorAction SilentlyContinue | Where-Object { $_ -match '^\s*ParentAllowlist\s*=' }) | Select-Object -First 1
        if ($line) { $effectiveFilter = ($line -replace '^\s*ParentAllowlist\s*=', '').Trim() }
    }
}
$filterSet = @()
if ($effectiveFilter) {
    $filterSet = ($effectiveFilter -split ',') | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ }
    WdLog ("父进程白名单生效: " + ($filterSet -join ', '))
}

# 编码要求：本文件必须保存为 UTF-8 with BOM（PS 5.1 对无 BOM 文件按 ANSI/GBK
# 误读，曾导致 C# 注释乱码吞换行、方法签名被并入注释行、编译连锁失败）。
if (-not ('AttachLinkWatchdog.Native' -as [type])) {
    try {
        Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace AttachLinkWatchdog
{
    public class Native
    {
        private delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lp);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc cb, IntPtr lp);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr h, StringBuilder sb, int max);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr h);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr h);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr h, uint dwId, ref Guid riid,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);

        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;

        private static bool IsClass(IntPtr h, string cls)
        {
            var sb = new StringBuilder(64);
            GetClassName(h, sb, 64);
            return sb.ToString() == cls;
        }

        // EXCEL7 窗口所属进程 PID（父进程白名单过滤用）
        public static long GetWindowProcessId(long hwnd)
        {
            uint pid;
            GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
            return pid;
        }

        // 枚举 XLMAIN 顶层窗口下的 EXCEL7 文档窗口句柄（含隐藏的；是否处理由主循环按 Window.Visible 决定）
        public static List<long> FindExcel7Windows()
        {
            var list = new List<long>();
            EnumWindows((top, lp) =>
            {
                if (IsClass(top, "XLMAIN"))
                {
                    EnumChildWindows(top, (child, lp2) =>
                    {
                        if (IsClass(child, "EXCEL7")) list.Add(child.ToInt64());
                        return true;
                    }, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
            return list;
        }

        // 目标窗口是否仍然存在（写操作前的复核用）
        public static bool IsWindowAlive(long hwnd) { return IsWindow(new IntPtr(hwnd)); }

        // 窗口自身及其祖先是否带 WS_VISIBLE（纯 Win32 查询，不涉及宿主）。只用来粗判「已判定为隐藏的窗口
        // 日后是否被显示了」——显示了就重新评估一次（是否真可见仍以 Window.Visible 为准）。注意实例整体不可见时
        // 所有文档窗口在 Win32 层面都不可见，所以它只对可见实例里的窗口有意义。
        public static bool IsWindowShown(long hwnd) { return IsWindowVisible(new IntPtr(hwnd)); }

        // 从 EXCEL7 窗口取 Excel Window 对象，再返回其 Application；同时带出该窗口自身的可见性
        // （1 / 0，读不到为 -1）。加载项自身的文档窗口 Visible=0，用户的工作窗口 Visible=1
        //（即使实例整体不可见也为 1，2026-11-24 实测）。
        // （用反射而非 dynamic：Add-Type 默认编译环境不含 Microsoft.CSharp）
        public static object GetApplicationFromWindow(long hwnd, out int windowVisible)
        {
            object win;
            var iid = IID_IDispatch;
            windowVisible = -1;
            int hr = AccessibleObjectFromWindow(new IntPtr(hwnd), OBJID_NATIVEOM, ref iid, out win);
            if (hr != 0 || win == null) return null;
            try
            {
                try
                {
                    windowVisible = Convert.ToInt32(win.GetType().InvokeMember("Visible",
                        System.Reflection.BindingFlags.GetProperty, null, win, null));
                }
                catch { }
                return win.GetType().InvokeMember("Application",
                    System.Reflection.BindingFlags.GetProperty, null, win, null);
            }
            finally
            {
                // Window 对象用完即释放，不留给 GC（原因见脚本里的 Release-Com）
                System.Runtime.InteropServices.Marshal.ReleaseComObject(win);
            }
        }
    }
}
'@
        WdLog ("Add-Type 完成, 类型可用=" + [bool]('AttachLinkWatchdog.Native' -as [type]))
    } catch {
        # 完整异常链 + LoaderExceptions（ReflectionTypeLoadException 时逐条列出缺失类型）
        WdLog ("Add-Type 失败: " + $_.Exception.GetType().FullName + " :: " + $_.Exception.Message)
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException; WdLog ("  inner: " + $ex.GetType().FullName + " :: " + $ex.Message) }
        if ($_.Exception.LoaderExceptions) {
            foreach ($le in ($_.Exception.LoaderExceptions | Select-Object -First 8)) { WdLog ("  loader: " + $le.Message) }
        }
        WdLog "致命错误：看门狗退出（类型不可用，后续周期必然全部失败）"
        exit 1
    }
}

function Release-Com($obj) {
    # 立即释放 RCW 持有的全部 COM 引用（循环到引用计数归零）。宿主在 Quit 之后要等外部 COM 引用全部
    # 释放才真正退出，而「不可达对象交给 GC 回收」在常驻进程里并不总是可靠（WORKLOG #27 会诊实测：看门狗
    # 先运行一段时间再附着时，一次 GC 回收不掉某些链式调用产生的中间对象，Excel 被钉在内存里不退）。
    # 因此看门狗碰过的每个 COM 对象都必须在这里显式放掉；非 COM 对象 / 已释放对象静默跳过。
    if ($null -eq $obj) { return }
    try { while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) -gt 0) { } } catch {}
}

function Find-AddInIndex($addins) {
    # 在 AddIns 集合里按 Name 定位 AttachLink.xlam（返回 1 起的下标，找不到返回 0）。
    # 用下标遍历而非 foreach：不产生需要回收的枚举器；每个 AddIn 对象查完即释放——释放放在 finally 里，
    # 属性读取抛异常（宿主繁忙 / RPC 拒绝）时同样不漏。
    $n = [int]$addins.Count
    for ($i = 1; $i -le $n; $i++) {
        $ai = $null
        try {
            $ai = $addins.Item($i)
            if ($ai.Name -ieq 'AttachLink.xlam') { return $i }
        } finally { Release-Com $ai }
    }
    return 0
}

function Get-LoadedAddInPath($wbs, $name) {
    # 取「当前已在位的、名为 $name 的加载项工作簿」的完整路径，读完立即释放。
    # 必须按名字取：Excel 的 Workbooks 集合不按序号枚举加载项工作簿（Count 不含它们，1..Count 遍历不到），
    # 但 Item('名字') 能取到——按序号查找会永远「找不到」，导致看门狗在已有同名副本时仍打开第二份（2026-11-23 实测）。
    #
    # 返回值：在位且路径可读时为完整路径，否则 $null。结果性质写入 $script:lookupState（三态）：
    #   found    = 在位，路径读取成功
    #   notfound = 明确不在位（Excel 对按名字取不到的工作簿抛 DISP_E_BADINDEX，HRESULT 0x8002000B「无效索引」）
    #   error    = 其他一切情况：在位但路径读不到、宿主繁忙 / RPC 拒绝、对象不可用……
    # 调用方只能把 notfound 当作「可以装载」；error 表示「无法确认」，绝不能当成不在位去打开第二份。
    $script:lookupState = 'error'
    $w = $null; $p = $null
    try {
        $w = $wbs.Item($name)
        $p = [string]$w.FullName
        if ($p) { $script:lookupState = 'found' } else { $p = $null }
    } catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        if ($ex.HResult -eq [int]0x8002000B) { $script:lookupState = 'notfound' }
        $p = $null
    } finally { Release-Com $w }
    return $p
}

function Test-SamePath($a, $b) {
    try { return ([IO.Path]::GetFullPath($a) -ieq [IO.Path]::GetFullPath($b)) } catch { return $false }
}

# ---- 隐藏文档窗口的状态机 ----
# 判定逻辑集中在这两个函数里，主循环只负责调用与记日志；它们不碰 Excel，由 scripts\test-watchdog-units.ps1
# 的 A2 段用假状态 + 受控时钟逐条回归（涉及「永久跳过」的路径必须有用例）。

function New-WindowState {
    # 某个文档窗口在「隐藏判定」上的状态；调用方按窗口键保存在 $hiddenWin 里，窗口消失即丢弃。
    return @{ since = $null; done = $false; rev = 0; shown = $false; lastCheck = $null; logged = $false; unknownSince = $null; unknownEscalated = $false }
}

function Get-HiddenDecision($st, [int]$winVisible, [bool]$shownNow, $now, [int]$intervalSeconds) {
    # 对一个尚未标记 seen 的窗口，按 Excel 报告的 Window.Visible（1 / 0，-1 = 读不到）决定本周期怎么做。
    # 返回：process = 按工作窗口处理；wait = 本周期不处理、不标 seen、下周期再看；skip = 标记 seen，不作为处理对象。
    #
    # 读不到按「未知」处理：只延迟、重试。不能按可见处理——宿主退出过程中，加载项自身的窗口正是这时候不该碰的
    # 对象；也不能按隐藏处理——实例整体不可见时，Win32 可见性帮不上忙（工作窗口在 Win32 层面同样不可见），
    # 判定下去会让正常的工作窗口被永久跳过，隐藏的自动化实例就再也装载不了。
    if ($winVisible -lt 0) { return 'wait' }
    $st.unknownSince = $null; $st.unknownEscalated = $false
    if ($winVisible -ne 0) { return 'process' }
    # Excel 报告为隐藏：宽限一个周期再判定（用户的工作窗口在创建瞬间可能尚未显示，立刻判定会让这个实例永远
    # 不被装载）。重新评估时 since 已由 Test-HiddenReevaluate 置为触发时刻，同样再等一个周期。
    if ($null -eq $st.since) { $st.since = $now; return 'wait' }
    if ((($now) - ([datetime]$st.since)).TotalSeconds -lt $intervalSeconds) { return 'wait' }
    $st.done = $true
    $st.shown = $shownNow          # 记下判定时的 Win32 可见性：之后只有「由不可见变为可见」才算一次新的跳变
    $st.lastCheck = $now
    return 'skip'
}

function Test-HiddenReevaluate($st, [bool]$shownNow, $now, [int]$maxReeval, [int]$recheckSeconds) {
    # 对一个已标记 seen、已判定为隐藏的窗口，每个周期只用 Win32 可见性（不向宿主发调用）决定是否重新评估。
    # 触发条件二选一：
    #   ① 真实的「不可见 → 可见」跳变（用户取消隐藏、工具先隐藏后显示、实例由不可见变为可见）——立即触发，
    #      次数用完为止（$maxReeval）。计数只在跳变时消耗：Win32 持续可见而没有新跳变时不消耗。
    #   ② 窗口 Win32 可见，且距上次评估已超过 $recheckSeconds——低频只读重查，不设次数上限。
    #      「Win32 可见、Excel 却报告隐藏」这种不一致不能把窗口永久跳过。
    # Win32 一直不可见的已判定窗口（加载项自身的窗口、被隐藏的工作簿、不可见实例里的隐藏窗口）不会被重查：
    # 它们本来就不是处理对象；等它们显示出来（跳变）再说。
    # 返回 $true 时已把状态复位为本轮待评估（done = false、since = 触发时刻），调用方去掉 seen 并记一行日志。
    $turned = ($shownNow -and -not $st.shown)
    $st.shown = $shownNow
    if (-not $shownNow) { return $false }
    $sinceCheck = [double]::MaxValue
    if ($null -ne $st.lastCheck) { $sinceCheck = (($now) - ([datetime]$st.lastCheck)).TotalSeconds }
    if (-not (($turned -and $st.rev -lt $maxReeval) -or ($sinceCheck -ge $recheckSeconds))) { return $false }
    if ($turned -and $st.rev -lt $maxReeval) { $st.rev = $st.rev + 1 }     # 计数不越过上限：达到上限后的跳变只经由低频重查触发
    $st.lastCheck = $now
    $st.done = $false
    $st.since = $now
    return $true
}

function Provision-App($app) {
    # 登记装载（幂等）。注意：AddIns 集合的 Installed=True 只反映登记状态，
    # 自动化实例可能「已登记勾选但从未装载」（本机实测根因）。真正的强制
    # 装载靠下方 Workbooks.Open：xlam 为 IsAddin 工作簿，Open 无可见窗口、
    # 不切换用户活动窗口；已装载时 Open 返回现有工作簿，无副作用。
    #
    # 返回 $true = 已得出确定结论，结论见 $script:provState；$false = 登记或打开过程出错（已记日志）。
    #   trusted  = 宿主里在位的那份就是 $XlamPath（本周期由看门狗打开并复核，或此前已装载）——只有这个状态允许执行初始化宏
    #   foreign  = 宿主里在位的是一份【异路径】的同名副本 → 不打开第二份（两份同名副本相撞会让宿主退不干净），
    #              也绝不执行它的宏：那份可能来自非受信任位置或是旧版本，而看门狗服务的是管理员权限的 Excel
    #              （提权进程不得执行用户可写目录里的代码）
    #   mismatch = 看门狗执行了 Open，但打开后读到的在位副本路径与 $XlamPath 不符（期间宿主装入了同名副本 /
    #              Open 返回了别的工作簿）→ 不执行宏
    #   unknown  = 无法确认在位情况（查询出错或路径读不到）→ 不打开、不执行宏；调用方不标记 seen，下个周期重试
    #   failed   = 登记或打开时出错
    $addins = $null; $entry = $null; $added = $null; $wbs = $null; $wb = $null
    $script:provState = 'failed'; $script:foreignPath = ''
    $done = $false
    try {
        $wbs = $app.Workbooks
        $loaded = Get-LoadedAddInPath $wbs 'AttachLink.xlam'
        if ($script:lookupState -eq 'error') {
            $script:provState = 'unknown'
        } elseif ($script:lookupState -eq 'found') {
            # 已有同名加载项在位：不登记、不打开第二份，只判定它是不是 $XlamPath 本身
            if (Test-SamePath $loaded $XlamPath) { $script:provState = 'trusted' }
            else { $script:provState = 'foreign'; $script:foreignPath = $loaded }
        } else {
            # 明确不在位：登记，并从 $XlamPath 打开
            $addins = $app.AddIns
            $idx = Find-AddInIndex $addins
            if ($idx -gt 0) { $entry = $addins.Item($idx) }
            if ($null -eq $entry) {
                $added = $addins.Add($XlamPath, $false)
                $idx = Find-AddInIndex $addins
                if ($idx -gt 0) { $entry = $addins.Item($idx) } else { $entry = $added; $added = $null }
            }
            # 只对「登记的就是 $XlamPath 本身」的条目置 Installed=True：按名字找到的条目可能登记在别的路径，
            # 置位会让 Excel 从那个路径装载一份未经核验的副本。读不到登记路径时同样不置位。
            if ($null -ne $entry) {
                $entryPath = ''
                try { $entryPath = [string]$entry.FullName } catch { $entryPath = '' }
                $entryInstalled = $true
                try { $entryInstalled = [bool]$entry.Installed } catch { $entryInstalled = $true }
                if ((Test-SamePath $entryPath $XlamPath) -and (-not $entryInstalled)) { $entry.Installed = $true }
            }
            $wb = $wbs.Open($XlamPath)
            # 打开后立即复核：在位的那份必须就是 $XlamPath（检查与打开之间宿主可能已装入同名副本）
            $opened = ''
            try { $opened = [string]$wb.FullName } catch { $opened = '' }
            if (Test-SamePath $opened $XlamPath) { $script:provState = 'trusted' }
            else { $script:provState = 'mismatch'; $script:foreignPath = $opened }
        }
        $done = $true
    } catch {
        WdLog ("装载失败（AddIns 登记 / Workbooks.Open）: " + $_.Exception.Message)
    } finally {
        Release-Com $wb; Release-Com $wbs; Release-Com $added; Release-Com $entry; Release-Com $addins
    }
    return $done
}

function Menu-Exists($app) {
    # 右键菜单 Tag 是否在位。下标遍历 + 逐个释放（见 Find-AddInIndex 的说明）；任何异常原样抛给主循环。
    # 注意：这只是「已初始化」的幂等标记，不是加载项健康检查。
    $found = $false
    $bars = $null; $bar = $null; $ctrls = $null; $ctrl = $null
    try {
        $bars = $app.CommandBars
        $bar = $bars.Item('Cell')
        $ctrls = $bar.Controls
        $n = [int]$ctrls.Count
        for ($i = 1; $i -le $n; $i++) {
            $ctrl = $null
            try {
                $ctrl = $ctrls.Item($i)
                if ($ctrl.Tag -eq 'AttachLinkCtx') { $found = $true }
            } finally { Release-Com $ctrl }
            if ($found) { break }
        }
    } finally {
        Release-Com $ctrl; Release-Com $ctrls; Release-Com $bar; Release-Com $bars
    }
    return $found
}

WdLog ("看门狗启动, pid=" + $PID + ", xlam=" + $XlamPath)      # 带上自身 pid：常驻与被测看门狗可能共用同一份日志，排障时要能区分
$seen = @{}
$cycle = 0
$lastCount = -1               # 上次记入日志的 EXCEL7 窗口数（变化才记周期行）
$lastBeat = Get-Date          # 上次记周期行/心跳的时刻
$xlamMissing = $false         # xlam 缺失状态（仅状态切换时记日志，另配心跳）
$hiddenSince = @{}            # 隐藏实例的首次见到时刻（pid → 时间）：宽限一个周期后再处理
$hiddenWin = @{}              # 文档窗口在「隐藏判定」上的状态（窗口键 → New-WindowState），逻辑见 Get-HiddenDecision / Test-HiddenReevaluate
$maxReeval = 3                # 已判定为隐藏的窗口因「Win32 由不可见变为可见」而立即重新评估的次数上限
$hiddenRecheckSeconds = 30    # Win32 可见却没有新跳变时的低频只读重查间隔（秒）；不设次数上限，避免窗口被永久跳过

while ($true) {
    $cycle++
    try {
        if (-not [IO.File]::Exists($XlamPath)) {
            # xlam 缺失（升级/卸载进行中、文件被移走）：不向任何宿主发 COM 调用，静默等待。
            # 旧版此时每 2 秒向宿主重试一次 Workbooks.Open——同事机日志 78 秒内连续
            # 39 次失败，且每次都打进正在打开文档的 WPS。
            if (-not $xlamMissing) {
                $xlamMissing = $true
                $lastBeat = Get-Date
                WdLog ("xlam 不存在，暂停附着直至文件就位: " + $XlamPath)
            } elseif (((Get-Date) - $lastBeat).TotalSeconds -ge $HeartbeatSeconds) {
                $lastBeat = Get-Date
                WdLog ("cycle " + $cycle + ": xlam 仍不存在，继续等待")
            }
        } else {
            if ($xlamMissing) {
                $xlamMissing = $false
                $lastCount = -1
                WdLog "xlam 已就位，恢复附着"
            }
            $wins = [AttachLinkWatchdog.Native]::FindExcel7Windows()
            if ($wins.Count -ne $lastCount -or ((Get-Date) - $lastBeat).TotalSeconds -ge $HeartbeatSeconds) {
                $lastCount = $wins.Count
                $lastBeat = Get-Date
                WdLog ("cycle " + $cycle + ": EXCEL7 窗口数=" + $wins.Count)
            }
            foreach ($h in $wins) {
                $winPid = [AttachLinkWatchdog.Native]::GetWindowProcessId($h)
                $key = [string]$h + ':' + $winPid
                if ($seen.ContainsKey($key)) {
                    # 已判定为「隐藏窗口、不作处理对象」的窗口：每个周期只做一次 Win32 可见性查询来决定是否重新评估
                    #（判定在 Test-HiddenReevaluate 里），不向宿主发调用。
                    $st0 = $hiddenWin[$key]
                    if ($null -ne $st0 -and $st0.done -and (Test-HiddenReevaluate $st0 ([AttachLinkWatchdog.Native]::IsWindowShown($h)) (Get-Date) $maxReeval $hiddenRecheckSeconds)) {
                        $seen.Remove($key)
                        WdLog ("先前隐藏的窗口已显示，重新评估: hwnd=" + $h + " pid=" + $winPid)
                    } else { continue }
                }

                # ---- 宿主身份校验：只附着 Microsoft Excel（EXCEL.EXE）。必须在取得
                # Application 对象之前完成，识别不了进程名时同样跳过（fail-closed）----
                $hostName = ''
                try { $hostName = [System.Diagnostics.Process]::GetProcessById([int]$winPid).ProcessName } catch {}
                if ($hostName -ine 'EXCEL') {
                    $seen[$key] = $true
                    if ($hostName) { $hostName = $hostName + '.exe' } else { $hostName = '(进程已退出/不可识别)' }
                    WdLog ("跳过 hwnd=" + $h + "：宿主进程=" + $hostName + "，仅附着 EXCEL.EXE")
                    continue
                }

                $app = $null
                try {
                    $winVisible = [int]-1
                    $app = [AttachLinkWatchdog.Native]::GetApplicationFromWindow($h, [ref]$winVisible)
                    if ($null -eq $app) { WdLog ("hwnd=" + $h + " 未取得 Application"); continue }

                    # ---- 隐藏 / 状态未知的文档窗口（判定在 Get-HiddenDecision 里）----
                    # 绝大多数隐藏窗口是加载项自身的窗口，此外是被隐藏的工作簿；它们都不是用户要右键的工作窗口。
                    # 宿主开始退出后若外部客户端还持着引用，进程会拖上几秒，这期间对加载项自身的窗口再执行一次
                    # 「打开 + 初始化」是一条可疑的致残留路径（2026-11-24 排查，见 WORKLOG #29），本版予以规避。
                    $st = $hiddenWin[$key]
                    if ($null -eq $st) { $st = New-WindowState; $hiddenWin[$key] = $st }
                    $decision = Get-HiddenDecision $st $winVisible ([AttachLinkWatchdog.Native]::IsWindowShown($h)) (Get-Date) $IntervalSeconds
                    if ($decision -eq 'wait') {
                        # 未知（Window.Visible 读不到）只延迟、重试，不标 seen。首次出现记一行，持续 60 秒再记一行，便于排查。
                        if ($winVisible -lt 0) {
                            if ($null -eq $st.unknownSince) {
                                $st.unknownSince = Get-Date
                                WdLog ("Window.Visible 读不到，本周期不处理（下周期重试）: hwnd=" + $h + " pid=" + $winPid)
                            } elseif (-not $st.unknownEscalated -and ((Get-Date) - $st.unknownSince).TotalSeconds -ge 60) {
                                $st.unknownEscalated = $true
                                WdLog ("Window.Visible 已持续 60 秒读不到，该窗口一直没有被处理: hwnd=" + $h + " pid=" + $winPid)
                            }
                        }
                        continue
                    }
                    if ($decision -eq 'skip') {
                        $seen[$key] = $true
                        if (-not $st.logged) { $st.logged = $true; WdLog ("隐藏窗口（加载项自身或被隐藏的工作簿），不作为处理对象: hwnd=" + $h + " pid=" + $winPid) }
                        continue
                    }
                    $hiddenWin.Remove($key)      # Excel 报告为可见：按工作窗口处理

                    # ---- 隐藏实例宽限一个周期 ----
                    # 不可见的自动化实例（各类工具、测试脚本）往往寿命很短，或者会自己装载另一份同名加载项；
                    # 看门狗抢在前面装载，轻则白做，重则两份同名副本相撞、宿主退不干净（2026-11-23 实测）。
                    # 所以隐藏实例首次出现时先只记时间，下一个周期仍在才处理。可见实例（用户正在用的、
                    # 鼎信诺拉起后显示出来的）不延迟。读不到 Visible 时按可见处理，不延迟。
                    $visible = $true
                    try { $visible = [bool]$app.Visible } catch { $visible = $true }
                    if (-not $visible) {
                        $first = $hiddenSince[[string]$winPid]
                        if (-not $first) {
                            $hiddenSince[[string]$winPid] = Get-Date
                            WdLog ("隐藏实例，宽限一个周期再处理: pid=" + $winPid)
                            continue
                        }
                        if (((Get-Date) - $first).TotalSeconds -lt $IntervalSeconds) { continue }
                    }

                    # 父进程白名单：仅附着白名单进程拉起的 Excel（风险收敛，见仓库 Issue #1 / T13）
                    if ($filterSet.Count -gt 0) {
                        $excelPid2 = [int]$winPid
                        $parentObj = Get-CimInstance Win32_Process -Filter "ProcessId=$excelPid2" -ErrorAction SilentlyContinue
                        $pname = ""
                        if ($parentObj -and $parentObj.ParentProcessId) {
                            $pname = (Get-CimInstance Win32_Process -Filter "ProcessId=$($parentObj.ParentProcessId)" -ErrorAction SilentlyContinue).Name
                        }
                        if (-not $pname) {
                            # 父进程查询失败（CIM/WMI 临时故障）：本周期跳过且【不标 seen】，待下轮重查
                            WdLog ("hwnd=" + $h + " 父进程未知，本周期跳过（不缓存，下轮重查）")
                            continue
                        }
                        if ($filterSet -notcontains $pname.ToLowerInvariant()) {
                            $seen[$key] = $true
                            WdLog ("跳过 hwnd=" + $h + "：父进程 " + $pname + " 不在白名单")
                            continue
                        }
                    }

                    # ---- 已随宿主正常启动装载（右键菜单在位）→ 不做任何装载动作 ----
                    # v1.4.3 起安装器默认部署看门狗，普通双击打开的 Excel 启动时已由 OPENn 装好加载项：
                    # 再 Open + Run 一次属多余，还会在用户正在工作的会话里动右键菜单与快捷键。
                    # 鼎信诺拉起的自动化实例没有菜单，照常走下方装载流程。
                    # 注意：这里不要再加任何链式 COM 访问（如读取在位副本路径）。看门狗先运行一段时间再附着时，
                    # 这类调用产生的中间对象一次 GC 回收不掉，会把该 Excel 钉在内存里不退（WORKLOG #27 会诊实测）。
                    if (Menu-Exists $app) {
                        $seen[$key] = $true
                        WdLog ("已在位，无需装载: hwnd=" + $h + " pid=" + $winPid)
                        continue
                    }

                    # 写操作（打开加载项 / 执行宏）前复核目标窗口仍在：宿主若已开始退出，文档窗口首当其冲被销毁，
                    # 此时再 Open / Run 会打断退出流程。窗口已不在就什么都不做（下个周期它也不会再被枚举到）。
                    if (-not [AttachLinkWatchdog.Native]::IsWindowAlive($h)) { continue }

                    if (Provision-App $app) {
                        if ($script:provState -eq 'trusted' -and -not [AttachLinkWatchdog.Native]::IsWindowAlive($h)) {
                            # 装载期间目标窗口已消失（宿主开始退出）：不执行宏、不标记 seen（该窗口下个周期不会再被枚举到）
                            WdLog ("装载后目标窗口已消失（宿主正在退出），不执行初始化: hwnd=" + $h + " pid=" + $winPid)
                        } elseif ($script:provState -eq 'trusted') {
                            # 只有确认在位的那份就是 $XlamPath 本身，才执行初始化（显式正向判断，不靠排除法）。
                            # xlam 装载发生在工作簿打开之后，初始化时序已错过，补一次（幂等）。
                            try { $null = $app.Run('AttachLink_Initialize') }
                            catch { WdLog ("Run AttachLink_Initialize 失败: " + $_.Exception.Message) }
                            if (Menu-Exists $app) {
                                $seen[$key] = $true
                                $ver = ""
                                try { $ver = $app.Run('AttachLink_GetVersion') } catch {}
                                WdLog ("已装载并初始化: hwnd=" + $h + " pid=" + $winPid + $(if ($ver) { " 版本=v" + $ver } else { "" }))
                            } else {
                                WdLog ("装载后未检测到菜单, hwnd=" + $h + " pid=" + $winPid + "（下轮重试）")
                            }
                        } elseif ($script:provState -eq 'unknown') {
                            # 无法确认在位情况（宿主繁忙 / 路径不可读）：不打开、不执行宏，也不标记 seen，下个周期重试
                            WdLog ("无法确认在位的 AttachLink.xlam（查询出错或路径不可读），本周期不处理: hwnd=" + $h + " pid=" + $winPid)
                        } else {
                            # foreign：宿主里是一份异路径的同名副本；mismatch：打开后复核不符。都不打开第二份、不执行宏，
                            # 对该窗口不再重试（见 Provision-App）。
                            $seen[$key] = $true
                            if ($script:provState -eq 'mismatch') {
                                WdLog ("打开后复核不符（在位的不是看门狗配置的那份），不执行宏: hwnd=" + $h + " pid=" + $winPid + " 在位=" + $script:foreignPath)
                            } else {
                                WdLog ("同名异路径副本已装载且未初始化，不执行其宏、不打开第二份，跳过: hwnd=" + $h + " pid=" + $winPid + " 副本=" + $script:foreignPath)
                            }
                        }
                    }
                } catch {
                    WdLog ("附着失败 hwnd=" + $h + " : " + $_.Exception.Message + "（退避重试）")
                    Start-Sleep -Milliseconds 800
                } finally {
                    # 用完立即释放对宿主的对象引用（$app 在这里；其余对象在各函数的 finally 里已逐个 Release-Com）。
                    # 宿主在 Quit 之后要等外部 COM 引用全部释放才真正退出；旧版把引用一直持有到下次附着或 GC，
                    # 被附着过的 Excel 关闭后进程被钉在内存里不退（2026-11-23 实测残留 2~14 分钟；同事机 WPS
                    # 残留进程不断累积同源）。末尾的 GC 只作兜底，正确性不再依赖它。
                    Release-Com $app
                    $app = $null
                    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
                }
            }
            # 清理已消失进程 / 窗口的宽限记录（隐藏实例宽限表按 pid、隐藏窗口状态表按窗口键，宿主或窗口没了就删）
            $livePids = @{}; $liveKeys = @{}
            foreach ($h in $wins) {
                $p = [string][AttachLinkWatchdog.Native]::GetWindowProcessId($h)
                $livePids[$p] = $true
                $liveKeys[([string]$h + ':' + $p)] = $true
            }
            foreach ($k in @($hiddenSince.Keys)) { if (-not $livePids.ContainsKey($k)) { $hiddenSince.Remove($k) } }
            foreach ($k in @($hiddenWin.Keys)) { if (-not $liveKeys.ContainsKey($k)) { $hiddenWin.Remove($k) } }
        }
    } catch {
        WdLog ("cycle " + $cycle + " 异常: " + $_.Exception.Message)
    }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }
    Start-Sleep -Seconds $IntervalSeconds
}
