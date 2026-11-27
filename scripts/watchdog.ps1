# =====================================================================
# AttachLink 看门狗（Watchdog）v1.0
# ---------------------------------------------------------------------
# 背景：鼎信诺等审计软件通过 COM 自动化（CreateObject）拉起 Excel 时，
# Excel 跳过一切加载项初始化（.xlam / XLSTART / COM 加载项的 OnConnection
# 均不触发，本机逐项实测证实）。因此装载必须来自 Excel 进程外部。
#
# 机制：周期性枚举 Excel 的 EXCEL7 文档窗口（XLMAIN 的子窗口），通过
# oleacc!AccessibleObjectFromWindow(OBJID_NATIVEOM, IID_IDispatch) 取得
# 该窗口的 Excel Window 对象 → .Application → 按名核对在位副本 →
# Workbooks.Open(xlam) 强制装载（只装载、不登记，见下）→ 显式 Run 初始化
# → 验证右键菜单 Tag 在位。该通道是微软文档化的
# Office Native OM 附着方式，不依赖 ROT、不注入。
#
# 只装载、不登记（v1.6.0）：装载动作只有 Workbooks.Open，不触碰 AddIns 集合。
# 原因：① 自动化实例里 Installed=True 只是登记位、从不产生装载效果（本机实测，
# 登记了也「从未装载」）；② 登记会在宿主退出时把 OPENn / Add-in Manager 写脏
# （WORKLOG #30），是纯粹的状态污染。切换后会话内装载功能不变、加载项列表零差异
# （技术预研 docs\AttachLink技术预研-看门狗只装载不登记可行性.md）。用户级安装
# （install.ps1）的 OPENn 登记是交互式 Excel 启动装载的设计通道，不在此列、不受影响。
# 装载不留副作用（v1.6.0）：打开带 Workbook_Open 的加载项，Excel 在该事件处理结束时会把宿主的
# DisplayAlerts / ScreenUpdating 复位为 True；看门狗装载前记下原值、装载后恢复（Get-HostUiState /
# Restore-HostUiState），不把宿主（别的程序）关掉的警告重新打开。v1.6.4 起 AttachLink 自己不再带
# Workbook_Open、也不再处理「打开工作簿」事件，这两处复位不再发生；写回保留作兜底（宿主里若有别的带
# Workbook_Open 的加载项，或被装载的是旧版加载项）。
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
# 已在位即不动（v1.6.0 加固）：右键菜单 Tag 在位不可反推装载态——Excel 会把非临时的
# 内置命令栏控件在正常退出时写入 %APPDATA%\Microsoft\Excel\Excel15.xlb，此后每个新实例
# 都带着它启动；残留的 AttachLinkCtx 控件形成「死菜单」（菜单在位、加载项工作簿 notfound，
# 旧版据此判「已在位」而漏装载，2026-11-24 受控复现）。因此菜单在位后再用 Workbooks 三态
# 确认加载项工作簿：在位的就是 $XlamPath 才算「已在位」；在位的是异路径同名副本则维持不动；
# 明确 notfound（死菜单）按未装载处理：装载前先由看门狗【从外部】删除残留控件（根除——实测非临时控件在
# Workbook_Open 里被加载项删除后，实例正常退出时仍会被写回缓存，死菜单会在每个新实例重现；外部删除 +
# 正常退出才真正清缓存），再打开配置路径副本并初始化；
# 查询出错不动作、不标 seen、下周期重试（日志限频，避免宿主繁忙时刷日志）。判定集中在纯函数
# Get-PresenceDecision 里，由 scripts\test-watchdog-units.ps1 的假对象用例逐条回归。
# 两级都装的机器上，普通双击 Excel 的会话由用户级登记装载的是库目录副本（%APPDATA%\Microsoft\AddIns），
# 而看门狗配置的是 Program Files 副本——这类会话的判定是 foreign（在位的是异路径同名副本：同样标记完成、
# 不 Open、不 Run），每个交互式会话会记一行「已在位，无需装载 … 在位副本=…」，属正常；只有看门狗自己
# 装载过的会话、或只装了机器级的机器上，普通会话的判定才是 inplace。
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
# 只附着用户看得见的会话（v1.6.2）：Application.Visible 为假的实例不再附着，每个周期只读一次它的可见性，
# 变可见之后再按常规处理。AttachLink 是给人右键用的，用户看不见的 Excel 用不上右键菜单，往里装只有风险：
# 鼎信诺归档在一个独立的隐藏 Excel 进程里做，旧版会对它发起附着（2026-11-27 实测，WORKLOG #35）。
# 回归脚本用 -AttachHiddenInstances 回到旧规则（隐藏实例宽限一个周期后照常附着）；产品部署不带这个开关。
#
# 附着退避（v1.6.1）：某个文档窗口附着失败之后——宿主正忙（被呼叫方拒绝接收呼叫），或装载后没能确认初始化——
# 按窗口退避：2、4、8、16 秒，之后封顶 30 秒；退避期内不向这个窗口发任何调用，别的窗口照常处理，成功一次即清零。
# 旧版是固定等 800 毫秒后每个周期重试：鼎信诺归档期间 8 秒内对同一个进程被连续拒绝 3 次（2026-11-27 实测）。
# 成功判定从严（v1.6.1）：初始化宏执行成功、菜单 Tag 在位、版本宏返回非空，三者齐备才标「已装载并初始化」
# 并不再处理该窗口；缺任何一项都记一行并按上面的退避重试。宏名限定到工作簿（'AttachLink.xlam'!宏名）。
#
# 部署：install-elevated.ps1（管理员）将本脚本复制到
# C:\Program Files\AttachLink\ 并注册为登录计划任务（最高权限运行，
# 可同时服务普通 Excel 与管理员权限的鼎信诺 Excel）。v1.6.3 起计划任务的直接动作是同目录的
# AttachLink.WatchdogLauncher.exe（无窗口启动器，源码 src\WatchdogLauncher.cs）：它启动本脚本并盯着它运行，
# 本脚本异常退出 10 秒后拉起，连续 3 次没能稳定运行才放弃；启动器被结束时本脚本随它的作业对象一起结束。
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
    [string]$LogFile = '',
    # 仅供回归使用：隐藏实例（Application.Visible 为假）照旧附着（宽限一个周期）。产品部署不带这个开关——
    # v1.6.2 起看门狗只附着用户看得见的会话：隐藏的自动化工作进程（如鼎信诺归档时另起的 Excel）不再被装载。
    [switch]$AttachHiddenInstances
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

function Get-PresenceDecision([bool]$menuExists, [string]$lookupState, [bool]$samePath) {
    # 在位判定的纯函数（不碰 COM），主循环在菜单 Tag 在位时结合 Get-LoadedAddInPath 的三态调用它分流。
    # 参数：$menuExists = 右键菜单 Tag 是否在位；$lookupState = Workbooks 三态查询结果
    # （found / notfound / error，见 Get-LoadedAddInPath）；$samePath = 在位副本路径是否就是 $XlamPath。
    # 返回五种结论之一：
    #   notloaded = 菜单不在位 → 直接走装载流程（无论三态如何，此时尚未做三态查询）
    #   inplace   = 菜单在位且在位的就是 $XlamPath → 标记完成，不再 Open / Run
    #   foreign   = 菜单在位但在位的是异路径同名副本 → 维持现状不动它（与 Provision-App 同一纪律）
    #   deadmenu  = 菜单在位但加载项工作簿明确 notfound（命令栏缓存残留的「死菜单」）→ 按未装载走装载流程
    #               （主循环先调用 Remove-DeadMenuControls 从外部删除残留控件，再装载；原因见该函数的说明）
    #   unknown   = 菜单在位但三态查询出错 / 路径读不到 → 不标 seen，下周期重试
    if (-not $menuExists) { return 'notloaded' }
    if ($lookupState -eq 'found') { if ($samePath) { return 'inplace' } else { return 'foreign' } }
    if ($lookupState -eq 'notfound') { return 'deadmenu' }
    return 'unknown'
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

function Get-BackoffDelay([double]$prevSeconds, [double]$baseSeconds, [double]$maxSeconds) {
    # 附着失败后的下次重试间隔（秒）：首次取 base，之后逐次翻倍，封顶 max。纯函数，判定与调用方解耦，
    # 由 scripts\test-watchdog-units.ps1 逐条回归。prevSeconds <= 0 视为首次；base 大于 max 时以 max 为准。
    $d = if ($prevSeconds -gt 0) { $prevSeconds * 2 } else { $baseSeconds }
    if ($d -lt $baseSeconds) { $d = $baseSeconds }
    if ($d -gt $maxSeconds) { $d = $maxSeconds }
    return $d
}

function Get-InstanceDecision([int]$appVisible, [bool]$allowHidden, $firstSeen, $now, [int]$intervalSeconds) {
    # 实例可见性判定（v1.6.2：只附着用户看得见的会话）。纯函数，判定与调用方解耦，由 scripts\test-watchdog-units.ps1 回归。
    #   appVisible：1 = Application.Visible 为真；0 = 为假；-1 = 读不到
    #   返回 'attach'（按工作窗口处理）或 'defer'（本周期不处理、不标 seen，下周期再看）
    # allowHidden（仅回归用）= 旧规则：不可见实例首次出现只记时间，宽限一个周期后照常处理；读不到按可见处理。
    # 产品规则：看不见就不附着。AttachLink 是给人右键用的，用户看不见的 Excel 用不上右键菜单，往里装只有风险——
    # 鼎信诺归档在一个独立的隐藏 Excel 进程里做，旧版会对它发起附着（2026-11-27 实测，WORKLOG #35）。
    if ($appVisible -eq 1) { return 'attach' }
    if ($allowHidden) {
        if ($appVisible -lt 0) { return 'attach' }
        if ($null -eq $firstSeen) { return 'defer' }
        if (($now - [datetime]$firstSeen).TotalSeconds -lt $intervalSeconds) { return 'defer' }
        return 'attach'
    }
    return 'defer'
}

function Get-HostUiState($app) {
    # 宿主两项界面设置的原值快照（看门狗装载前调用）。
    # Excel 在 Workbook_Open 事件处理结束时会把 DisplayAlerts / ScreenUpdating 一律复位为 True（2026-11-26 实测：
    # 处理过程里读到的还是原值、在处理里再设一次也无济于事；任何带 Workbook_Open 的加载项都会触发，与
    # AttachLink 的代码无关）。看门狗把加载项装进别的程序拉起的 Excel 之后，宿主关掉的警告会被重新打开——
    # 随后的覆盖保存、删表等操作一旦弹框，隐藏的自动化实例就卡死（实测一次卡了 22 分钟，直到有人关掉那个
    # 「是否替换」）。所以装载前记下原值、装载后恢复（Restore-HostUiState）。读不到的项保持 $null，恢复时跳过。
    $alerts = $null; $screen = $null
    try { $alerts = $app.DisplayAlerts } catch {}
    try { $screen = $app.ScreenUpdating } catch {}
    return @{ alerts = $alerts; screen = $screen }
}

function Restore-HostUiState($app, $state) {
    # 把宿主的 DisplayAlerts / ScreenUpdating 恢复成装载前的值（Get-HostUiState 的快照）；快照为 $null 或某项
    # 没读到时不恢复该项。直接写回、不先比较：读当前值若静默失败（PS 5.1 下属性读取失败可能得到空值而
    # 不上抛）会误判成「已一致」。
    if ($null -eq $state) { return }
    if ($null -ne $state.alerts) { try { $app.DisplayAlerts = [bool]$state.alerts } catch {} }
    if ($null -ne $state.screen) { try { $app.ScreenUpdating = [bool]$state.screen } catch {} }
}

function Provision-App($app) {
    # 只装载、不登记（v1.6.0，R1）。AddIns 集合的 Installed=True 只是登记位，
    # 自动化实例可能「已登记勾选但从未装载」（本机实测根因），登记对装载没有任何效果；
    # 而登记会在宿主退出时把 OPENn / Add-in Manager 写脏（WORKLOG #30）。
    # 因此本函数不触碰 AddIns 集合，只做强制装载：xlam 为 IsAddin 工作簿，
    # Workbooks.Open 无可见窗口、不切换用户活动窗口；已装载时 Open 返回现有工作簿，无副作用。
    #
    # 返回 $true = 已得出确定结论，结论见 $script:provState；$false = 打开过程出错（已记日志）。
    #   trusted  = 宿主里在位的那份就是 $XlamPath（本周期由看门狗打开并复核，或此前已装载）——只有这个状态允许执行初始化宏
    #   foreign  = 宿主里在位的是一份【异路径】的同名副本 → 不打开第二份（两份同名副本相撞会让宿主退不干净），
    #              也绝不执行它的宏：那份可能来自非受信任位置或是旧版本，而看门狗服务的是管理员权限的 Excel
    #              （提权进程不得执行用户可写目录里的代码）
    #   mismatch = 看门狗执行了 Open，但打开后读到的在位副本路径与 $XlamPath 不符（期间宿主装入了同名副本 /
    #              Open 返回了别的工作簿）→ 不执行宏
    #   unknown  = 无法确认在位情况（查询出错或路径读不到）→ 不打开、不执行宏；调用方不标记 seen，下个周期重试
    #   failed   = 打开时出错
    $wbs = $null; $wb = $null; $uiState = $null
    $script:provState = 'failed'; $script:foreignPath = ''
    $done = $false
    try {
        $wbs = $app.Workbooks
        $loaded = Get-LoadedAddInPath $wbs 'AttachLink.xlam'
        if ($script:lookupState -eq 'error') {
            $script:provState = 'unknown'
        } elseif ($script:lookupState -eq 'found') {
            # 已有同名加载项在位：不打开第二份，只判定它是不是 $XlamPath 本身
            if (Test-SamePath $loaded $XlamPath) { $script:provState = 'trusted' }
            else { $script:provState = 'foreign'; $script:foreignPath = $loaded }
        } else {
            # 明确不在位：从 $XlamPath 打开（只装载、不登记）
            $uiState = Get-HostUiState $app
            $wb = $wbs.Open($XlamPath)
            # 打开后立即复核：在位的那份必须就是 $XlamPath（检查与打开之间宿主可能已装入同名副本）
            $opened = ''
            try { $opened = [string]$wb.FullName } catch { $opened = '' }
            if (Test-SamePath $opened $XlamPath) { $script:provState = 'trusted' }
            else { $script:provState = 'mismatch'; $script:foreignPath = $opened }
        }
        $done = $true
    } catch {
        WdLog ("装载失败（Workbooks.Open）: " + $_.Exception.Message)
    } finally {
        # 宿主显示设置的恢复放在 finally：打开带 Workbook_Open 的加载项时，Excel 在该事件处理结束时已把
        # DisplayAlerts / ScreenUpdating 复位为 True（v1.6.4 起 AttachLink 自身不再触发，写回留作兜底）——
        # 复核不符要恢复，Workbooks.Open 中途抛错（事件可能已经跑完）同样要恢复；没走到打开这一步
        # （$uiState 为 $null）则不碰。
        if ($null -ne $uiState) { Restore-HostUiState $app $uiState }
        Release-Com $wb; Release-Com $wbs
    }
    return $done
}

function Menu-Exists($app) {
    # 右键菜单 Tag 是否在位。下标遍历 + 逐个释放（每个控件查完即在 finally 里释放，属性读取抛异常
    # ——宿主繁忙 / RPC 拒绝——时同样不漏）；任何异常原样抛给主循环。
    # 注意：这只是「已初始化」的幂等标记，不是加载项健康检查。
    # 读不到 ≠ 不在位（v1.6.0）：控件数读不到时直接抛异常；控件的 Tag 读不到时先记一笔、继续扫描——找到带标签
    # 的控件就说明菜单在位（途中别的控件读不到不影响这个结论），没找到且有控件读不到才是「无法确认菜单不在位」，
    # 抛异常交给主循环走「附着失败（退避重试）」。PS 5.1 下属性读取失败可能不上抛而静默得到空值（名为 Count
    # 的属性实测如此），真实 Excel 里无标签控件的 Tag 是空串而不是 $null——所以 $null 只可能是读不到。若当成
    # 「0 个控件 / 无标签」返回 $false，已在运行的会话会被当成未装载再装载、初始化一次；而遇到第一个读不到的
    # 控件就抛，则第三方加载项插在前面的某个控件若一直读不到，菜单明明在位也会永远装载不上。
    $found = $false
    $unknown = 0
    $bars = $null; $bar = $null; $ctrls = $null; $ctrl = $null
    try {
        $bars = $app.CommandBars
        $bar = $bars.Item('Cell')
        $ctrls = $bar.Controls
        $nRaw = $ctrls.Count
        if ($null -eq $nRaw) { throw '命令栏控件数读不到（宿主繁忙）' }
        $n = [int]$nRaw
        for ($i = 1; $i -le $n; $i++) {
            $ctrl = $null
            try {
                $ctrl = $ctrls.Item($i)
                $tag = $ctrl.Tag
                if ($null -eq $tag) { $unknown++ }
                elseif ($tag -eq 'AttachLinkCtx') { $found = $true }
            } finally { Release-Com $ctrl }
            if ($found) { break }
        }
        if ($unknown -gt 0 -and -not $found) { throw ('有 ' + $unknown + ' 个控件的 Tag 读不到（宿主繁忙），且没有找到带标签的菜单项') }
    } finally {
        Release-Com $ctrl; Release-Com $ctrls; Release-Com $bar; Release-Com $bars
    }
    return $found
}

function Remove-DeadMenuControls($app) {
    # 死菜单根除（v1.6.0 R2）：从【外部】删除 Cell 命令栏上带 AttachLinkCtx 标签的控件，返回删掉的个数；
    # -1 = 一个都没删成、且存在读不到 / 删不掉的控件——命令栏整体不可访问，或宿主繁忙导致个别控件读不到 Tag
    #       （读不到的那个可能就是死菜单，不能谎报「未找到」），调用方按「暂未删除、稍后重试」处理。
    # 为什么不能只靠装载后的初始化去删（2026-11-26 实测，WORKLOG #34）：非临时的 AttachLinkCtx 控件如果是在
    # Workbooks.Open 触发的 Workbook_Open 里被加载项的 RegisterMenu 删除，只是当前会话里不见了，实例正常
    # 退出时仍会被写回 Excel15.xlb——缓存里的死菜单永远清不掉，每个新实例都会再带着它启动。由外部客户端
    # 删除并正常退出才会真正清缓存。只在 deadmenu 结论下调用（菜单 Tag 在位、加载项工作簿明确 notfound，
    # 即这些控件点了也只会报宏不可用），删除后照常走装载流程，由初始化重建正式菜单。
    # COM 纪律同 Menu-Exists：下标遍历（倒序——删除不改变尚未遍历部分的下标）、每个控件用完即在 finally 里
    # 释放、不留链式访问；单个控件 Tag 读不到 / 删不掉只跳过它，不影响其余控件。
    $removed = 0; $failed = 0
    $bars = $null; $bar = $null; $ctrls = $null; $ctrl = $null
    try {
        $bars = $app.CommandBars
        $bar = $bars.Item('Cell')
        $ctrls = $bar.Controls
        # 控件数读不到时显式按「不可访问」处理：PS 5.1 对名为 Count 的属性有特殊处理，读取失败不一定上抛、
        # 可能静默得到空值（实测）——不能把它当成 0 个控件。
        $nRaw = $ctrls.Count
        if ($null -eq $nRaw) {
            $removed = -1
        } else {
            $n = [int]$nRaw
            for ($i = $n; $i -ge 1; $i--) {
                $ctrl = $null
                try {
                    $ctrl = $ctrls.Item($i)
                    # Tag 读到 $null 按「读不到」计一次失败：真实 Excel 里无标签控件的 Tag 是空串而不是 $null
                    # （实测 Cell 栏 51 个内置控件全部为 ""），而 PS 5.1 下属性读取失败不一定上抛、可能静默得到
                    # $null——读不到的那个可能就是死菜单，不能当成「无标签」放过
                    $tag = $ctrl.Tag
                    if ($null -eq $tag) { $failed++ }
                    elseif ($tag -eq 'AttachLinkCtx') { [void]$ctrl.Delete(); $removed++ }
                } catch {
                    # 取不到该控件，或读 Tag / 删除时抛异常（宿主繁忙）：跳过它，继续处理其余控件
                    $failed++
                } finally { Release-Com $ctrl }
            }
        }
    } catch {
        $removed = -1
    } finally {
        Release-Com $ctrl; Release-Com $ctrls; Release-Com $bar; Release-Com $bars
    }
    if ($removed -eq 0 -and $failed -gt 0) { return -1 }
    return $removed
}

WdLog ("看门狗启动, pid=" + $PID + ", xlam=" + $XlamPath)      # 带上自身 pid：常驻与被测看门狗可能共用同一份日志，排障时要能区分
$seen = @{}
$cycle = 0
$lastCount = -1               # 上次记入日志的 EXCEL7 窗口数（变化才记周期行）
$lastBeat = Get-Date          # 上次记周期行/心跳的时刻
$xlamMissing = $false         # xlam 缺失状态（仅状态切换时记日志，另配心跳）
$hiddenSince = @{}            # 隐藏实例的首次见到时刻（pid → 时间）：产品规则下只用于「首次记一行、变可见时清掉」；-AttachHiddenInstances 下作宽限计时
$hiddenWin = @{}              # 文档窗口在「隐藏判定」上的状态（窗口键 → New-WindowState），逻辑见 Get-HiddenDecision / Test-HiddenReevaluate
$maxReeval = 3                # 已判定为隐藏的窗口因「Win32 由不可见变为可见」而立即重新评估的次数上限
$hiddenRecheckSeconds = 30    # Win32 可见却没有新跳变时的低频只读重查间隔（秒）；不设次数上限，避免窗口被永久跳过
$lastPresenceUnknownLog = $null   # 上次记「菜单在位但无法确认装载状态」日志的时刻（限频 60 秒一行，全局一条）
$backoff = @{}                  # 附着退避表（v1.6.1）：窗口键 → @{ until = 下次可试时刻; delay = 本次间隔秒数 }，间隔见 Get-BackoffDelay
$backoffMaxSeconds = 30         # 退避间隔上限（秒）
$initMacro = "'AttachLink.xlam'!AttachLink_Initialize"       # 宏名限定到工作簿（v1.6.1）
$versionMacro = "'AttachLink.xlam'!AttachLink_GetVersion"

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

                # ---- 附着退避（v1.6.1）：这个窗口上一次附着失败后，退避期内不向宿主发任何调用 ----
                $bo = $backoff[$key]
                if ($null -ne $bo -and (Get-Date) -lt $bo.until) { continue }

                $app = $null
                $attachFailed = $false
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

                    # ---- 实例可见性（v1.6.2：只附着用户看得见的会话；判定在 Get-InstanceDecision 里）----
                    # Application.Visible 为假的实例是隐藏的自动化实例：各类工具、测试脚本，以及鼎信诺归档时另起的那个 Excel。
                    # 它们用不上右键菜单，往里装只有风险，所以不再附着，等它变可见再说（每个周期只读一次 Visible）。
                    # 回归脚本用 -AttachHiddenInstances 回到旧规则（首次出现只记时间，宽限一个周期后照常处理）。
                    $appVisible = -1
                    try { $appVisible = [int][bool]$app.Visible } catch { $appVisible = -1 }
                    $hvPidKey = [string]$winPid
                    $decisionV = Get-InstanceDecision $appVisible $AttachHiddenInstances $hiddenSince[$hvPidKey] (Get-Date) $IntervalSeconds
                    if ($decisionV -eq 'defer') {
                        if ($null -eq $hiddenSince[$hvPidKey]) {
                            $hiddenSince[$hvPidKey] = Get-Date
                            if ($AttachHiddenInstances) { WdLog ("隐藏实例，宽限一个周期再处理: pid=" + $winPid) }
                            elseif ($appVisible -eq 0) { WdLog ("隐藏实例，暂不附着（等它可见）: pid=" + $winPid) }
                            else { WdLog ("实例可见性暂时无法确认，本周期不处理（下周期重试）: pid=" + $winPid) }
                        }
                        continue
                    }
                    if ($null -ne $hiddenSince[$hvPidKey]) {
                        # 先前按隐藏记下的实例现在可见了：清掉记录；产品规则下记一行，说明从这一刻起开始处理
                        $hiddenSince.Remove($hvPidKey)
                        if (-not $AttachHiddenInstances) { WdLog ("先前隐藏的实例已可见，开始处理: pid=" + $winPid) }
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

                    # ---- 在位判定（v1.6.0 加固，R2）：菜单 Tag + Workbooks 三态联合确认 ----
                    # 菜单在位不可反推装载态：Excel 会把非临时的命令栏控件写入 Excel15.xlb 缓存，
                    # 残留的 AttachLinkCtx 控件形成「死菜单」（菜单在位、加载项工作簿 notfound，旧版
                    # 据此漏装载，2026-11-24 受控复现）。因此菜单在位后再用 Workbooks 三态确认加载项
                    # 工作簿在位情况，由纯函数 Get-PresenceDecision 分流：
                    #   inplace   = 在位的就是 $XlamPath → 标记完成，不再 Open / Run（普通双击 Excel 的常态）
                    #   foreign   = 在位的是异路径同名副本 → 维持不动（与 Provision-App 同一纪律）
                    #   deadmenu  = 加载项工作簿明确 notfound → 按未装载走下方装载流程（装载前先由看门狗从外部
                    #               删除残留控件——根除缓存，见 Remove-DeadMenuControls）
                    #   unknown   = 查询出错 / 路径读不到 → 不动作、不标 seen，下周期重试（日志限频 60 秒一行）
                    # 菜单不在位（notloaded）不做三态查询，直接走装载流程。
                    # COM 纪律（WORKLOG #27）：取到的 Workbooks 存局部变量、finally Release-Com，
                    # 不新增链式访问；Get-LoadedAddInPath 内部读完即释放。
                    $deadMenu = $false
                    if (Menu-Exists $app) {
                        $wbsL = $null
                        $lookupState = 'error'; $lookupPath = $null
                        try {
                            $wbsL = $app.Workbooks
                            $lookupPath = Get-LoadedAddInPath $wbsL 'AttachLink.xlam'
                            $lookupState = $script:lookupState
                        } catch {
                            $lookupState = 'error'    # 连 Workbooks 都取不到（宿主正忙）→ 按无法确认处理
                        } finally {
                            Release-Com $wbsL
                        }
                        $samePath = $false
                        if ($lookupState -eq 'found') { $samePath = Test-SamePath $lookupPath $XlamPath }
                        $presence = Get-PresenceDecision $true $lookupState $samePath
                        if ($presence -eq 'inplace') {
                            $seen[$key] = $true
                            WdLog ("已在位，无需装载: hwnd=" + $h + " pid=" + $winPid)
                            continue
                        } elseif ($presence -eq 'foreign') {
                            $seen[$key] = $true
                            WdLog ("已在位，无需装载: hwnd=" + $h + " pid=" + $winPid + " 在位副本=" + $lookupPath)
                            continue
                        } elseif ($presence -eq 'unknown') {
                            # 宿主繁忙（如用户正在编辑单元格）会每个周期都走到这里：限频 60 秒一行，
                            # 避免每 2 秒刷一行把事件行淹没。不标 seen，下周期重试。
                            $nowU = Get-Date
                            if ($null -eq $script:lastPresenceUnknownLog -or (($nowU - [datetime]$script:lastPresenceUnknownLog).TotalSeconds -ge 60)) {
                                $script:lastPresenceUnknownLog = $nowU
                                WdLog ("菜单在位但无法确认加载项装载状态（查询出错或路径不可读），本周期不处理: hwnd=" + $h + " pid=" + $winPid)
                            }
                            continue
                        }
                        # deadmenu：菜单在位但加载项工作簿明确不在位（命令栏缓存残留），按未装载处理
                        WdLog ("菜单在位但加载项未装载（命令栏缓存残留），按未装载处理: hwnd=" + $h + " pid=" + $winPid)
                        $deadMenu = $true
                    }

                    # 写操作（删除死菜单控件 / 打开加载项 / 执行宏）前复核目标窗口仍在：宿主若已开始退出，文档窗口首当其冲被销毁，
                    # 此时再动命令栏或 Open / Run 会打断退出流程。窗口已不在就什么都不做（下个周期它也不会再被枚举到）。
                    if (-not [AttachLinkWatchdog.Native]::IsWindowAlive($h)) { continue }

                    if ($deadMenu) {
                        # 死菜单根除：装载之前先由看门狗从外部删掉残留控件（为什么必须从外部删见 Remove-DeadMenuControls）。
                        # 删不掉不影响本周期装载——装载后的初始化仍会在会话内清掉死菜单，缓存留待下个实例再试。
                        $deadRemoved = Remove-DeadMenuControls $app
                        if ($deadRemoved -gt 0) {
                            WdLog ("已删除死菜单控件 " + $deadRemoved + " 个（实例正常退出后不再写入命令栏缓存）: hwnd=" + $h + " pid=" + $winPid)
                        } elseif ($deadRemoved -eq 0) {
                            WdLog ("未找到可删除的死菜单控件（期间已被其它方清理），继续装载: hwnd=" + $h + " pid=" + $winPid)
                        } else {
                            WdLog ("死菜单控件暂未删除（宿主正忙，命令栏不可访问），本周期仍照常装载，缓存留待下个实例再试: hwnd=" + $h + " pid=" + $winPid)
                        }
                    }

                    if (Provision-App $app) {
                        if ($script:provState -eq 'trusted' -and -not [AttachLinkWatchdog.Native]::IsWindowAlive($h)) {
                            # 装载期间目标窗口已消失（宿主开始退出）：不执行宏、不标记 seen（该窗口下个周期不会再被枚举到）
                            WdLog ("装载后目标窗口已消失（宿主正在退出），不执行初始化: hwnd=" + $h + " pid=" + $winPid)
                        } elseif ($script:provState -eq 'trusted') {
                            # 只有确认在位的那份就是 $XlamPath 本身，才执行初始化（显式正向判断，不靠排除法）。
                            # 显式初始化（幂等）：v1.6.4 起加载项没有 Workbook_Open，这是外部装载唯一的初始化入口
                            # （菜单、快捷键、应用事件都由它建立）；对旧版加载项则不能依赖其 Workbook_Open 一定执行过
                            # （事件关闭时它不执行，已装载时 Workbooks.Open 返回现有工作簿、也不再触发）。
                            # 成功判定从严（v1.6.1）：初始化宏执行成功、菜单 Tag 在位、版本宏返回非空，三者齐备才标 seen。
                            $initOk = $false
                            try { $null = $app.Run($initMacro); $initOk = $true }
                            catch { WdLog ("Run AttachLink_Initialize 失败: " + $_.Exception.Message) }
                            $menuOk = $false
                            try { $menuOk = [bool](Menu-Exists $app) } catch { WdLog ("装载后菜单读不到（退避重试）: " + $_.Exception.Message) }
                            $ver = ''
                            if ($initOk -and $menuOk) {
                                try { $ver = [string]$app.Run($versionMacro) } catch { $ver = '' }
                            }
                            if ($initOk -and $menuOk -and $ver) {
                                $seen[$key] = $true
                                WdLog ("已装载并初始化: hwnd=" + $h + " pid=" + $winPid + " 版本=v" + $ver)
                            } else {
                                # 没能确认就不标 seen，按退避重试。旧写法是初始化失败后只要菜单在位就算成功，版本只用于日志。
                                $attachFailed = $true
                                WdLog ("装载后未能确认初始化（初始化" + $(if ($initOk) { '成功' } else { '失败' }) + "、菜单" + $(if ($menuOk) { '在位' } else { '不在位' }) + "、版本" + $(if ($ver) { $ver } else { '读不到' }) + "），下轮重试: hwnd=" + $h + " pid=" + $winPid + "（退避重试）")
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
                    $attachFailed = $true
                    WdLog ("附着失败 hwnd=" + $h + " : " + $_.Exception.Message + "（退避重试）")
                } finally {
                    # 用完立即释放对宿主的对象引用（$app 在这里；其余对象在各函数的 finally 里已逐个 Release-Com）。
                    # 宿主在 Quit 之后要等外部 COM 引用全部释放才真正退出；旧版把引用一直持有到下次附着或 GC，
                    # 被附着过的 Excel 关闭后进程被钉在内存里不退（2026-11-23 实测残留 2~14 分钟；同事机 WPS
                    # 残留进程不断累积同源）。末尾的 GC 只作兜底，正确性不再依赖它。
                    Release-Com $app
                    $app = $null
                    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
                    # 退避记账（v1.6.1）：失败的窗口按 Get-BackoffDelay 排下次可试的时刻；得出任何非失败的结论即清零。
                    # 旧版这里是「固定等 800 毫秒再按周期重试」，会拖慢同一轮里别的窗口。
                    if ($attachFailed) {
                        $boPrevDelay = 0
                        if ($null -ne $backoff[$key]) { $boPrevDelay = $backoff[$key].delay }
                        $boNextDelay = Get-BackoffDelay $boPrevDelay $IntervalSeconds $backoffMaxSeconds
                        $backoff[$key] = @{ until = (Get-Date).AddSeconds($boNextDelay); delay = $boNextDelay }
                    } else {
                        $backoff.Remove($key)
                    }
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
            foreach ($k in @($backoff.Keys)) { if (-not $liveKeys.ContainsKey($k)) { $backoff.Remove($k) } }
        }
    } catch {
        WdLog ("cycle " + $cycle + " 异常: " + $_.Exception.Message)
    }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }
    Start-Sleep -Seconds $IntervalSeconds
}
