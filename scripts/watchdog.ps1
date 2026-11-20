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
    [string]$ParentAllowlist = ''
)

$ErrorActionPreference = 'Continue'
$logFile = Join-Path $env:TEMP 'AttachLinkWatchdog.log'

function WdLog([string]$m) {
    try { [IO.File]::AppendAllText($logFile, (Get-Date -Format 'HH:mm:ss') + ' ' + $m + "`r`n") } catch {}
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

        // 枚举 XLMAIN 顶层窗口下的 EXCEL7 文档窗口句柄
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

        // 从 EXCEL7 窗口取 Excel Window 对象，再返回其 Application
        // （用反射而非 dynamic：Add-Type 默认编译环境不含 Microsoft.CSharp）
        public static object GetApplicationFromWindow(long hwnd)
        {
            object win;
            var iid = IID_IDispatch;
            int hr = AccessibleObjectFromWindow(new IntPtr(hwnd), OBJID_NATIVEOM, ref iid, out win);
            if (hr != 0 || win == null) return null;
            return win.GetType().InvokeMember("Application",
                System.Reflection.BindingFlags.GetProperty, null, win, null);
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

function Provision-App($app) {
    # 登记装载（幂等）。注意：AddIns 集合的 Installed=True 只反映登记状态，
    # 自动化实例可能「已登记勾选但从未装载」（本机实测根因）。真正的强制
    # 装载靠下方 Workbooks.Open：xlam 为 IsAddin 工作簿，Open 无可见窗口、
    # 不切换用户活动窗口；已装载时 Open 返回现有工作簿，无副作用。
    $entry = $null
    foreach ($ai in $app.AddIns) {
        if ($ai.Name -ieq 'AttachLink.xlam') { $entry = $ai; break }
    }
    if ($null -eq $entry) {
        $null = $app.AddIns.Add($XlamPath, $false)
        foreach ($ai in $app.AddIns) {
            if ($ai.Name -ieq 'AttachLink.xlam') { $entry = $ai; break }
        }
    }
    if ($null -ne $entry -and -not $entry.Installed) { $entry.Installed = $true }
    try {
        $null = $app.Workbooks.Open($XlamPath)
    } catch {
        WdLog ("Workbooks.Open(xlam) 失败: " + $_.Exception.Message)
        return $false
    }
    return $true
}

function Menu-Exists($app) {
    foreach ($c in $app.CommandBars.Item('Cell').Controls) {
        if ($c.Tag -eq 'AttachLinkCtx') { return $true }
    }
    return $false
}

WdLog ("看门狗启动, xlam=" + $XlamPath)
$seen = @{}
$cycle = 0

while ($true) {
    $cycle++
    try {
        $wins = [AttachLinkWatchdog.Native]::FindExcel7Windows()
        WdLog ("cycle " + $cycle + ": EXCEL7 窗口数=" + $wins.Count)
        foreach ($h in $wins) {
            $key = [string]$h + ':' + [AttachLinkWatchdog.Native]::GetWindowProcessId($h)
            if ($seen.ContainsKey($key)) { continue }
            try {
                $app = [AttachLinkWatchdog.Native]::GetApplicationFromWindow($h)
                if ($null -eq $app) { WdLog ("hwnd=" + $h + " 未取得 Application"); continue }

                # 父进程白名单：仅附着白名单进程拉起的 Excel（风险收敛，见仓库 Issue #1 / T13）
                if ($filterSet.Count -gt 0) {
                    $excelPid2 = [int]($key -replace '^\d+:', '')
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

                if (Provision-App $app) {
                    # xlam 装载发生在工作簿打开之后，初始化时序已错过，显式补一次（幂等）
                    try { $null = $app.Run('AttachLink_Initialize') }
                    catch { WdLog ("Run AttachLink_Initialize 失败: " + $_.Exception.Message) }
                    if (Menu-Exists $app) {
                        $seen[$key] = $true
                        $ver = ""
                        try { $ver = $app.Run('AttachLink_GetVersion') } catch {}
                        WdLog ("已装载并初始化: hwnd=" + $h + $(if ($ver) { " 版本=v" + $ver } else { "" }))
                    } else {
                        WdLog ("装载后未检测到菜单, hwnd=" + $h + "（下轮重试）")
                    }
                }
            } catch {
                WdLog ("附着失败 hwnd=" + $h + " : " + $_.Exception.Message + "（退避重试）")
                Start-Sleep -Milliseconds 800
            }
        }
    } catch {
        WdLog ("cycle " + $cycle + " 异常: " + $_.Exception.Message)
    }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }
    Start-Sleep -Seconds $IntervalSeconds
}
