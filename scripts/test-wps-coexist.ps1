# =====================================================================
# AttachLink 看门狗 × WPS 共存回归（v1.4.3）
# ---------------------------------------------------------------------
# 背景：WPS 表格（et.exe / wps.exe）为兼容自动化复用 Excel 的 XLMAIN /
# EXCEL7 窗口类名。看门狗若只按类名附着，会向 WPS 注入 xlam，导致 WPS 卡在
# 「正在打开文档…」或反复弹「同名文件已打开」（2026-11-23 同事机日志，WORKLOG #27）。
#
# 验证的契约：看门狗对「不属于 EXCEL.EXE 的 XLMAIN/EXCEL7 窗口」零接触——
#   在看门狗自己的日志里，这类窗口只允许出现一行「跳过 …仅附着 EXCEL.EXE」；
#   不得出现「未取得 Application / 附着失败 / Workbooks.Open / Run / 已装载」
#   （这些是取 Application 之后每条路径的必经日志，出现即说明宿主校验失效）。
#   另验证 xlam 缺失时不处理任何窗口。
#
# 两段：
#   A. 仿冒宿主（任何机器可跑，不需要 WPS / Excel，约 20 秒）：本测试进程自建
#      隐藏的 XLMAIN 顶层窗口 + EXCEL7 子窗口，跑看门狗若干周期后核对独立日志。
#      旧版看门狗（无宿主校验）会对该窗口记「未取得 Application / 附着失败」→ FAIL。
#   B. 真实 WPS（装有 WPS 时执行；-SkipRealWps 强制跳过）：经 COM 拉起独立
#      WPS 实例开一个工作簿，核对看门狗跳过且实例未被触碰。**机器级常驻看门狗
#      仍是旧版时本段自动跳过**——旧版会向测试实例弹「同名文件已打开」
#      （2026-11-23 本机实踩），须先重新部署修复版。
#
# 隔离手段：看门狗一律写独立日志（-LogFile），并以 -ParentAllowlist 指向不存在的
# 父进程名——本机上真实 Excel 窗口因此至多被读取 Application，绝不会被注入。
#
# 用法：powershell -ExecutionPolicy Bypass -File test-wps-coexist.ps1 [-SkipRealWps] [-ForceCleanWps]
#   -SkipRealWps   跳过 B 段
#   -ForceCleanWps 明确授权：B 段测试实例 Quit 后若仍有测试期间新出现的 et/wps 进程，强制结束它们
#                  （默认不强杀——这些进程未必都属于测试实例，用户恰好另开的 WPS 窗口会被误伤；残留即判 FAIL）
# 期望输出末行：RESULT: PASS
# 编码：本文件含中文，必须保存为 UTF-8 with BOM（WORKLOG 根因 #11）。
# =====================================================================
param([switch]$SkipRealWps, [switch]$ForceCleanWps)

$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$wdPath = Join-Path $PSScriptRoot 'watchdog.ps1'
if (-not (Test-Path $wdPath)) { throw "未找到 $wdPath" }
$base = Join-Path $env:TEMP 'opencode\wdcoexist'
Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path (Join-Path $base 'missing') | Out-Null

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class AttachLinkCoexist
{
    private delegate bool EnumProc(IntPtr h, IntPtr lp);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc cb, IntPtr lp);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lp);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool DestroyWindow(IntPtr h);

    // 与看门狗 FindExcel7Windows 同构的枚举，返回 "EXCEL7句柄|EXCEL7进程ID|XLMAIN句柄|XLMAIN进程ID"
    public static List<string> FindXl()
    {
        var list = new List<string>();
        EnumWindows((top, lp) =>
        {
            if (Is(top, "XLMAIN"))
            {
                uint tpid; GetWindowThreadProcessId(top, out tpid);
                EnumChildWindows(top, (child, lp2) =>
                {
                    if (Is(child, "EXCEL7"))
                    {
                        uint cpid; GetWindowThreadProcessId(child, out cpid);
                        list.Add(child.ToInt64() + "|" + cpid + "|" + top.ToInt64() + "|" + tpid);
                    }
                    return true;
                }, IntPtr.Zero);
            }
            return true;
        }, IntPtr.Zero);
        return list;
    }
    private static bool Is(IntPtr h, string cls)
    {
        var sb = new StringBuilder(64); GetClassName(h, sb, 64); return sb.ToString() == cls;
    }

    // ---- 仿冒宿主：在本进程注册 XLMAIN / EXCEL7 类并建隐藏窗口 ----
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WNDCLASS
    {
        public uint style; public IntPtr lpfnWndProc; public int cbClsExtra; public int cbWndExtra;
        public IntPtr hInstance; public IntPtr hIcon; public IntPtr hCursor; public IntPtr hbrBackground;
        public string lpszMenuName; public string lpszClassName;
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern ushort RegisterClassW(ref WNDCLASS wc);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateWindowExW(uint ex, string cls, string name, uint style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr param);
    [DllImport("user32.dll")] private static extern IntPtr DefWindowProcW(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandleW(string name);

    public static IntPtr Top, Child;
    private static WndProcShim _proc;     // 持有委托：窗口过程 thunk 所指的委托被回收会导致本进程崩溃
    public static void CreateFake()
    {
        IntPtr inst = GetModuleHandleW(null);
        _proc = DefWindowProcW;
        WNDCLASS wc = new WNDCLASS();
        wc.lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_proc);
        wc.hInstance = inst;
        wc.lpszClassName = "XLMAIN"; RegisterClassW(ref wc);
        wc.lpszClassName = "EXCEL7"; RegisterClassW(ref wc);
        // 隐藏的顶层 XLMAIN + 其下的隐藏子窗口 EXCEL7（WS_CHILD = 0x40000000）
        Top = CreateWindowExW(0, "XLMAIN", "AttachLink coexist test fake host", 0, 0, 0, 0, 0, IntPtr.Zero, IntPtr.Zero, inst, IntPtr.Zero);
        Child = CreateWindowExW(0, "EXCEL7", "fake book", 0x40000000, 0, 0, 0, 0, Top, IntPtr.Zero, inst, IntPtr.Zero);
    }
    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate IntPtr WndProcShim(IntPtr h, uint m, IntPtr w, IntPtr l);
}
'@

$myProc = (Get-Process -Id $PID).ProcessName + '.exe'
$noneParent = 'attachlink-coexist-test-no-such-parent.exe'

function Invoke-Watchdog([string]$tag, [string]$xlam, [int]$cycles) {
    # 独立日志 + 空父进程白名单：不与机器级常驻看门狗共用日志，也不向本机真实 Excel 注入
    $log = Join-Path $base ("wd-$tag.log")
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $wdPath),
        '-MaxCycles', "$cycles", '-LogFile', ('"{0}"' -f $log), '-XlamPath', ('"{0}"' -f $xlam),
        '-ParentAllowlist', $noneParent)
    $p = Start-Process powershell -ArgumentList $argList -PassThru -WindowStyle Hidden
    $p | Wait-Process -Timeout (30 + 4 * $cycles) -ErrorAction SilentlyContinue
    if (-not $p.HasExited) { try { Stop-Process -Id $p.Id -Force } catch {} }
    if (-not (Test-Path $log)) { return @() }
    return @([IO.File]::ReadAllLines($log, [Text.Encoding]::UTF8))
}

# =====================================================================
'=== A. 仿冒宿主：非 EXCEL.EXE 进程的 XLMAIN/EXCEL7 窗口必须被跳过 ==='
[AttachLinkCoexist]::CreateFake()
if ([AttachLinkCoexist]::Top -eq [IntPtr]::Zero -or [AttachLinkCoexist]::Child -eq [IntPtr]::Zero) {
    Check $false '仿冒窗口创建失败（环境问题）'
} else {
    $child = [string][AttachLinkCoexist]::Child.ToInt64()
    $pidMine = [uint32]$PID
    $found = @([AttachLinkCoexist]::FindXl() | Where-Object { $_.StartsWith($child + '|') })
    Check ($found.Count -eq 1 -and $found[0].StartsWith($child + '|' + $pidMine + '|')) "仿冒 EXCEL7 可被看门狗的枚举算法看到（hwnd=$child pid=$pidMine）"

    $dummyXlam = Join-Path $base 'dummy.xlam'
    [IO.File]::WriteAllBytes($dummyXlam, [byte[]]@(1))

    '--- A1. xlam 在位：宿主校验生效 ---'
    $lines = Invoke-Watchdog 'A1' $dummyXlam 4
    Check (@($lines | Where-Object { $_ -match '看门狗启动' }).Count -eq 1) '看门狗已启动（独立日志）'
    $skipPattern = '跳过 hwnd=' + $child + '：宿主进程=' + [regex]::Escape($myProc) + '，仅附着 EXCEL\.EXE'
    $skipLines = @($lines | Where-Object { $_ -match $skipPattern })
    Check ($skipLines.Count -eq 1) "仿冒窗口被跳过且只记一次：宿主进程=$myProc（实际 $($skipLines.Count) 行）"
    $touchLines = @($lines | Where-Object { $_.Contains("hwnd=$child") -and $_ -notmatch $skipPattern })
    Check ($touchLines.Count -eq 0) "仿冒窗口零接触：无取对象/失败重试/注入日志（实际 $($touchLines.Count) 行）"
    $injLines = @($lines | Where-Object { $_ -match 'Workbooks\.Open|Run AttachLink|已装载并初始化|装载后未检测到菜单' })
    Check ($injLines.Count -eq 0) "空父进程白名单下未向本机任何真实 Excel 窗口注入（实际 $($injLines.Count) 行）"

    '--- A2. xlam 缺失：不枚举窗口、不向任何宿主发调用 ---'
    $lines2 = Invoke-Watchdog 'A2' (Join-Path $base 'missing\AttachLink.xlam') 3
    $missingLines = @($lines2 | Where-Object { $_ -match 'xlam 不存在，暂停附着直至文件就位' })
    Check ($missingLines.Count -eq 1) "xlam 缺失只记一次暂停日志（实际 $($missingLines.Count) 行）"
    $otherLines = @($lines2 | Where-Object { $_ -notmatch 'Add-Type 完成|看门狗启动|父进程白名单生效|xlam 不存在，暂停附着' })
    Check ($otherLines.Count -eq 0) "xlam 缺失期间无任何窗口被处理（多余行 $($otherLines.Count)）"

    '--- A3. 周期行心跳（信息项，不作断言）---'
    $cycleLines = @($lines | Where-Object { $_ -match 'cycle \d+: EXCEL7 窗口数=' })
    "    info: 4 周期内周期行 $($cycleLines.Count) 行（默认 30s 心跳，仅窗口数变化时记）"
    '    A1 独立日志全文：'
    $lines | ForEach-Object { '      ' + $_ }

    [void][AttachLinkCoexist]::DestroyWindow([AttachLinkCoexist]::Top)
}

# =====================================================================
'=== B. 真实 WPS：WPS 表格窗口必须被跳过（实例未被触碰）==='
$wpsType = $null
try { $wpsType = [type]::GetTypeFromProgID('ket.Application') } catch {}
if ($SkipRealWps -or -not $wpsType) {
    "SKIP: " + $(if ($SkipRealWps) { '按要求跳过' } else { '本机未安装 WPS 表格（ket.Application 未注册）' })
} else {
    # 安全闸：机器级常驻看门狗若仍是旧版（无宿主校验），会向测试实例弹窗——不跑
    $pfWd = 'C:\Program Files\AttachLink\watchdog.ps1'
    $deployedOld = $false
    if (Test-Path $pfWd) {
        $deployedOld = -not ([IO.File]::ReadAllText($pfWd, [Text.Encoding]::UTF8).Contains('仅附着 EXCEL.EXE'))
    }
    if ($deployedOld) {
        'SKIP: 机器级看门狗仍为旧版（' + $pfWd + ' 无宿主校验），真实 WPS 段不执行——请先重新部署修复版（AttachLinkSetup.exe → 鼎信诺管理员环境部署）'
    } else {
        $w = $null; $wb = $null; $instPids = @()
        try {
            $beforePids = @(Get-Process -Name et, wps -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
            $w = New-Object -ComObject ket.Application
            try { $w.DisplayAlerts = $false } catch {}
            function Release-Com($obj) {
                if ($null -eq $obj) { return }
                try { while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) -gt 0) {} } catch {}
            }
            function Get-WpsAddInInstalled($app) {
                # WPS 的 AddIns.Item('名字') 取不到时只报错，改为逐个比对名字。返回 'True' / 'False'；读不到返回 'n/a'。
                # 优先按下标遍历并逐个释放（不依赖 GC——被枚举过的集合只靠 GC 回收不可靠）；宿主不支持按下标取时退回枚举。
                $addins = $null
                try {
                    $addins = $app.AddIns
                    $n = [int]$addins.Count
                    for ($i = 1; $i -le $n; $i++) {
                        $ai = $null
                        try { $ai = $addins.Item($i); if ($ai.Name -ieq 'AttachLink.xlam') { return ([string]([bool]$ai.Installed)) } } finally { Release-Com $ai }
                    }
                    return 'n/a'
                } catch {
                    try { foreach ($a in $app.AddIns) { if ($a.Name -ieq 'AttachLink.xlam') { return ([string]([bool]$a.Installed)) } } } catch {}
                    return 'n/a'
                } finally { Release-Com $addins }
            }
            $instBefore = Get-WpsAddInInstalled $w
            $wb = $w.Workbooks.Add()          # 零工作簿实例没有 EXCEL7 窗口
            $deadline = (Get-Date).AddSeconds(15); $wins = @()
            while ($wins.Count -eq 0 -and (Get-Date) -lt $deadline) {
                Start-Sleep -Milliseconds 500
                # 只认「EXCEL7 属主是本次新出现的 et/wps 进程」的窗口：即测试实例自己的。
                # 必须按宿主进程名过滤——本机若有别的 Excel 实例在测试期间冒出新窗口，只按「新 pid」会把它误认成 WPS 窗口。
                $wins = @([AttachLinkCoexist]::FindXl() | Where-Object {
                        $f = $_.Split('|')
                        $ownerPid = [int]$f[1]
                        ($beforePids -notcontains $ownerPid) -and ((Get-Process -Id $ownerPid -ErrorAction SilentlyContinue).ProcessName -in @('et', 'wps'))
                    })
            }
            if ($wins.Count -eq 0) {
                'SKIP: WPS 实例未暴露顶层 XLMAIN/EXCEL7 窗口（当前窗口管理模式下看门狗本就看不到），无可验证对象'
            } else {
                # 清理名单只收「本次拉起」的进程：EXCEL7 属主必为新进程；XLMAIN 属主若为既有进程（用户的 WPS 壳）则不碰
                $instPids = @($wins | ForEach-Object { $_.Split('|') } | ForEach-Object {
                        [int]$_[1]
                        if ($beforePids -notcontains [int]$_[3]) { [int]$_[3] }
                    } | Select-Object -Unique)
                $books = @($wins | ForEach-Object { $_.Split('|')[0] })
                '    测试实例窗口：' + (($wins | ForEach-Object { $f = $_.Split('|'); 'EXCEL7=' + $f[0] + '(pid ' + $f[1] + ') under XLMAIN=' + $f[2] + '(pid ' + $f[3] + ')' }) -join '； ')
                $linesB = Invoke-Watchdog 'B' (Join-Path $base 'dummy.xlam') 4
                foreach ($b in $books) {
                    $hostPattern = '跳过 hwnd=' + $b + '：宿主进程=(et|wps)\.exe，仅附着 EXCEL\.EXE'
                    Check (@($linesB | Where-Object { $_ -match $hostPattern }).Count -eq 1) "WPS 文档窗口 hwnd=$b 被看门狗跳过"
                    Check (@($linesB | Where-Object { $_.Contains("hwnd=$b") -and $_ -notmatch $hostPattern }).Count -eq 0) "WPS 文档窗口 hwnd=$b 零接触"
                }
                $alive = $false
                try { $null = [string]$w.Caption; $alive = $true } catch {}
                Check $alive 'WPS 测试实例仍可响应 COM 调用（未被卡死）'
                $instAfter = Get-WpsAddInInstalled $w
                Check ($instBefore -ne 'n/a' -and $instAfter -ne 'n/a' -and $instAfter -eq $instBefore) "WPS 内 AttachLink 加载项 Installed 状态可读且未被改动（$instBefore → $instAfter；读不到按不通过处理）"
                # 按名字取：加载项工作簿不在 Workbooks 的序号枚举里，逐个遍历会漏判
                $xlamOpen = $false; $bkFound = $null
                try { $bkFound = $w.Workbooks.Item('AttachLink.xlam'); $xlamOpen = $true } catch { $xlamOpen = $false }
                if ($null -ne $bkFound) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($bkFound) } catch {} }
                Check (-not $xlamOpen) 'WPS 测试实例内未被打开 AttachLink.xlam'
                '    B 段独立日志：'
                $linesB | ForEach-Object { '      ' + $_ }
            }
        } catch {
            Check $false "真实 WPS 段执行异常: $($_.Exception.Message.Split("`n")[0])"
        } finally {
            $wbsW = $null
            # 判断 COM 对象变量是否还有值要用 $null -ne：已与 RCW 分离的对象放在 if ($x) 里求布尔值会抛 InvalidComObjectException
            if ($null -ne $w) {
                # 标记已保存 → Quit → 显式释放（循环到计数归零）。按下标遍历，不产生枚举器
                try { $wbsW = $w.Workbooks; $n = [int]$wbsW.Count; for ($i = 1; $i -le $n; $i++) { $bki = $null; try { $bki = $wbsW.Item($i); $bki.Saved = $true } catch {} finally { Release-Com $bki } } } catch {}
                try { $w.Quit() } catch {}
                Release-Com $wbsW
                Release-Com $w
            }
            if ($null -ne $wb) { Release-Com $wb }
            Remove-Variable w, wb, bk, wbsW, bki -ErrorAction SilentlyContinue
            [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
            Start-Sleep -Seconds 3
            # 默认只 Quit、不强杀：$instPids 只是「测试期间新出现的 et/wps 进程」，无法证明都属于本测试实例——
            # 用户若恰好在测试期间另开一个 WPS 窗口，强杀会丢失未保存内容。残留即判 FAIL 并打印 PID 交人工处理；
            # 确认测试机无人在用时，可用 -ForceCleanWps 明确授权清理。
            $left = @($instPids | Where-Object { $p = Get-Process -Id $_ -ErrorAction SilentlyContinue; $p -and $p.ProcessName -in @('et', 'wps') })
            if ($left.Count -gt 0) {
                $desc = ($left | ForEach-Object { 'pid=' + $_ + '(' + (Get-Process -Id $_ -ErrorAction SilentlyContinue).ProcessName + ')' }) -join ' '
                Check $false ('WPS 测试实例 Quit 后仍有进程存活: ' + $desc + '（未自动结束；确认不是自己在用的 WPS 后，可手动结束或加 -ForceCleanWps）')
                if ($ForceCleanWps) { foreach ($id in $left) { try { Stop-Process -Id $id -Force } catch {}; "    已按 -ForceCleanWps 结束 pid=$id" } }
            }
        }
    }
}

'--- 清理 ---'
try { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue } catch {}
"RESULT: " + $(if ($ok) { 'PASS' } else { 'FAIL' })
