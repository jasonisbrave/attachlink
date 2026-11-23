# =====================================================================
# AttachLink 全链路回归测试（非提权）
# ---------------------------------------------------------------------
# 场景：用 COM 自动化（CreateObject）拉起 Excel——精确复刻鼎信诺的启动
# 方式（跳过一切加载项初始化），再由看门狗完成外部装载。验证：
#   基线（装载前右键菜单不存在） → 看门狗附着 → 菜单 Tag 在位 → SelfTest
#   → 关后即退（被测看门狗全程存活；Quit 后脚本仍持着引用 5 秒）
#   → 预热生命周期（被测看门狗空转 50 秒后再附着）
# 用法：powershell -ExecutionPolicy Bypass -File test-chain.ps1
# 期望输出末行：RESULT: PASS
# 注意：本文件含中文，必须保存为 UTF-8 with BOM（PS5.1 按 ANSI 误读无
# BOM 文件会破坏解析，教训见 WORKLOG 根因 #11）。
#
# COM 引用纪律（v1.4.3）：本脚本碰过的每个 Excel 对象都显式 ReleaseComObject 到计数归零，不依赖 GC。
#   被枚举过的集合（AddIns、CommandBars.Controls）只靠 GC 回收不可靠——「关后即退」检查曾因此偶发假失败
#   （退不掉的引用是测试自己留下的，2026-11-24 对照实验证实），也让检查结果无法归因于看门狗。
#
# 关于 [6] 的定位：它验证的是「看门狗附着过、且在实例退出过程中一直活着时，实例仍能正常退出」。
#   它不是某个具体缺陷的必现探测器（旧版看门狗在这一步多半也能通过）；隐藏窗口规则的确定性回归在
#   scripts\test-watchdog-units.ps1 的 A2 / E 段。
# =====================================================================
# 可选参数：-AssumeSameVersion  跳过「常驻看门狗与被测脚本同版」的检查。仅在本机常驻看门狗已暂停、要单独验证
#           工作区版本时使用；常驻看门狗在运行时不要加——它会对同一实例做自己的动作，结果无法归因。
param([switch]$AssumeSameVersion)
$ErrorActionPreference = 'Stop'

$wdPath = Join-Path $PSScriptRoot 'watchdog.ps1'
if (-not (Test-Path $wdPath)) { throw "未找到 $wdPath" }
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class AttachLinkChainPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid); }'

function Release-Com($obj) {
    if ($null -eq $obj) { return }
    try { while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) -gt 0) { } } catch {}
}
function Test-Menu($app) {
    # 右键菜单 Tag 是否在位。集合与条目逐个释放，不留给 GC。
    $found = $false
    $bars = $null; $bar = $null; $ctrls = $null
    try {
        $bars = $app.CommandBars
        $bar = $bars.Item('Cell')
        $ctrls = $bar.Controls
        $n = [int]$ctrls.Count
        for ($i = 1; $i -le $n; $i++) {
            $ctrl = $null
            try { $ctrl = $ctrls.Item($i); if ($ctrl.Tag -eq 'AttachLinkCtx') { $found = $true } } finally { Release-Com $ctrl }
            if ($found) { break }
        }
    } finally { Release-Com $ctrls; Release-Com $bar; Release-Com $bars }
    return $found
}
function Get-AddInInstalled($app) {
    # AttachLink.xlam 的 Installed 状态；读不到返回 $null
    $result = $null
    $addins = $null
    try {
        $addins = $app.AddIns
        $n = [int]$addins.Count
        for ($i = 1; $i -le $n; $i++) {
            $ai = $null
            try { $ai = $addins.Item($i); if ($ai.Name -ieq 'AttachLink.xlam') { $result = [bool]$ai.Installed; break } } finally { Release-Com $ai }
        }
    } catch { $result = $null } finally { Release-Com $addins }
    return $result
}
function Save-AllWorkbooks($wbs) {
    # 按下标遍历（不产生枚举器），标记为已保存，免得 Quit 时弹保存提示
    $n = 0
    try { $n = [int]$wbs.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $w = $null
        try { $w = $wbs.Item($i); $w.Saved = $true } catch {} finally { Release-Com $w }
    }
}
function Wait-Exit([uint32]$ProcId, [int]$Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        if (-not (Get-Process -Id $ProcId -ErrorAction SilentlyContinue)) { return $true }
        Start-Sleep -Seconds 1
    }
    return $false
}
function Get-WdLogLines([string]$Log) {
    if ($Log -and (Test-Path $Log)) { return @([IO.File]::ReadAllLines($Log, [Text.Encoding]::UTF8)) }
    return @()
}
function Close-TestInstance($app, $books, [uint32]$ProcId) {
    # 收尾：Quit + 显式释放；仍不退的按进程号结束——只碰本测试自己拉起的实例
    if ($null -ne $app) { try { Save-AllWorkbooks $books } catch {}; try { $app.Quit() } catch {} }
    Release-Com $books
    Release-Com $app
    if ($ProcId -gt 0 -and (Get-Process -Id $ProcId -ErrorAction SilentlyContinue)) {
        if (-not (Wait-Exit $ProcId 10)) { try { Stop-Process -Id $ProcId -Force } catch {}; "    已清理残留测试实例 pid=$ProcId" }
    }
}

# 本机有机器级部署时用 Program Files 副本，否则用用户级库目录副本（与常驻看门狗用同一份，见 [3] 的说明）
$trustedXlam = 'C:\Program Files\AttachLink\AttachLink.xlam'
if (-not (Test-Path $trustedXlam)) { $trustedXlam = Join-Path $env:APPDATA 'Microsoft\AddIns\AttachLink.xlam' }

# 机器级常驻看门狗会附着同一个测试实例，它的行为会混进 [6] / [7] 的结果。只有它与被测脚本逐字节一致（哈希相同）时，
# 结果才归因于被测脚本；版本不一致（更旧或更新）就无法判断它的行为，两段生命周期检查跳过并明说原因。
$pfWd = 'C:\Program Files\AttachLink\watchdog.ps1'
$deployedDiffers = $false
if ((-not $AssumeSameVersion) -and (Test-Path $pfWd)) { $deployedDiffers = ((Get-FileHash $pfWd).Hash -ne (Get-FileHash $wdPath).Hash) }
if ($AssumeSameVersion) {
    '*** 警告：已加 -AssumeSameVersion，跳过了「机器级常驻看门狗与被测脚本同版」的检查。'
    '***        只有在常驻看门狗已暂停时才有意义；它在运行的话，下面的结果无法归因于被测脚本。'
}

$menu = $false; $installed = $null; $pinOk = $true; $warmOk = $true
$wd = $null; $wd2 = $null; $log1 = $null; $log2 = $null
$xl = $null; $wbs = $null; $testPid = [uint32]0
$xl2 = $null; $wbs2 = $null; $pid2 = [uint32]0
try {
    '[1] COM 自动化拉起 Excel（复刻鼎信诺启动方式，跳过加载项初始化）'
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $wbs = $xl.Workbooks                                   # 零工作簿实例没有 EXCEL7 窗口，先建一个
    $wb = $wbs.Add()
    Release-Com $wb; $wb = $null
    "    hwnd=$($xl.Hwnd)"

    # 测试实例的进程 ID（用于「关后即退」检查与残留清理）
    [void][AttachLinkChainPid]::GetWindowThreadProcessId([IntPtr][long]$xl.Hwnd, [ref]$testPid)
    "    pid=$testPid"

    '[2] 基线：该实例不应有右键菜单'
    $baseline = Test-Menu $xl
    if ($baseline) {
        "    baseline menu = True（机器级看门狗抢先装载所致，属预期竞态，不影响结论）"
    } else {
        "    baseline menu = False"
    }

    '[3] 启动被测看门狗（独立日志，运行到本测试结束；本机若有同版本常驻看门狗，[4] 的装载可能由它先完成——[1]~[6] 验证的是「看门狗机制能装载并释放」，被测脚本自身的归因由 [6] / [7] 负责）'
    # 被测看门狗与机器级常驻看门狗必须用同一份 xlam：二者会在 1 秒内先后附着同一实例，若各带一份不同的同名副本，
    # 后到者会再打开第二份、撞出「两份同名副本」的局面（实例随后退不干净）。
    $log1 = Join-Path $env:TEMP ('AttachLinkTestChain_wd_' + [guid]::NewGuid().ToString('N') + '.log')
    $wd = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $wdPath), '-LogFile', ('"{0}"' -f $log1), '-XlamPath', ('"{0}"' -f $trustedXlam)) -PassThru -WindowStyle Hidden

    '[4] 轮询验证菜单与装载状态（最多 25 秒）'
    $deadline = (Get-Date).AddSeconds(25)
    while ((Get-Date) -lt $deadline -and -not $menu) {
        Start-Sleep -Milliseconds 1000
        try {
            $menu = Test-Menu $xl
            $installed = Get-AddInInstalled $xl
        } catch { Start-Sleep -Milliseconds 500 }
    }
    "    AddIn Installed = $installed"
    "    右键菜单 Tag 在位 = $menu"

    if ($menu) {
        '[5] 冒烟：在目标实例中运行 SelfTest'
        $t = $xl.Run('AttachLink_SelfTest')
        "    AttachLink_SelfTest = $t"
    }

    '[6] 清理与「关后即退」检查：被测看门狗全程存活；Quit 后本脚本仍持着 Application / Workbooks 引用 5 秒'
    # 退出过程中客户端仍持着引用 5 秒——真实客户端（鼎信诺、各类脚本）常在 Quit 之后才慢慢放引用。这几秒里看门狗的
    # 周期照常运行（加载项自身的隐藏窗口通常就是在这期间第一次被它看到）：它不得钉住实例，也不得打断退出。
    if ($deployedDiffers) {
        '    SKIP 关后即退检查：机器级常驻看门狗与被测脚本版本不一致（它的行为会混进结果），请先重新部署'
        Close-TestInstance $xl $wbs $testPid
        $xl = $null; $wbs = $null
    } else {
        # 被测看门狗必须参与了本实例的处理（独立日志里有本实例 pid 的记录），并且在检查全程存活——
        # 否则「实例能退出」说明不了任何问题：没有看门狗参与时实例当然能退出。
        $deadline = (Get-Date).AddSeconds(15); $wdLines = @()
        while ((Get-Date) -lt $deadline -and $wdLines.Count -eq 0) {
            Start-Sleep -Milliseconds 500
            $wdLines = @(Get-WdLogLines $log1 | Where-Object { $_ -match ('pid=' + $testPid + '\b') })
        }
        if (-not $wd -or $wd.HasExited) {
            $pinOk = $false; '    FAIL: 被测看门狗在 [6] 开始前已退出，本次关后即退检查无效'
        } elseif ($wdLines.Count -eq 0) {
            $pinOk = $false; "    FAIL: 被测看门狗的日志里没有本实例（pid=$testPid）的记录——它没有参与处理，本次关后即退检查无效"
        } else {
            '    被测看门狗对本实例的记录（截至 Quit 前）：'
            $wdLines | ForEach-Object { '      ' + $_ }
            try { Save-AllWorkbooks $wbs } catch {}
            $xl.Quit()
            Start-Sleep -Seconds 5
            Release-Com $wbs
            Release-Com $xl
            $wbs = $null; $xl = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
            $exited = Wait-Exit $testPid 20
            if ($wd.HasExited) { $pinOk = $false; '    FAIL: 被测看门狗在退出检查期间退出，检查无效' }
            elseif ($exited) {
                $after = @(Get-WdLogLines $log1 | Where-Object { $_ -match ('pid=' + $testPid + '\b') })
                '    pass: 测试实例在 Quit（含 5 秒持引用）后 25 秒内退出，被测看门狗全程存活（未钉住进程、未打断退出）'
                $after | Where-Object { $wdLines -notcontains $_ } | ForEach-Object { '      （退出期间）' + $_ }
            } else { $pinOk = $false; '    FAIL: 测试实例 Quit（释放引用后 20 秒，合计 25 秒）仍未退出（看门狗持有宿主引用或打断了退出）' }
        }
        Close-TestInstance $xl $wbs $testPid      # FAIL / 异常路径的兜底：正常路径上 $xl 已为 $null，只做残留清理
        $xl = $null; $wbs = $null
    }
    if ($wd -and -not $wd.HasExited) { try { Stop-Process -Id $wd.Id -Force } catch {} }
    Remove-Item $log1 -Force -ErrorAction SilentlyContinue

    '[7] 预热生命周期检查：看门狗先空转再附着（走「已在位」分支），且看门狗仍在运行时 Excel 须能退出'
    # 背景（WORKLOG #27 会诊）：常驻看门狗运行一段时间后附着时，链式 COM 调用产生的中间对象一次 GC 回收不掉，
    # 会把宿主 Excel 钉住不退。上面 [3]~[6] 是冷启动附着，恰好测不出来。
    if ($deployedDiffers) {
        '    SKIP 预热生命周期检查：机器级常驻看门狗与被测脚本版本不一致（它的行为会混进结果），请先重新部署'
    } else {
        $log2 = Join-Path $env:TEMP ('AttachLinkTestChain_wd_' + [guid]::NewGuid().ToString('N') + '.log')
        $wd2 = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $wdPath), '-LogFile', ('"{0}"' -f $log2), '-XlamPath', ('"{0}"' -f $trustedXlam)) -PassThru -WindowStyle Hidden
        '    被测看门狗空转 50 秒…'
        Start-Sleep -Seconds 50
        $xl2 = New-Object -ComObject Excel.Application        # 隐藏实例
        $wbs2 = $xl2.Workbooks
        $wb2 = $wbs2.Add(); Release-Com $wb2; $wb2 = $null
        $wb2 = $wbs2.Open($trustedXlam); Release-Com $wb2; $wb2 = $null   # 模拟正常启动时已装载加载项
        $null = $xl2.Run('AttachLink_Initialize')
        [void][AttachLinkChainPid]::GetWindowThreadProcessId([IntPtr][long]$xl2.Hwnd, [ref]$pid2)
        $deadline = (Get-Date).AddSeconds(25); $inplace = 0
        # 只认本实例（pid=$pid2）的「已在位」记录：本机别的 Excel 窗口也会让被测看门狗写同类日志行
        $pattern = '已在位，无需装载: hwnd=\d+ pid=' + $pid2 + '\b'
        while ((Get-Date) -lt $deadline -and $inplace -lt 1) {
            Start-Sleep -Milliseconds 500
            $inplace = @(Get-WdLogLines $log2 | Where-Object { $_ -match $pattern }).Count
        }
        if ($inplace -ge 1) { '    pass: 预热后的看门狗对本实例走了「已在位，无需装载」分支' }
        else { $warmOk = $false; "    FAIL: 未见本实例（pid=$pid2）的「已在位，无需装载」日志，未覆盖到目标分支" }
        try { Save-AllWorkbooks $wbs2 } catch {}
        $xl2.Quit()
        Release-Com $wbs2
        Release-Com $xl2
        $wbs2 = $null; $xl2 = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
        if ($wd2.HasExited) {
            # 被测看门狗提前退出：它持有的引用随进程释放，「实例能退出」不再说明任何问题——检查无效，按不合格处理
            $warmOk = $false
            '    FAIL: 被测看门狗在退出检查前已退出，本次预热生命周期检查无效'
        } else {
            $exited = Wait-Exit $pid2 20
            if ($exited -and $wd2.HasExited) { $warmOk = $false; '    FAIL: Excel 已退出，但被测看门狗在等待期间退出，检查无效' }
            elseif ($exited) { '    pass: Excel 在 Quit 后 20 秒内退出，且被测看门狗仍在运行（未被钉住）' }
            else { $warmOk = $false; '    FAIL: Excel 在 Quit 后 20 秒仍未退出（预热后的看门狗钉住了进程）' }
        }
        Close-TestInstance $xl2 $wbs2 $pid2      # FAIL 路径的兜底；正常路径上 $xl2 已为 $null
        $xl2 = $null; $wbs2 = $null
    }
} finally {
    '--- 清理 ---'
    Close-TestInstance $xl $wbs $testPid
    Close-TestInstance $xl2 $wbs2 $pid2
    foreach ($w in @($wd2, $wd)) { if ($w -and -not $w.HasExited) { try { Stop-Process -Id $w.Id -Force } catch {} } }
    Remove-Item $log1, $log2 -Force -ErrorAction SilentlyContinue
}
"RESULT: " + $(if ($menu -and $installed -and $pinOk -and $warmOk) { 'PASS' } else { 'FAIL' })
