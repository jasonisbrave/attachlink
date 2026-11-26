# =====================================================================
# AttachLink 全链路回归测试（非提权）
# ---------------------------------------------------------------------
# 场景：用 COM 自动化（CreateObject）拉起 Excel——精确复刻鼎信诺的启动
# 方式（跳过一切加载项初始化），再由看门狗完成外部装载。验证：
#   基线（装载前右键菜单不存在） → 看门狗附着 → 菜单 Tag 在位 → SelfTest
#   → 会话内已装载（在位副本路径 = 配置路径；v1.6.0 R1 起不再要求 AddIn Installed）
#   → 登记零差异（实例退出后，HKCU OPEN/OPENn 与 Add-in Manager 的 AttachLink 相关值
#     与测试前完全一致；实例内 AddIns 的 Installed 不被看门狗改动）
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
function Remove-MenuTagControls($app) {
    # 删除右键菜单上所有带 AttachLinkCtx 标签的控件（仅用于清理命令栏缓存残留；集合与条目逐个释放）
    $bars = $null; $bar = $null; $ctrls = $null
    try {
        $bars = $app.CommandBars
        $bar = $bars.Item('Cell')
        $ctrls = $bar.Controls
        for ($i = [int]$ctrls.Count; $i -ge 1; $i--) {
            $ctrl = $null
            try { $ctrl = $ctrls.Item($i); if ($ctrl.Tag -eq 'AttachLinkCtx') { $ctrl.Delete() } } finally { Release-Com $ctrl }
        }
    } finally { Release-Com $ctrls; Release-Com $bar; Release-Com $bars }
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
function Get-SessionLoadedPath($wbs) {
    # 会话内装载判定（v1.6.0 R1）：按名字取 AttachLink.xlam 工作簿，返回其完整路径；
    # 不在位 / 路径读不到返回 $null。取到的工作簿读完即释放（不依赖 GC）。
    $w = $null
    try { $w = $wbs.Item('AttachLink.xlam'); return [string]$w.FullName } catch { return $null } finally { Release-Com $w }
}
function Read-HostUiState($app) {
    # 读宿主的 DisplayAlerts / ScreenUpdating。宿主正忙时（例如看门狗正在装载）COM 调用会被拒绝：
    # 重试几次；仍读不到返回 $null——调用方必须按「读不到」处理，不能当成 True 或 False。
    for ($k = 0; $k -lt 8; $k++) {
        try {
            $alertsValue = $app.DisplayAlerts
            $screenValue = $app.ScreenUpdating
            if ($null -ne $alertsValue -and $null -ne $screenValue) { return @{ alerts = [bool]$alertsValue; screen = [bool]$screenValue } }
        } catch {}
        Start-Sleep -Milliseconds 150
    }
    return $null
}
function Get-AttachLinkRegistrySnapshot {
    # 加载项列表登记的「与 AttachLink 相关」快照（名称+数据，排序后比对，只读）：
    # HKCU 各 Office 版本 Excel\Options 的 OPEN/OPENn 中含 AttachLink 的值，
    # 以及同级 Add-in Manager 下名字含 AttachLink 的值。R1（只装载不登记）验收：
    # 看门狗装载过的实例退出后，此处必须与测试前完全一致（WORKLOG #30 的登记残留被绕开）。
    $snap = New-Object System.Collections.Generic.List[string]
    foreach ($ver in @('16.0', '15.0', '14.0')) {
        $optKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
        if (Test-Path $optKey) {
            foreach ($p in (Get-ItemProperty -LiteralPath $optKey).PSObject.Properties) {
                if ($p.Name -match '^OPEN(\d+)?$' -and $p.Value -is [string] -and $p.Value -match 'AttachLink') {
                    $snap.Add("OPT $ver $($p.Name)=[$($p.Value)]")
                }
            }
        }
        $amKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Add-in Manager"
        if (Test-Path $amKey) {
            foreach ($n in (Get-Item -LiteralPath $amKey).GetValueNames()) {
                if ($n -match 'AttachLink') {
                    $v = (Get-ItemProperty -LiteralPath $amKey -Name $n).$n
                    $snap.Add("AM $ver [$n]=[$v]")
                }
            }
        }
    }
    return @($snap | Sort-Object)
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

$menu = $false; $installed = $null; $sessionLoaded = $false; $loadedPath = $null; $selfTestOk = $false
$pinOk = $true; $warmOk = $true; $regOk = $true; $installedUntouched = $true; $installedBefore = $null; $hostUiUntouched = $true; $hostUiPreloaded = $false; $hostUiRace = $false; $hostUiLoadedEarly = $false
$snapBefore = Get-AttachLinkRegistrySnapshot          # 登记零差异（R1）的「前」快照：实例创建之前
$wd = $null; $wd2 = $null; $log1 = $null; $log2 = $null
$xl = $null; $wbs = $null; $testPid = [uint32]0
$xl2 = $null; $wbs2 = $null; $pid2 = [uint32]0
try {
    '[1] COM 自动化拉起 Excel（复刻鼎信诺启动方式，跳过加载项初始化）'
    $xl = New-Object -ComObject Excel.Application
    # 宿主界面设置（R1「装载不留副作用」的断言用）：模拟别的程序拉起 Excel 后自己关掉了警告与屏幕刷新。必须在出现
    # 工作簿窗口【之前】设——零工作簿实例没有 EXCEL7 窗口，看门狗此时不可能附着；这样看门狗之后的任何快照都晚于
    # 这里的设置，从构造上排除「带着设置之前的快照装载、再把旧值恢复回去」的竞态（2026-11-26 实测遇到过一次）。
    # Excel 在加载项的 Workbook_Open 事件处理结束时会把这两项一律复位为 True，看门狗装载后必须恢复成这里的原值（[4] 断言）。
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
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
        # 三态区分（真三态，不是「一切异常都算死菜单」）：
        #   notfound = Workbooks.Item 抛 DISP_E_BADINDEX（0x8002000B）→ 加载项工作簿不在位 → 命令栏缓存残留（死菜单），清理；
        #   found    = 取到工作簿 → 机器级看门狗抢先装载（属预期竞态，不清理）；
        #   unknown  = 其余异常（宿主正忙等）→ 无法确认在位情况 → 不清理、如实打印，按已有菜单继续。
        # Excel 会把非临时的内置命令栏控件写入 %APPDATA%\Microsoft\Excel\Excel15.xlb，此后每个新实例
        # （含 COM 自动化拉起的）都带着它启动；看门狗据菜单 Tag 判「已在位」而不装载（2026-11-24 受控复现）。
        # 残留来自在实例里创建非临时 AttachLinkCtx 控件并正常退出的测试 / 探针——测试与探针不得这样做。
        $wbState = 'unknown'
        $w0 = $null
        try { $w0 = $wbs.Item('AttachLink.xlam'); if ($null -ne $w0) { $wbState = 'found' } } catch {
            $hr = 0
            try { $hr = [int]$_.Exception.InnerException.HResult } catch {}
            if ($hr -eq 0) { try { $hr = [int]$_.Exception.HResult } catch {} }
            if ($hr -eq -2147352565) { $wbState = 'notfound' }   # DISP_E_BADINDEX (0x8002000B)
        } finally { Release-Com $w0 }
        if ($wbState -eq 'notfound') {
            "    baseline menu = True，但加载项工作簿不在位 → 命令栏缓存残留的 AttachLinkCtx 控件（死菜单）"
            "    已从本实例移除残留控件（本实例随后正常退出会把缓存写干净）；看门狗下个周期应正常装载"
            Remove-MenuTagControls $xl
            $baseline = $false
        } elseif ($wbState -eq 'found') {
            "    baseline menu = True（机器级看门狗抢先装载所致，属预期竞态，不影响结论）"
            $hostUiPreloaded = $true      # 加载项在基线检查时已由机器级看门狗装载：两项设置早于窗口出现、快照必然晚于设置，[4] 照常断言（输出里注明）
        } else {
            "    无法确认在位情况（Workbooks.Item 返回 0x{0:X8}），按已有菜单继续" -f $hr
        }
    } else {
        "    baseline menu = False"
    }

    # 测试前读一次登记位（R1 断言：看门狗不再改它；登记位 ≠ 装载态，只作对照不做装载判据）
    $installedBefore = Get-AddInInstalled $xl

    # 宿主界面设置已在 [1] 建实例时、出现工作簿窗口之前设好。这里观察最多 6 秒：若这期间加载项由机器级常驻看门狗
    # 装载完成，这次装载就是有效证据的来源（它的快照必然晚于设置）；6 秒内没人来装载，则由 [3] 启动的被测看门狗来装载。
    # 两项在装载后是否保持原值，统一由 [4] 判定。
    $settleUntil = (Get-Date).AddSeconds(6)
    $uiReads = 0
    while ((Get-Date) -lt $settleUntil) {
        Start-Sleep -Milliseconds 300
        $uiNow = Read-HostUiState $xl
        if ($null -eq $uiNow) { continue }        # 宿主正忙读不到（多半是看门狗在装载）：下一轮再看
        $uiReads++
        if (Get-SessionLoadedPath $wbs) { $hostUiLoadedEarly = $true; break }
    }
    if ($uiReads -eq 0) { $hostUiRace = $true }    # 整个观察窗口一次都没读到（宿主一直忙）：无法确认设置是否生效，按没有判别力处理
    "    观察窗口：成功读取 $uiReads 次；由常驻看门狗在窗口内装载=$hostUiLoadedEarly"

    '[3] 启动被测看门狗（独立日志，运行到本测试结束；本机若有同版本常驻看门狗，[4] 的装载可能由它先完成——[1]~[6] 验证的是「看门狗机制能装载并释放」，被测脚本自身的归因由 [6] / [7] 负责）'
    # 被测看门狗与机器级常驻看门狗必须用同一份 xlam：二者会在 1 秒内先后附着同一实例，若各带一份不同的同名副本，
    # 后到者会再打开第二份、撞出「两份同名副本」的局面（实例随后退不干净）。
    $log1 = Join-Path $env:TEMP ('AttachLinkTestChain_wd_' + [guid]::NewGuid().ToString('N') + '.log')
    $wd = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $wdPath), '-LogFile', ('"{0}"' -f $log1), '-XlamPath', ('"{0}"' -f $trustedXlam)) -PassThru -WindowStyle Hidden

    '[4] 轮询验证菜单与装载状态（最多 25 秒）'
    $deadline = (Get-Date).AddSeconds(25)
    while ((Get-Date) -lt $deadline -and -not ($menu -and $sessionLoaded)) {
        Start-Sleep -Milliseconds 1000
        try {
            $menu = Test-Menu $xl
            $loadedPath = Get-SessionLoadedPath $wbs
            $installed = Get-AddInInstalled $xl
            if ($loadedPath -and ($loadedPath -ieq $trustedXlam)) { $sessionLoaded = $true }
        } catch { Start-Sleep -Milliseconds 500 }
    }
    $installedAfter = $installed
    "    会话内已装载（在位副本 = 配置路径）= $sessionLoaded" + $(if ($loadedPath) { "（在位路径: $loadedPath）" } else { "（尚未取到在位副本）" })
    "    AddIn Installed（登记位，看门狗不应改动）: 测试前 = $installedBefore，装载后 = $installedAfter"
    "    右键菜单 Tag 在位 = $menu"

    # 装载前后宿主的 DisplayAlerts / ScreenUpdating 必须不变（上面已设为 False）：看门狗须恢复原值——否则别的
    # 程序关掉的警告会被重新打开，一个「是否替换」提示框就能把隐藏的自动化实例卡死（实测卡了 22 分钟，WORKLOG #34）
    if ($deployedDiffers) {
        '    SKIP 宿主设置检查：机器级常驻看门狗与被测脚本版本不一致（装载可能由它完成），请先重新部署'
    } elseif ($hostUiRace) {
        '    SKIP 宿主设置检查：设置后的 6 秒观察窗口内一次都没能读到宿主设置（宿主一直忙），无法确认设置是否生效'
    } else {
        $uiAfter = Read-HostUiState $xl
        if ($null -eq $uiAfter) {
            '    SKIP 宿主设置检查：装载后读不到宿主设置（宿主正忙），本断言没有判别力'
        } else {
            $hostAlerts = $uiAfter.alerts; $hostScreen = $uiAfter.screen
            "    装载后宿主 DisplayAlerts=$hostAlerts ScreenUpdating=$hostScreen（建实例时、出现工作簿窗口之前均设为 False）" + $(if ($hostUiPreloaded) { '（基线检查前已由机器级常驻看门狗装载）' } elseif ($hostUiLoadedEarly) { '（由机器级常驻看门狗装载）' } else { '' })
            if ($hostAlerts -eq $false -and $hostScreen -eq $false) {
                '    pass: 装载前后宿主的 DisplayAlerts / ScreenUpdating 不变（看门狗恢复原值，不把宿主关掉的警告重新打开）'
            } else {
                $hostUiUntouched = $false
                "    FAIL: 看门狗装载后宿主设置被改动（DisplayAlerts=$hostAlerts ScreenUpdating=$hostScreen，装载前均为 False）"
            }
        }
    }

    if ($menu) {
        '[5] 冒烟：在目标实例中运行 SelfTest'
        $t = $xl.Run('AttachLink_SelfTest')
        $selfTestOk = [bool]$t
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

        # ---- R1 只装载不登记：实例已退出（登记此刻落盘），加载项列表必须与测试前完全一致 ----
        $snapAfter6 = Get-AttachLinkRegistrySnapshot
        $regDiff = @(Compare-Object $snapBefore $snapAfter6)
        if ($regDiff.Count -eq 0) {
            '    pass: 登记零差异（实例退出后，OPEN/OPENn 与 Add-in Manager 的 AttachLink 相关值与测试前完全一致）'
        } else {
            $regOk = $false
            '    FAIL: 装载会话退出后加载项列表发生变动（看门狗应只装载不登记，WORKLOG #30 残留应被绕开）：'
            $regDiff | ForEach-Object { '      ' + $_.InputObject }
        }
        if ($installedBefore -eq $installedAfter) {
            "    pass: 实例内 AddIns 的 Installed 未被看门狗改动（$installedBefore → $installedAfter）"
        } else {
            $installedUntouched = $false
            "    FAIL: 实例内 AddIns 的 Installed 被改动（$installedBefore → $installedAfter）——看门狗不应再碰登记"
        }
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

        # R1 终态复核：预热实例也已退出，加载项列表仍须与测试前一致
        $snapFinal = Get-AttachLinkRegistrySnapshot
        $regDiffF = @(Compare-Object $snapBefore $snapFinal)
        if ($regDiffF.Count -eq 0) {
            '    pass: 终态登记零差异（预热实例退出后，AttachLink 相关登记仍与测试前完全一致）'
        } else {
            $regOk = $false
            '    FAIL: 预热实例退出后加载项列表发生变动（看门狗应只装载不登记）：'
            $regDiffF | ForEach-Object { '      ' + $_.InputObject }
        }
    }
} finally {
    '--- 清理 ---'
    Close-TestInstance $xl $wbs $testPid
    Close-TestInstance $xl2 $wbs2 $pid2
    foreach ($w in @($wd2, $wd)) { if ($w -and -not $w.HasExited) { try { Stop-Process -Id $w.Id -Force } catch {} } }
    # 提前终止时 $log2 可能尚未赋值：先滤掉空值，否则 Remove-Item 会报参数绑定错误（-ErrorAction 管不住）
    Remove-Item (@($log1, $log2) | Where-Object { $_ }) -Force -ErrorAction SilentlyContinue
}
if ($menu -and $sessionLoaded -and $selfTestOk -and $pinOk -and $warmOk -and $regOk -and $installedUntouched -and $hostUiUntouched) { "RESULT: PASS"; exit 0 } else { "RESULT: FAIL"; exit 1 }
