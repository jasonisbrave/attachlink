# =====================================================================
# AttachLink 看门狗「在位判定加固」回归（v1.6.0 R2，含死菜单根除）：死菜单（命令栏缓存残留）场景
# ---------------------------------------------------------------------
# 验证：Excel 命令栏里残留了非临时的 AttachLinkCtx 控件时（菜单在位、加载项工作簿不在位——「死菜单」，
# 见开发交接 3.3 第 9 条），看门狗不再据此判「已在位」，而是：
#   ① 识别为死菜单；
#   ② 装载之前先【从外部】删除残留控件——根除。只靠装载后的初始化在会话内删除不够：实测非临时控件在
#      Workbook_Open 里被加载项删除后，实例正常退出时仍会被写回 Excel15.xlb，死菜单在每个新实例重现
#      （2026-11-26，WORKLOG #34）；
#   ③ 打开配置路径副本并初始化，菜单可用；
#   ④ 实例正常退出后，新实例不再带着该控件启动（缓存已清）。
#
# 做法：COM 自动化拉起隐藏 Excel（复刻鼎信诺启动方式，加载项未装载），建工作簿出 EXCEL7 窗口后【立刻】
# 在 Cell 命令栏放一个【非临时】的 AttachLinkCtx 占位控件，抢在看门狗首次附着之前（隐藏实例有一个周期的宽限）；
# 随后在看门狗日志里等「按未装载处理」「已删除死菜单控件」「已装载并初始化」三行，并在实例内核实 SelfTest
# 可用、一级菜单恰好是 5 个正式项。最后新起实例核实缓存已清。
#
# v1.6.2 起常驻看门狗不再附着隐藏实例：装载改由本脚本自己启动的被测看门狗完成（带 -AttachHiddenInstances 回到
# 旧规则，独立日志）；下文所说「看门狗日志」自闸检查之后均指这份独立日志，常驻看门狗的主日志只用于运行前的闸。
#
# ■ 本脚本是「任何测试 / 探针都不得创建非临时的 AttachLinkCtx 控件」这条红线（开发交接 3.3 第 9 条与
#   3.5 节）的【唯一受控例外】——它测的就是非临时残留；换成临时控件，第 ④ 步的断言就没有判别力
#   （临时控件本来就不会进缓存，2026-11-26 实测）。受控措施：
#     ① 运行前先核实缓存干净（[0]）；不干净则先清掉（删除 + 正常退出会重写缓存）并判 FAIL，不从脏缓存起测；
#     ② 占位控件一放进命令栏就开始跟踪：先设标签、再设标题。收尾时按「带 AttachLinkCtx 标签、但标题不是
#        正式菜单项」识别并从外部删除（时序未命中、看门狗没有处理时它还在——必须删掉再正常退出）；
#     ③ 认不出它（标签没设成功）或删不干净时，直接结束实例而不正常退出——被强杀的 Excel 不会写命令栏缓存；
#     ④ 结束后再核实缓存干净（[2]）；不干净则清掉并判 FAIL；
#     ⑤ 读控件数、读标签读不到时一律按「无法确认」处理（抛异常），不当成「没有」。
#   任何路径下都不得把占位控件留在缓存里。
#
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test-deadmenu.ps1
#       换 .NET 版被测看门狗：powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test-deadmenu.ps1 -WatchdogExe dist\AttachLink.Watchdog.exe
# 期望末行：RESULT: PASS
# 前置：机器级常驻看门狗与被测脚本同版（哈希一致，否则 SKIP——版本不一致时它会对同一实例做自己的动作，
#       结果无法归因）；看门狗日志在刷新；运行前没有 EXCEL.EXE。
# 可选参数：-AssumeSameVersion  跳过同版检查（仅在常驻看门狗已暂停、单独验证工作区版本时使用）；
#           -WaitSeconds <秒>  等看门狗处理的上限（默认 30）；
#           -WatchdogExe <路径>  用该可执行文件当被测看门狗（相对路径按仓库根解析；给 dist\AttachLink.Watchdog.exe
#           即换 .NET 版，日志措辞与 PowerShell 版逐项等价，仅少「Add-Type 完成」一行）。不给就是原来的 PowerShell 版。
# 注意：含中文，必须 UTF-8 with BOM。本脚本只按自己记录的进程号结束自己拉起的实例，不会结束其它 Excel。
# =====================================================================
param([switch]$AssumeSameVersion, [int]$WaitSeconds = 30, [string]$WatchdogExe = '')
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
# -WatchdogExe：相对路径按仓库根（scripts 的上一级）解析；文件不存在直接失败，不让测试带着空开关跑。
if ($WatchdogExe) {
    if (-not [IO.Path]::IsPathRooted($WatchdogExe)) { $WatchdogExe = Join-Path (Split-Path -Parent $PSScriptRoot) $WatchdogExe }
    if (-not (Test-Path $WatchdogExe)) { throw "未找到被测看门狗可执行文件：$WatchdogExe" }
}

$wdPath = Join-Path $PSScriptRoot 'watchdog.ps1'
if (-not (Test-Path $wdPath)) { throw "未找到 $wdPath" }
$pfWd = 'C:\Program Files\AttachLink\watchdog.ps1'
$wdLog = Join-Path $env:TEMP 'AttachLinkWatchdog.log'
$trustedXlam = 'C:\Program Files\AttachLink\AttachLink.xlam'
if (-not (Test-Path $trustedXlam)) { $trustedXlam = Join-Path $env:APPDATA 'Microsoft\AddIns\AttachLink.xlam' }
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class AttachLinkDeadMenuPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid); }'

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; Write-Host "FAIL: $label" } else { Write-Host "pass: $label" } }
function Say($m) { Write-Host $m }
function Release-Com($obj) {
    if ($null -eq $obj) { return }
    try { while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) -gt 0) { } } catch {}
}
function Get-CellControls($app) {
    # 取 Cell 命令栏的一级控件集合，存进哈希表返回（调用方负责 Release-Com 其中的 bars / bar / ctrls 三个对象）。
    # 读不到控件数时抛异常（PS 5.1 下属性读取失败可能不上抛而静默得到空值；真实 Excel 里这个数一定读得到）
    $h = @{ bars = $null; bar = $null; ctrls = $null }
    try {
        $h.bars = $app.CommandBars
        $h.bar = $h.bars.Item('Cell')
        $h.ctrls = $h.bar.Controls
        $nRaw = $h.ctrls.Count
        if ($null -eq $nRaw) { throw '命令栏控件数读不到（宿主繁忙）' }
        return $h
    } catch {
        Release-Com $h.ctrls; Release-Com $h.bar; Release-Com $h.bars
        throw
    }
}
function Get-TaggedTop($app) {
    # Cell 命令栏【一级】控件里带 AttachLinkCtx 标签的标题（按出现顺序；集合与条目逐个释放，不留给 GC）。
    # 读不到控件数 / 标签时抛异常：读不到不等于没有（真实 Excel 里无标签控件的 Tag 是空串而不是 $null）
    $caps = @()
    $h = Get-CellControls $app
    $ctrl = $null
    try {
        $n = [int]$h.ctrls.Count
        for ($i = 1; $i -le $n; $i++) {
            $ctrl = $null
            try {
                $ctrl = $h.ctrls.Item($i)
                $tag = $ctrl.Tag
                if ($null -eq $tag) { throw ('控件 Tag 读不到（宿主繁忙），下标 ' + $i) }
                if ($tag -eq 'AttachLinkCtx') { $caps += [string]$ctrl.Caption }
            } finally { Release-Com $ctrl }
        }
    } finally { Release-Com $ctrl; Release-Com $h.ctrls; Release-Com $h.bar; Release-Com $h.bars }
    # 不用一元逗号包裹返回值：空数组包一层后到调用方会变成「含一个空数组的单元素数组」（实测）
    return $caps
}
function Remove-TaggedControls($app) {
    # 从外部删除 Cell 命令栏上所有带 AttachLinkCtx 标签的一级控件（只用于缓存检查实例：那里没有正式菜单，
    # 带标签的都是残留）。倒序下标遍历（删除不改变未遍历部分的下标）；集合与条目逐个释放。返回删除个数。
    $removed = 0
    $h = Get-CellControls $app
    $ctrl = $null
    try {
        $n = [int]$h.ctrls.Count
        for ($i = $n; $i -ge 1; $i--) {
            $ctrl = $null
            try {
                $ctrl = $h.ctrls.Item($i)
                $tag = $ctrl.Tag
                if ($null -eq $tag) { throw ('控件 Tag 读不到（宿主繁忙），下标 ' + $i) }
                if ($tag -eq 'AttachLinkCtx') { [void]$ctrl.Delete(); $removed++ }
            } finally { Release-Com $ctrl }
        }
    } finally { Release-Com $ctrl; Release-Com $h.ctrls; Release-Com $h.bar; Release-Com $h.bars }
    return $removed
}
function Remove-PlaceholderControls($app, [string[]]$keepCaptions) {
    # 收尾清理：从外部删除「带 AttachLinkCtx 标签、但标题不是正式菜单项」的一级控件——即测试放入的占位控件
    #（无论它的标题是否设置成功；正式菜单项的标题见 $keepCaptions）。删除后再数一遍，返回 @{ removed = n; left = m }。
    # 读不到时抛异常（调用方按「无法确认」处理）。
    $removed = 0; $left = 0
    $h = Get-CellControls $app
    $ctrl = $null
    try {
        $n = [int]$h.ctrls.Count
        for ($i = $n; $i -ge 1; $i--) {
            $ctrl = $null
            try {
                $ctrl = $h.ctrls.Item($i)
                $tag = $ctrl.Tag
                if ($null -eq $tag) { throw ('控件 Tag 读不到（宿主繁忙），下标 ' + $i) }
                if ($tag -eq 'AttachLinkCtx' -and ($keepCaptions -notcontains ([string]$ctrl.Caption))) { [void]$ctrl.Delete(); $removed++ }
            } finally { Release-Com $ctrl }
        }
        $n2 = [int]$h.ctrls.Count
        for ($i = 1; $i -le $n2; $i++) {
            $ctrl = $null
            try {
                $ctrl = $h.ctrls.Item($i)
                $tag2 = $ctrl.Tag
                if ($null -eq $tag2) { throw ('控件 Tag 读不到（宿主繁忙），下标 ' + $i) }
                if ($tag2 -eq 'AttachLinkCtx' -and ($keepCaptions -notcontains ([string]$ctrl.Caption))) { $left++ }
            } finally { Release-Com $ctrl }
        }
    } finally { Release-Com $ctrl; Release-Com $h.ctrls; Release-Com $h.bar; Release-Com $h.bars }
    return @{ removed = $removed; left = $left }
}
function Get-SessionLoadedPath($wbs) {
    # 按名字取 AttachLink.xlam 工作簿的完整路径；不在位 / 读不到返回 $null（读完即释放）
    $w = $null
    try { $w = $wbs.Item('AttachLink.xlam'); return [string]$w.FullName } catch { return $null } finally { Release-Com $w }
}
function Save-AllWorkbooks($wbs) {
    $n = 0
    try { $n = [int]$wbs.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $w = $null
        try { $w = $wbs.Item($i); [void]$w.Saved; $w.Saved = $true } catch {} finally { Release-Com $w }
    }
}
function Test-Alive([uint32]$ProcId) {
    # 进程是否仍在：只把「进程号已不存在」判为已退出，其余查询错误按「仍在」处理（保守，不把查不到当成已退出）
    try { $null = [System.Diagnostics.Process]::GetProcessById([int]$ProcId); return $true }
    catch [System.ArgumentException] { return $false }
    catch { return $true }
}
function Wait-Exit([uint32]$ProcId, [int]$Seconds) {
    $waitUntil = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $waitUntil) {
        if (-not (Test-Alive $ProcId)) { return $true }
        Start-Sleep -Milliseconds 500
    }
    return $false
}
function Get-WdLogLines {
    if (Test-Path $wdLog) { return @([IO.File]::ReadAllLines($wdLog, [Text.Encoding]::UTF8)) }
    return @()
}
function Get-ExcelPid($app) {
    [uint32]$p = 0
    [void][AttachLinkDeadMenuPid]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$p)
    return $p
}
function New-App {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    return $app
}
function New-HiddenInstance {
    # COM 自动化拉起的隐藏实例 + 一个工作簿（零工作簿实例没有 EXCEL7 窗口，看门狗不附着）
    $app = New-App
    $wbs = $app.Workbooks
    $wb = $wbs.Add()
    Release-Com $wb
    return @{ app = $app; wbs = $wbs }
}
function Stop-InstanceByPid([uint32]$ProcId) {
    # 按进程号结束本测试拉起的 Excel。先确认该进程号对应的是一个仍在运行的 EXCEL——进程号可能来自调用方
    # 几十秒前的记录（-KnownPid），实例若已自行退出、进程号被回收分给了别的进程，盲目结束会误伤。
    # 返回 killed（已结束）/ gone（该进程号上已经没有本测试的 Excel，无须处理）
    $gp = $null
    try { $gp = [System.Diagnostics.Process]::GetProcessById([int]$ProcId) } catch { return 'gone' }
    if ($gp.ProcessName -ine 'EXCEL') { return 'gone' }
    try { Stop-Process -Id $ProcId -Force } catch {}
    [void](Wait-Exit $ProcId 10)
    return 'killed'
}
function Close-App($app, $wbs, [switch]$Kill, [uint32]$KnownPid = 0) {
    # 正常退出：标记已保存 → Quit → 显式释放全部引用；仍不退的按进程号结束（只碰本测试自己拉起的实例）。
    # -Kill：不 Quit，直接按进程号结束——用于「状态无法确认、不得把当前状态写回命令栏缓存」的场合
    #（被强杀的 Excel 不会写 Excel15.xlb）。返回 exited / killed / gone / none。调用后不得再使用传入的对象。
    # KnownPid：调用方已知的进程号。宿主繁忙时 Hwnd 可能读成空、现取不到进程号——这时用它，保证「直接结束」不落空。
    [uint32]$procId = 0
    try { $procId = Get-ExcelPid $app } catch {}
    if ($procId -eq 0) { $procId = $KnownPid }
    if (-not $Kill) {
        if ($null -ne $wbs) { try { Save-AllWorkbooks $wbs } catch {} }
        if ($null -ne $app) { try { $app.Quit() } catch {} }
    }
    Release-Com $wbs
    Release-Com $app
    if ($procId -eq 0) { return 'none' }
    if ((-not $Kill) -and (Wait-Exit $procId 20)) { return 'exited' }
    return (Stop-InstanceByPid $procId)
}
function Invoke-CacheCheck {
    # 零工作簿实例（看门狗不附着）读当前命令栏缓存，返回启动时就带的 AttachLinkCtx 一级控件个数；
    # 有则从外部删掉并正常退出（缓存被重写）。读不到 / 删不干净时无法确认缓存状态：判 FAIL，并直接结束
    # 该实例（不正常退出——否则可能把无法确认的状态写回缓存）。
    $app = $null
    $count = 0; $broken = $false
    try {
        $app = New-App
        $caps = @(Get-TaggedTop $app)
        $count = $caps.Count
        if ($count -gt 0) {
            Say ('    缓存里的残留控件: ' + ($caps -join ' | '))
            [void](Remove-TaggedControls $app)
            $leftCount = @(Get-TaggedTop $app).Count
            if ($leftCount -gt 0) { throw ('残留控件删不干净（还剩 ' + $leftCount + ' 个）') }
        }
    } catch {
        $broken = $true
        Check $false ('读取 / 清理命令栏缓存时出错：' + $_.Exception.Message + '——无法确认缓存状态，本检查实例直接结束、不正常退出')
    }
    if ($null -ne $app) {
        $how = Close-App $app $null -Kill:$broken
        if ($how -eq 'killed' -and -not $broken) { Check $false '缓存检查用的实例退出后仍在运行（已清理）' }
        if ($how -eq 'none' -and $broken) { Check $false '无法确认缓存检查用的实例已结束（进程号未知），请人工确认没有残留的隐藏 Excel' }
    }
    if ($broken) { return $null }      # 读不到 / 删不干净：没有可信的个数可报（FAIL 已在上面记录），调用方不得据此打印 pass
    return $count
}

# ---- 闸：常驻看门狗与被测脚本同版、看门狗存活、运行前没有 Excel ----
$deployedDiffers = $false
if ((-not $AssumeSameVersion) -and (Test-Path $pfWd)) { $deployedDiffers = ((Get-FileHash $pfWd).Hash -ne (Get-FileHash $wdPath).Hash) }
if ($deployedDiffers) {
    Say 'SKIP: 机器级常驻看门狗与被测脚本版本不一致——它会对同一实例做自己的动作，结果无法归因（请先重新部署）'
    Say 'RESULT: SKIP'; exit 0
}
if ($AssumeSameVersion) { Say '*** 警告：已加 -AssumeSameVersion，跳过了同版检查；常驻看门狗在运行时结果无法归因于被测脚本。' }
if (-not (Test-Path $wdLog)) { Say ('SKIP: 看门狗日志不存在（' + $wdLog + '）'); Say 'RESULT: SKIP'; exit 0 }
if (((Get-Date) - (Get-Item $wdLog).LastWriteTime).TotalSeconds -gt 90) {
    Say 'SKIP: 看门狗日志超过 90 秒未刷新，看门狗可能没有在运行'; Say 'RESULT: SKIP'; exit 0
}
$runningExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($runningExcel.Count -gt 0) {
    Say 'FAIL 前置：运行前已有 EXCEL.EXE 在运行，为避免误伤用户的实例不执行：'
    $runningExcel | ForEach-Object { Say ('    pid=' + $_.Id + '  ' + $_.MainWindowTitle) }
    $script:ok = $false
}

# ---- 启动被测看门狗（v1.6.2 起常驻看门狗不再附着隐藏实例，装载由它完成）----
# 带 -AttachHiddenInstances（回归开关，隐藏实例照旧附着——本测试的实例就是隐藏的，且靠它有一个周期的宽限
# 来抢放占位控件）；-XlamPath 与常驻看门狗同一份可信副本。启动后把 Get-WdLogLines 读取的 $wdLog 改指这份
# 独立日志：闸检查已经用主日志做完了，之后的等待逻辑原样可用、只针对被测看门狗。
$testWd = $null
$testWdLog = Join-Path $env:TEMP ('AttachLinkTestDeadMenu_wd_' + [guid]::NewGuid().ToString('N') + '.log')
try {
    # 给了 -WatchdogExe 就启动 .NET 版被测看门狗（参数名与 PowerShell 版相同）；否则启动原 PowerShell 版。
    # 日志路径、xlam 路径、-AttachHiddenInstances 与原写法逐字保持一致。
    if ($WatchdogExe) {
        $testWd = Start-Process -FilePath $WatchdogExe -ArgumentList @('-LogFile', ('"{0}"' -f $testWdLog), '-XlamPath', ('"{0}"' -f $trustedXlam), '-AttachHiddenInstances') -PassThru -WindowStyle Hidden
    } else {
        $testWd = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $wdPath), '-LogFile', ('"{0}"' -f $testWdLog), '-XlamPath', ('"{0}"' -f $trustedXlam), '-AttachHiddenInstances') -PassThru -WindowStyle Hidden
    }
} catch {
    Check $false ('被测看门狗启动失败: ' + $_.Exception.Message)
}
$wdLog = $testWdLog
$logWaitUntil = (Get-Date).AddSeconds(10)
while (-not (Test-Path $wdLog) -and (Get-Date) -lt $logWaitUntil) { Start-Sleep -Milliseconds 500 }
if (-not (Test-Path $wdLog)) { Check $false ('被测看门狗的独立日志 10 秒内未出现（' + $wdLog + '），后续等待必然失败') }

$placeholder = 'AttachLink 死菜单占位（仅测试）'
$expectedTop = @('插入附件链接（可多选）', '插入文件夹引用', '插入底稿引用', '归档迁移（批量）', '附件管理')

function Invoke-Attempt([int]$n) {
    # 结论写入 $script:verdict：PASS / RACE（看门狗抢在占位控件之前完成装载，时序未命中）/ FAIL。本函数不向管道输出任何对象
    $script:verdict = 'FAIL'
    $inst = $null; [uint32]$procId = 0; $phState = ''; $exitedEarly = $false
    $logStart = @(Get-WdLogLines).Count
    try {
        $inst = New-HiddenInstance
        $procId = Get-ExcelPid $inst.app
        if ($procId -eq 0) {
            # 取不到进程号就无法保证收尾（宿主繁忙时 Hwnd 可能读成空；进程号 0 会被 Test-Alive 当成「仍在」）——
            # 在放入占位控件【之前】中止：此时没有任何东西需要清理，finally 里的正常退出是安全的
            throw '取不到测试实例的进程号（宿主繁忙），无法保证收尾——在放入占位控件之前中止'
        }
        $bars = $null; $bar = $null; $ctrls = $null; $c = $null
        try {
            $bars = $inst.app.CommandBars
            $bar = $bars.Item('Cell')
            $ctrls = $bar.Controls
            $c = $ctrls.Add(1, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, $false)
            $phState = 'placed'              # 从这一刻起命令栏上有一个非临时控件：无论后面成败，收尾都必须处理它
            $c.Tag = 'AttachLinkCtx'        # 先设标签：即使下一步设置标题失败，收尾也能靠标签认出它
            $phState = 'tagged'
            $c.Caption = $placeholder
            $phState = 'captioned'
        } finally { Release-Com $c; Release-Com $ctrls; Release-Com $bar; Release-Com $bars }
        Say ('    第 {0} 次：实例 pid={1}（工作簿 {2} 个），已放置非临时占位控件，等待看门狗处理（最多 {3} 秒）' -f $n, $procId, $inst.wbs.Count, $WaitSeconds)

        $pidPattern = 'pid=' + $procId + '(\D|$)'
        # 注意：PowerShell 变量名不区分大小写——起名前核对没有只差大小写的同名变量（$deadline 曾与 $deadLog 的旧名 $deadLine 相撞，
        # 等待循环一次都没执行；见开发交接第九节）
        $waitUntil = (Get-Date).AddSeconds($WaitSeconds)
        $deadLog = $null; $purgeLog = $null; $initLog = $null
        while ((Get-Date) -lt $waitUntil) {
            Start-Sleep -Milliseconds 500
            if (-not (Test-Alive $procId)) { $exitedEarly = $true; break }
            $lines = @(Get-WdLogLines | Select-Object -Skip $logStart)
            $deadLog = @($lines | Where-Object { $_ -match '按未装载处理' -and $_ -match $pidPattern }) | Select-Object -First 1
            $purgeLog = @($lines | Where-Object { $_ -match '已删除死菜单控件' -and $_ -match $pidPattern }) | Select-Object -First 1
            $initLog = @($lines | Where-Object { $_ -match '已装载并初始化' -and $_ -match $pidPattern }) | Select-Object -First 1
            if ($initLog) { break }
        }
        $mine = @(Get-WdLogLines | Select-Object -Skip $logStart | Where-Object { $_ -match $pidPattern })
        foreach ($l in $mine) { Say ('    看门狗: ' + $l) }
        if ($exitedEarly) {
            Check $false ('第 {0} 次：实例 pid={1} 在等待期间意外退出（崩溃或被关闭），场景未得到验证' -f $n, $procId)
            return
        }
        if ($initLog -and -not $deadLog) {
            Say ('    第 {0} 次：时序未命中——看门狗在占位控件放好之前就完成了装载（死菜单分支没有被触发），重试' -f $n)
            $script:verdict = 'RACE'
            return
        }
        Check ([bool]$deadLog) ('第 {0} 次：看门狗识别出死菜单，按未装载处理' -f $n)
        Check ([bool]$purgeLog) ('第 {0} 次：装载之前看门狗已从外部删除死菜单控件（根除）' -f $n)
        Check ([bool]$initLog) ('第 {0} 次：随后装载并初始化' -f $n)
        if (-not ($deadLog -and $initLog)) { return }

        Start-Sleep -Milliseconds 800
        $selfTest = $null
        try { $selfTest = $inst.app.Run('AttachLink_SelfTest') } catch { $selfTest = 'ERR: ' + $_.Exception.Message }
        Check ($selfTest -eq $true) ('第 {0} 次：死菜单修复后 SelfTest 可用（实际 {1}）' -f $n, $selfTest)
        $loadedPath = Get-SessionLoadedPath $inst.wbs
        Check ($loadedPath -and ($loadedPath -ieq $trustedXlam)) ('第 {0} 次：在位副本就是配置路径（{1}）' -f $n, $loadedPath)
        $caps = @(Get-TaggedTop $inst.app)
        Say ('    一级菜单: ' + ($caps -join ' | '))
        $same = ($caps.Count -eq $expectedTop.Count)
        if ($same) { for ($i = 0; $i -lt $expectedTop.Count; $i++) { if ($caps[$i] -ne $expectedTop[$i]) { $same = $false } } }
        Check $same ('第 {0} 次：占位控件已被正式菜单替换（一级 5 项、顺序正确、无占位标题、无重复）' -f $n)
        if ($same) { $script:verdict = 'PASS' }
    } finally {
        if ($null -ne $inst) {
            $mustKill = $false
            if (-not $exitedEarly) {
                # 实例已经不在了就没有东西需要清理；否则按占位控件的跟踪状态收尾
                if ($phState -eq 'placed') {
                    # 占位控件已在命令栏上、但标签没设成功：认不出它，也就删不掉——不能正常退出
                    #（否则一个无名按钮会被写进命令栏缓存），直接结束实例
                    $mustKill = $true
                    Check $false ('第 {0} 次：占位控件的标签没有设置成功，无法识别与清理——直接结束实例（不正常退出），避免把无名按钮写进命令栏缓存' -f $n)
                } elseif ($phState -ne '') {
                    # 正常路径上占位控件已被看门狗从外部删除；时序未命中 / 看门狗没有处理时它还在——
                    # 必须先从外部删掉再正常退出，不得把它写进命令栏缓存
                    try {
                        $cleaned = Remove-PlaceholderControls $inst.app $expectedTop
                        if ($cleaned.removed -gt 0) { Say ('    （收尾）从外部删除了仍在校的占位控件 {0} 个' -f $cleaned.removed) }
                        if ($cleaned.left -gt 0) {
                            $mustKill = $true
                            Check $false ('第 {0} 次：占位控件删不干净（还剩 {1} 个）——直接结束实例（不正常退出），避免写进命令栏缓存' -f $n, $cleaned.left)
                        }
                    } catch {
                        $mustKill = $true
                        Check $false ('第 {0} 次：收尾清理占位控件时出错（{1}）——直接结束实例（不正常退出），避免写进命令栏缓存' -f $n, $_.Exception.Message)
                    }
                }
            }
            $how = Close-App $inst.app $inst.wbs -Kill:$mustKill -KnownPid $procId
            $inst = $null
            if ($how -eq 'killed') {
                if ($mustKill) { Say ('    （收尾）已直接结束测试实例 pid={0}' -f $procId) }
                else { Check $false ('第 {0} 次：测试实例 pid={1} 退出后仍在运行（已清理）' -f $n, $procId) }
            } elseif ($how -eq 'none') {
                Check $false ('第 {0} 次：无法确认测试实例已退出（进程号未知）——已发出 Quit，请人工确认没有残留的隐藏 Excel' -f $n)
            } elseif ($how -eq 'gone') {
                Say ('    （收尾）测试实例 pid={0} 在收尾前已自行退出' -f $procId)
            } elseif ($how -eq 'exited' -and $script:verdict -eq 'PASS') { Say ('    第 {0} 次：实例 pid={1} 已正常退出' -f $n, $procId) }
        }
    }
}

try {
    # ---- [0] 基线：缓存必须干净（新实例启动时不带 AttachLinkCtx 控件）----
    Say '[0] 基线检查：新实例启动时不带 AttachLinkCtx 控件（命令栏缓存干净）'
    $n0 = Invoke-CacheCheck
    if ($null -eq $n0) {
        Say '    基线：缓存状态无法确认（读取 / 清理出错，FAIL 已由缓存检查记录）'
    } elseif ($n0 -gt 0) {
        Check $false ('基线：命令栏缓存里有 {0} 个残留的 AttachLinkCtx 控件（已删除并正常退出，缓存已重写）——请查明是谁留下了非临时控件（红线见开发交接 3.3 第 9 条）' -f $n0)
    } else {
        Check $true '基线：新实例启动不带 AttachLinkCtx 控件（缓存干净）'
    }

    # ---- [1] 场景：放非临时占位控件 → 等看门狗识别、从外部删除、装载 ----
    Say '[1] 死菜单场景（最多 3 次，时序未命中则重试）'
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        Invoke-Attempt $attempt
        if ($script:verdict -ne 'RACE') { break }
        Start-Sleep -Seconds 2
    }
    if ($script:verdict -eq 'RACE') { Check $false '连续 3 次时序未命中（看门狗均抢在占位控件之前完成装载），场景未得到验证' }

    # ---- [2] 终态：缓存必须干净——这是「根除」的判别性断言（只做会话内修复时，占位控件会在这里重现）----
    Say '[2] 终态检查：实例正常退出后，新实例不再带 AttachLinkCtx 控件（死菜单已从命令栏缓存根除）'
    Start-Sleep -Seconds 1
    $n2 = Invoke-CacheCheck
    if ($null -eq $n2) {
        Say '    终态：缓存状态无法确认（读取 / 清理出错，FAIL 已由缓存检查记录）'
    } elseif ($n2 -gt 0) {
        Check $false ('终态：新实例仍带着 {0} 个 AttachLinkCtx 控件启动——看门狗只修复了会话、没有清掉缓存（已删除并正常退出，缓存已重写）' -f $n2)
    } else {
        Check $true '终态：新实例启动不带 AttachLinkCtx 控件（看门狗的删除已落到命令栏缓存，死菜单根除）'
    }
} catch {
    Check $false ('测试执行异常: ' + $_.Exception.Message)
    # 异常路径的兜底：核实并清理缓存，任何情况下都不留下占位控件
    try { $nx = Invoke-CacheCheck; if ($null -ne $nx -and $nx -gt 0) { Say ('    （异常兜底）已清理缓存里的残留控件 {0} 个' -f $nx) } } catch { Say ('    （异常兜底）缓存检查也失败了: ' + $_.Exception.Message) }
}

# ---- 收尾：结束本脚本启动的被测看门狗（按自己记下的进程号；闸之后的路径一定执行到这里）----
if ($testWd -and -not $testWd.HasExited) {
    try { Stop-Process -Id $testWd.Id -Force } catch {}
    Say ('    （收尾）已结束被测看门狗 pid=' + $testWd.Id)
}
if ($ok) { Remove-Item $testWdLog -Force -ErrorAction SilentlyContinue }   # PASS 才删日志；FAIL 时留在 %TEMP% 供排查

$leftExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($leftExcel.Count -gt 0) { Say ('留意：运行结束后仍有 EXCEL.EXE 在运行: ' + (($leftExcel | ForEach-Object { $_.Id }) -join ', ') + '（未处理。本脚本只按自己记录的进程号结束自己拉起的实例，不会结束其它 Excel；请确认这些是否为本测试遗留）') }
if ($ok) { Say 'RESULT: PASS'; exit 0 } else { Say 'RESULT: FAIL'; exit 1 }
