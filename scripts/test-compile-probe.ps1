# =====================================================================
# AttachLink 逐入口编译探针（长期回归工具）
# ---------------------------------------------------------------------
# 背景（WORKLOG 第六坑）：VBA 按需编译只编译被调过程——构建自检
# （AttachLink_SelfTest）通过不代表其他入口可编译。本脚本从
# src\AttachLink.bas 枚举全部 Public Sub/Function，逐入口经
# Application.Run 调用一次；任何编译错误（用户定义类型未定义、
# 重名、语法错误等）都会在该入口暴露。
#
# ★ 新增公开过程必须同步登记到下方 $entryTable（脚本会校验
#   「源码枚举 ↔ 登记表」一致，缺登记即 FAIL）。
#
# 运行形态：
#   - phase1：零活动工作簿（UI 包装在 ActiveWorkbook Is Nothing 处早退）
#   - phase2：带已保存夹具工作簿（Core 函数给安全参数，走早退/失败串路径）
#   - modal ：必然弹模态框的入口（ClearSelection/RenameLink）——子进程
#             运行，超时即证明「已编译、阻塞在对话框」，按 PID 清理
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File test-compile-probe.ps1
# 期望末行：RESULT: PASS
# 注意：含中文，必须 UTF-8 with BOM。
# =====================================================================
param(
    [string]$Child = '',          # 内部：子进程模式（探 modal 入口）
    [string]$ChildDir = ''
)
$ErrorActionPreference = 'Stop'
$attach = [string]([char]0x9644 + [char]0x4EF6)   # 附件
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
$basSrc = Join-Path $PSScriptRoot '..\src\AttachLink.bas'
if (-not (Test-Path $xlam)) { throw "未找到 $xlam，请先运行 build\build.ps1。" }

Add-Type -Namespace WP -Name U32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
function Get-XlPid($xl) {
    $procId = 0
    [void][WP.U32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$procId)
    return $procId
}

# ---------------------------------------------------------------------
# 子进程模式：探测必然弹模态框的入口。能跑到对话框 = 编译通过。
# ---------------------------------------------------------------------
if ($Child -ne '') {
    New-Item -ItemType Directory -Force -Path $ChildDir | Out-Null
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    (Get-XlPid $xl) | Out-File (Join-Path $ChildDir 'child.pid') -Encoding ascii
    $null = $xl.Workbooks.Open($xlam)          # 第一时间装载 dist 副本（防常驻看门狗抢装）
    $attDir = Join-Path $ChildDir $attach
    New-Item -ItemType Directory -Force -Path $attDir | Out-Null
    Set-Content (Join-Path $attDir 'a.pdf') -Value 'probe-content'
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(1, 1), ($attach + '\a.pdf'), '', 'PROBE')
    $wb.SaveAs((Join-Path $ChildDir 'child.xlsx'), 51)
    $ws.Range('A1').Select()
    # ready 标志：子进程已抵达入口调用点。父进程超时判定必须以此为据——
    # 仅当标志已写且子进程仍存活（阻塞在模态框）时才判「已编译」；
    # 无声卡死/启动即崩溃不得被误判为通过
    Set-Content (Join-Path $ChildDir 'child.ready') -Value '1'
    switch ($Child) {
        'AttachLink_ClearSelection' { $xl.Run('AttachLink_ClearSelection') }
        'AttachLink_RenameLink'     { $xl.Run('AttachLink_RenameLink') }
        default                     { $xl.Run($Child) }
    }
    'CHILD-COMPLETED'                          # 未阻塞（入口自己早退）也算编译通过
    try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
    exit 0
}

# ---------------------------------------------------------------------
# 入口登记表：名称 → 运行形态与参数（新增公开过程必须同步登记）
# 参数令牌：WB=夹具工作簿  RANGE=夹具选区  SRCFILE=源文件  SRCXLSX=源底稿
#           NODIR=不存在目录  OUTSIDE=附件夹外路径  其余为字面字符串
# ---------------------------------------------------------------------
$entryTable = [ordered]@{
    'Auto_Open'                       = @{ mode = 'phase1' }
    'AttachLink_Initialize'           = @{ mode = 'phase1' }
    'AttachLink_EnsureMenu'           = @{ mode = 'phase1' }
    'AttachLink_Insert'               = @{ mode = 'phase1' }
    'AttachLink_BuildIndex'           = @{ mode = 'phase1' }
    'AttachLink_Repair'               = @{ mode = 'phase1' }
    'AttachLink_InsertFolder'         = @{ mode = 'phase1' }
    'AttachLink_DeleteAttachment'     = @{ mode = 'phase1' }
    'AttachLink_InsertCrossRef'       = @{ mode = 'phase1' }
    'AttachLink_ArchiveMigrate'       = @{ mode = 'phase1' }
    'AttachLink_UpdateVersion'        = @{ mode = 'phase1' }
    'AttachLink_SelfTest'             = @{ mode = 'phase1' }
    'AttachLink_GetVersion'           = @{ mode = 'phase1' }
    'AttachLink_ClearSelection'       = @{ mode = 'modal' }
    'AttachLink_RenameLink'           = @{ mode = 'modal' }
    'AttachLink_BuildIndexCore'       = @{ mode = 'phase2'; args = @('WB') }
    'NormalizeAttachmentLinks'        = @{ mode = 'phase2'; args = @('WB') }
    'AttachLink_InsertCore'           = @{ mode = 'phase2'; args = @('SRCFILE', 'RANGE') }
    'AttachLink_RepairCore'           = @{ mode = 'phase2'; args = @('WB', 'NODIR') }
    'AttachLink_InsertFolderCore'     = @{ mode = 'phase2'; args = @('NODIR') }
    'AttachLink_InsertCrossRefCore'   = @{ mode = 'phase2'; args = @('SRCXLSX', 'STR:Sheet1', 'RANGE2') }
    'AttachLink_DeleteAttachmentCore' = @{ mode = 'phase2'; args = @('WB', 'STR:A1|Sheet1', 'OUTSIDE') }
    'AttachLink_ArchiveMigrateCore'   = @{ mode = 'phase2'; args = @('NODIR') }
    'AttachLink_RenameDiag'           = @{ mode = 'phase2'; args = @('STR:probe') }
    'AttachLink_GetHashCount'         = @{ mode = 'phase2' }
    'AttachLink_ResetHashCount'       = @{ mode = 'phase2' }
    'AttachLink_GetProgressCount'     = @{ mode = 'phase2' }
    'AttachLink_ResetProgressCount'   = @{ mode = 'phase2' }
    'AttachLink_NotifyNormalizeFailure' = @{ mode = 'phase2' }   # 无参形态；带 WB 形态见 phase2b
    'AttachLink_TestRepointFail'      = @{ mode = 'phase2'; args = @('NUM:0') }   # 测试钩子归零（生产语义）
    'AttachLink_TestCancelInFolder'   = @{ mode = 'phase2'; args = @('NUM:0') }   # O-06 嵌套取消钩子归零
    'AttachLink_TestReentryProbe'     = @{ mode = 'phase2' }   # O-06 重入演练置位（生产不调用）
    'AttachLink_TestReentrySeen'      = @{ mode = 'phase2' }   # O-06 重入拦截结果读取
    'AttachLink_TestThrowAtInsert'    = @{ mode = 'phase2'; args = @('NUM:0') }   # V2 包装器兜底注入归零
    'AttachLink_TestGrowAfterCopy'    = @{ mode = 'phase2'; args = @('NUM:0') }   # V11b 复制后源增长注入归零
    'AttachLink_TestWebPathGate'      = @{ mode = 'phase2'; args = @('STR:https://example.test/wb.xlsx') }   # V10 网页闸
    'AttachLink_UpdateVersionCore'    = @{ mode = 'phase2'; args = @('ATTFILE', 'SRCFILE') }   # N-07：未登记基线 → 安全早退
    'Auto_Close'                      = @{ mode = 'phase2' }   # 最后：注销菜单与快捷键
}

# ---- 源码枚举 ↔ 登记表一致性 ----
$src = [IO.File]::ReadAllText($basSrc, [Text.Encoding]::UTF8)
$enumerated = @{}
foreach ($m in [regex]::Matches($src, '(?m)^\s*Public\s+(?:Sub|Function)\s+([A-Za-z0-9_]+)')) {
    $enumerated[$m.Groups[1].Value] = $true
}
$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }
foreach ($name in $enumerated.Keys) {
    if (-not $entryTable.Contains($name)) { Check $false "公开过程 $name 未登记到探针表（新增公开过程必须同步登记）" }
}
foreach ($name in $entryTable.Keys) {
    if (-not $enumerated.Contains($name)) { Check $false "探针表登记的 $name 在源码中不存在（请同步清理登记表）" }
}

# ---- 夹具 ----
$base = Join-Path $env:TEMP 'opencode\compileprobe'
Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
$wbDir = "$base\proj"
$attDir = "$wbDir\$attach"
New-Item -ItemType Directory -Force -Path $attDir, "$base\child" | Out-Null
Set-Content "$base\srcfile.pdf" -Value 'probe-src'
Set-Content "$attDir\probe-outside-target.pdf" -Value 'x'   # 供 OUTSIDE 令牌指向附件夹外（实际用 $base 下文件）
Set-Content "$base\outside.pdf" -Value 'outside'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xlPid = Get-XlPid $xl
"Excel PID: $xlPid"
$null = $xl.Workbooks.Open($xlam)              # 第一时间装载 dist 副本

# 超时兜底：10 分钟后脚本仍未结束（COM 卡死）→ 按 PID 结束本实例
$sentinel = "$base\done.flag"
$killer = Start-Process powershell -PassThru -WindowStyle Hidden -ArgumentList '-NoProfile', '-Command', `
    "Start-Sleep -Seconds 600; if (-not (Test-Path '$sentinel')) { Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue }"

function Resolve-Args($spec, $wb, $ws) {
    $out = @()
    foreach ($a in $spec) {
        switch -regex ($a) {
            '^WB$'      { $out += $wb }
            '^RANGE$'   { $out += $ws.Range('J1') }
            '^RANGE2$'  { $out += $ws.Range('J2') }
            '^SRCFILE$' { $out += "$base\srcfile.pdf" }
            '^SRCXLSX$' { $out += "$base\srcwb.xlsx" }
            '^NODIR$'   { $out += "$base\no-such-dir" }
            '^OUTSIDE$' { $out += "$base\outside.pdf" }
            '^ATTFILE$' { $out += "$attDir\probe-outside-target.pdf" }
            '^STR:(.*)$' { $out += $Matches[1] }
            '^NUM:(\d+)$' { $out += [long]$Matches[1] }
            default     { $out += $a }
        }
    }
    return $out
}

try {
    '--- phase1：零活动工作簿（UI 包装早退路径） ---'
    foreach ($name in $entryTable.Keys) {
        $e = $entryTable[$name]
        if ($e.mode -ne 'phase1') { continue }
        try {
            $null = $xl.Run($name)
            "OK   $name"
        } catch {
            Check $false "$name 调用失败：$($_.Exception.Message)"
        }
    }

    '--- phase2：夹具工作簿（Core 函数安全参数路径） ---'
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $wb.SaveAs("$wbDir\main.xlsx", 51)
    # 源底稿（InsertCrossRefCore 用）
    $swb = $xl.Workbooks.Add()
    $swb.SaveAs("$base\srcwb.xlsx", 51)
    $swb.Close($false)
    $wb.Activate()

    foreach ($name in $entryTable.Keys) {
        $e = $entryTable[$name]
        if ($e.mode -ne 'phase2') { continue }
        try {
            if ($e.Contains('args')) {
                # COM 方法参数须逐个展开（数组整体传入会报 DISP_E_PARAMNOTOPTIONAL）
                $argArr = @(Resolve-Args $e.args $wb $ws)
                switch ($argArr.Count) {
                    1 { $null = $xl.Run($name, $argArr[0]) }
                    2 { $null = $xl.Run($name, $argArr[0], $argArr[1]) }
                    3 { $null = $xl.Run($name, $argArr[0], $argArr[1], $argArr[2]) }
                    default { throw "探针表参数个数未支持：$($argArr.Count)" }
                }
            } else {
                $null = $xl.Run($name)
            }
            "OK   $name"
        } catch {
            Check $false "$name 调用失败：$($_.Exception.Message)"
        }
    }

    '--- phase2b：同一过程的其它调用形态（T14/T8 新增签名） ---'
    try {
        # BuildIndexCore 带 cancelAfter：夹具仅 1 个链接，钩子（=2）不触发 → 正常完成
        $null = $xl.Run('AttachLink_BuildIndexCore', $wb, 2)
        "OK   AttachLink_BuildIndexCore(WB, cancelAfter=2)"
    } catch {
        Check $false "AttachLink_BuildIndexCore(WB, cancelAfter) 调用失败：$($_.Exception.Message)"
    }
    try {
        $null = $xl.Run('AttachLink_NotifyNormalizeFailure', $wb)
        "OK   AttachLink_NotifyNormalizeFailure(WB)"
    } catch {
        Check $false "AttachLink_NotifyNormalizeFailure(WB) 调用失败：$($_.Exception.Message)"
    }

    '--- modal：必然弹模态框的入口（子进程 + 超时判定已编译） ---'
    foreach ($name in $entryTable.Keys) {
        $e = $entryTable[$name]
        if ($e.mode -ne 'modal') { continue }
        $cDir = "$base\child\$name"
        $outF = "$cDir.out"; $errF = "$cDir.err"
        $proc = Start-Process powershell -PassThru -WindowStyle Hidden `
            -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"", '-Child', $name, '-ChildDir', "`"$cDir`"" `
            -RedirectStandardOutput $outF -RedirectStandardError $errF
        $exited = $proc.WaitForExit(30000)
        if ($exited) {
            $outText = [string](Get-Content $outF -Raw -ErrorAction SilentlyContinue)
            if ($outText -match 'CHILD-COMPLETED') {
                "OK   $name（子进程完成）"
            } else {
                $errText = [string](Get-Content $errF -Raw -ErrorAction SilentlyContinue)
                Check $false "$name 子进程异常：$errText"
            }
        } else {
            # 超时只有在「子进程已就绪（ready 标志已写）且仍存活（阻塞在模态框）」时才判
            # 「已编译」；否则视为卡死/崩溃，不得把无声卡死当作编译通过
            $ready = Test-Path "$cDir\child.ready"
            $alive = -not $proc.HasExited
            $cpid = [int](Get-Content "$cDir\child.pid" -ErrorAction SilentlyContinue)
            if ($cpid -gt 0) { Stop-Process -Id $cpid -Force -ErrorAction SilentlyContinue }
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            if ($ready -and $alive) {
                "OK   $name（已编译，阻塞在模态框，已按 PID $cpid 清理）"
            } else {
                Check $false "$name 子进程超时但未见就绪标志（ready=$ready alive=$alive，疑似卡死而非模态框）"
            }
        }
    }
}
finally {
    New-Item -ItemType File -Force -Path $sentinel | Out-Null
    try { Stop-Process -Id $killer.Id -Force -ErrorAction SilentlyContinue } catch {}
    try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
    Start-Sleep 2
}

if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
