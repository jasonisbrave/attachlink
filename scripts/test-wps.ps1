# =====================================================================
# WPS（ket.Application）兼容性验证：晚绑定改造后的核心链路
# 注意：只 Quit 自己拉起的 COM 实例，不动用户已开的 WPS
# =====================================================================
$ErrorActionPreference = 'Continue'
$attach = [string]([char]0x9644 + [char]0x4EF6)
$idxName = $attach + [string]([char]0x6E05) + [string]([char]0x5355)
$base = "$env:TEMP\opencode\wpstest"
$wbDir = "$base\proj"
$attDir = "$wbDir\$attach"
$srcDir = "$base\src\证据包"
Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path "$srcDir\子目录", $wbDir, "$wbDir\$attach" | Out-Null
Set-Content "$srcDir\合同.pdf" -Value 'wps-contract'
Set-Content "$srcDir\子目录\发票.pdf" -Value 'wps-invoice'

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }

$wps = New-Object -ComObject ket.Application
$wps.Visible = $false
try { $wps.DisplayAlerts = $false } catch {}

$xlamPath = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'

'--- 1) 打开 xlam + SelfTest ---'
try {
    $null = $wps.Workbooks.Open($xlamPath)
    # v1.6.4：加载项不再带 Workbook_Open，打开不隐式初始化；菜单、快捷键、应用事件都由显式的 AttachLink_Initialize 建立
    $menuOnOpen = @(@($wps.CommandBars.Item('Cell').Controls) | Where-Object { [string]$_.Caption -eq '插入附件链接（可多选）' }).Count
    Check ($menuOnOpen -eq 0) "打开 xlam 不隐式初始化（Cell 菜单上 AttachLink 控件数 $menuOnOpen）"
    $null = $wps.Run('AttachLink_Initialize')
    $r = $wps.Run('AttachLink_SelfTest')
    Check ($r -eq $true) 'SelfTest（晚绑定改造后可编译）'
    $v = $wps.Run('AttachLink_GetVersion')
    Check ("$v" -eq '1.6.4') "GetVersion=1.6.4（实际 $v）"
} catch { Check $false "打开/SelfTest 失败: $($_.Exception.Message.Split("`n")[0])" }

'--- 2) 清单/体检（含标记回退） ---'
try {
    $wb = $wps.Workbooks.Add()
    Set-Content "$attDir\a.pdf" -Value 'wps-a'
    $ws = $wb.Worksheets.Item(1)
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(1,1), ($attach + '\a.pdf'), '', 'A')
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(2,1), ($attach + '\missing.pdf'), '', 'B')
    $wb.SaveAs("$wbDir\main.xlsx", 51)
    $summary = $wps.Run('AttachLink_BuildIndexCore', $wb)
    $sheet = $null
    try { $sheet = $wb.Worksheets.Item($idxName) } catch {}
    Check ($null -ne $sheet) '清单表生成（归属标记兼容）'
    if ($sheet) {
        $row = 1; $vals = @{}
        while ($row -le 12) {
            $c3 = $sheet.Cells.Item($row, 3).Value2
            if ($c3 -and "$c3" -like '*.pdf') { $vals["$c3"] = @{ E = [string]$sheet.Cells.Item($row, 5).Value2; F = [string]$sheet.Cells.Item($row, 6).Value2 } }
            $row++
        }
        Check ($vals["$attach\a.pdf"].E -eq '正常')          '存在性检查正常'
        Check ($vals["$attach\a.pdf"].F -like '*v1*')         "基线登记（实际 [$($vals["$attach\a.pdf"].F)]）"
        Check ($vals["$attach\missing.pdf"].E -eq '断链')     '断链检查'
    }
    "--- 汇总: $($summary -replace "`r`n", ' | ')"
    '--- 2b) 二次生成清单（验证标记回退后允许重建，同名表不被误拒） ---'
    try {
        $r2b = $wps.Run('AttachLink_BuildIndexCore', $wb)
        Check ($r2b -and ($r2b -notlike '已存在同名*')) '二次生成未被拒（名称标记回退有效）'
        $sheet2b = $null
        try { $sheet2b = $wb.Worksheets.Item($idxName) } catch {}
        Check ($null -ne $sheet2b) '清单表仍在（重建成功）'
    } catch { Check $false "二次清单失败: $($_.Exception.Message.Split("`n")[0])" }
} catch { Check $false "清单失败: $($_.Exception.Message.Split("`n")[0])" }

'--- 3) 插入文件夹核心 ---'
try {
    $r3 = $wps.Run('AttachLink_InsertFolderCore', $srcDir)
    "    raw 返回: [$r3]"
    Check ($r3 -and $r3 -notlike '失败：*') 'InsertFolderCore 执行'
    Check (Test-Path "$attDir\证据包\子目录\发票.pdf") '文件夹复制+子目录'
    Check (Test-Path "$attDir\_attachlink_baseline.csv") '基线文件生成'
    if (Test-Path "$attDir\_attachlink_baseline.csv") {
        Check ((Get-Content "$attDir\_attachlink_baseline.csv" -Encoding UTF8) -match '子目录') '基线含子目录文件'
    }
    "--- 汇总: $($r3 -replace "`r`n", ' | ')"
} catch { Check $false "文件夹失败: $($_.Exception.Message.Split("`n")[0])" }

'--- 4) 右键菜单（CommandBars 兼容性观察） ---'
try {
    $ctrls = @($wps.CommandBars.Item('Cell').Controls)
    $names = $ctrls | ForEach-Object { $_.Caption }
    "    顶层前5: $(($names | Select-Object -First 5) -join ' | ')"
    Check (@($ctrls | Where-Object { "$($_.Caption)" -eq '插入附件链接（可多选）' }).Count -eq 1) '菜单注册可见'
} catch { Check $false "CommandBars 不可用: $($_.Exception.Message.Split("`n")[0])" }

'--- 5) 哈希引擎（R4：CAPI 主路径）与性能守卫 ---'
# 安全规则：只动自己拉起的 et 实例。先经 Hwnd 取 PID；单次调用挂住超 90s 只结束自己的实例。
$script:R4MyPid = 0
try {
    Add-Type -Namespace WPSR4 -Name U32 -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);' -ErrorAction SilentlyContinue
    $hwnd = 0
    try { $hwnd = [int64]$wps.Hwnd } catch { $hwnd = 0 }
    if ($hwnd -ne 0) {
        $pidOut = [uint32]0
        [void][WPSR4.U32]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$pidOut)
        if ($pidOut -gt 0) { $script:R4MyPid = [int]$pidOut }
    }
} catch {}
"    本次拉起的 et PID: $script:R4MyPid（只 Quit/结束这一个实例；用户 wps.exe 不动）"

function Invoke-R4Timed($sb, [int]$timeoutMs) {
    $wd = $null
    if ($script:R4MyPid -gt 0) {
        $cmd = "Start-Sleep -Milliseconds $timeoutMs; Stop-Process -Id $($script:R4MyPid) -Force -ErrorAction SilentlyContinue"
        $wd = Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile','-WindowStyle','Hidden','-Command', $cmd -PassThru -WindowStyle Hidden
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = @{ Ms = -1; Result = $null; Error = ''; Hung = $false }
    try { $r.Result = & $sb } catch { $r.Error = $_.Exception.Message }
    $sw.Stop()
    if ($null -ne $wd) {
        $fired = $wd.HasExited
        try { if (-not $wd.HasExited) { $wd.Kill() } } catch {}
        if ($fired) { $r.Hung = $true }
    }
    $r.Ms = [math]::Round($sw.Elapsed.TotalMilliseconds, 1)
    return $r
}

$engDir = "$base\eng"
New-Item -ItemType Directory -Force -Path $engDir | Out-Null
$rnd = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$b1 = New-Object byte[] 1MB; $rnd.GetBytes($b1); [IO.File]::WriteAllBytes("$engDir\rand1mb.bin", $b1)
$b10 = New-Object byte[] 10MB; $rnd.GetBytes($b10); [IO.File]::WriteAllBytes("$engDir\rand10mb.bin", $b10)
$b5 = New-Object byte[] 5MB; $rnd.GetBytes($b5); [IO.File]::WriteAllBytes("$engDir\rand5mb.bin", $b5)
$rnd.Dispose()
[IO.File]::WriteAllText("$engDir\abc.bin", 'abc')

try {
    $st = [string]$wps.Run('AttachLink_HashEngineState')
    "    HashEngineState: $st"
    Check ($st -match 'capi=1') "WPS 下 capi=1（实际 $st）"
    Check ($st -match 'win64=0') "WPS 下 win64=0（实际 $st）"

    foreach ($sz in @('rand1mb', 'rand10mb')) {
        $fp = "$engDir\$sz.bin"
        $expected = ([string](Get-FileHash $fp -Algorithm SHA256).Hash).ToLower()
        $h1 = [string]$wps.Run('AttachLink_HashEngine', $fp, 1)
        Check ($h1 -eq $expected) "engine1(CAPI) $sz 与 Get-FileHash 一致"
    }

    $hAbcVba = [string]$wps.Run('AttachLink_HashEngine', "$engDir\abc.bin", 3)
    Check ($hAbcVba -eq 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad') 'engine3(纯VBA回退) abc 标准向量'

    $hCert = [string]$wps.Run('AttachLink_HashEngine', "$engDir\abc.bin", 2)
    "    engine2(certutil) 平台限制只打印不断言: [$hCert]"

    # 回退链（生产分派）：把本会话的 CAPI 标成「自检失败」后，Sha256File 在 WPS 下须改走纯 VBA 并得到正确哈希
    $bk = New-Object byte[] 1024; (New-Object System.Random 7).NextBytes($bk); [IO.File]::WriteAllBytes("$engDir\k1.bin", $bk)
    $null = $wps.Run('AttachLink_TestCapiState', -1)
    Check ([string]$wps.Run('AttachLink_HashEngineState') -match 'capi=-1') 'WPS 下钩子生效：capi=-1'
    $null = $wps.Run('AttachLink_TestCapiState', 1)
    Check ([string]$wps.Run('AttachLink_HashEngineState') -match 'capi=-1') 'WPS 下钩子拒绝把状态直接置成 1（状态仍为 -1）'
    foreach ($fp in @("$engDir\abc.bin", "$engDir\k1.bin")) {
        $expected = ([string](Get-FileHash $fp -Algorithm SHA256).Hash).ToLower()
        $hf = [string]$wps.Run('AttachLink_HashEngine', $fp, 0)
        Check ($hf -eq $expected) ("回退链（生产分派，CAPI 不可用→纯 VBA）{0} 与 Get-FileHash 一致" -f (Split-Path $fp -Leaf))
    }
    Check ([string]$wps.Run('AttachLink_HashEngine', "$engDir\abc.bin", 1) -eq '') 'WPS 下 CAPI 标成不可用后，直调 CAPI 返回空串（本会话不重试）'
    $null = $wps.Run('AttachLink_TestCapiState', 0)
    Check ([string]$wps.Run('AttachLink_HashEngine', "$engDir\abc.bin", 0) -eq 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad') 'WPS 下还原后重新自检，CAPI 恢复为主路径'

    # 性能守卫：5MB 新文件插入（新建底稿目录 + 全新附件夹）须 < 8s 且非失败（1.5.1 实测 23.3s）
    $perfDir = "$base\perf\proj"
    New-Item -ItemType Directory -Force -Path $perfDir | Out-Null
    $pwb = $wps.Workbooks.Add()
    $pwb.SaveAs("$perfDir\perf.xlsx", 51)   # 勿用 main.xlsx：第 2 节已打开同名工作簿，WPS 会拒存（0x8FE3001F）
    try { $pwb.Activate() } catch {}
    $pcell = $pwb.Worksheets.Item(1).Cells.Item(1, 1)
    $pr = Invoke-R4Timed { $wps.Run('AttachLink_InsertCore', "$engDir\rand5mb.bin", $pcell) } 90000
    "    InsertCore 5MB 耗时 $($pr.Ms) ms（上限 8000 ms）hung=$($pr.Hung) err=$($pr.Error)"
    Check (-not $pr.Hung) 'InsertCore 5MB 未挂起'
    Check ($pr.Ms -ge 0 -and $pr.Ms -lt 8000) "InsertCore 5MB 性能守卫 <8s（实际 $($pr.Ms) ms）"
    $sumLine = "$($pr.Result)".Split("`n")[0]
    Check ($sumLine -notlike '失败：*') "InsertCore 5MB 汇总非失败（实际 [$sumLine]）"
    Check (Test-Path "$perfDir\$attach\rand5mb.bin") '5MB 附件已复制入全新附件夹'
} catch { Check $false "哈希引擎段失败: $($_.Exception.Message.Split("`n")[0])" }

try { foreach ($w in @($wps.Workbooks)) { $w.Saved = $true }; $wps.Quit() } catch {}
Start-Sleep 3
if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
