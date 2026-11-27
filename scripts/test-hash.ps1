# =====================================================================
# AttachLink 附件清单内容基线（SHA-256）功能测试
# ---------------------------------------------------------------------
# 生命周期：登记基线 -> 篡改附件 -> 体检红字警示且基线不被洗白
#           -> 删除基线文件重建 -> 恢复一致。另测断链行不受影响。
# 用法：powershell -ExecutionPolicy Bypass -File test-hash.ps1
# 期望末行：RESULT: PASS（夹具自动清理；FAIL 保留现场）
# 注意：含中文，必须 UTF-8 with BOM。
# =====================================================================
$ErrorActionPreference = 'Stop'
$attach = [string]([char]0x9644 + [char]0x4EF6)   # 附件
$idxName = $attach + [string]([char]0x6E05) + [string]([char]0x5355)   # 附件清单 = 附件 + 清 + 单
$base = Join-Path $env:TEMP 'opencode\hashtest'
$wbDir = "$base\proj"                             # 底稿所在目录
$attDir = "$wbDir\$attach"
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
$baseline = "$attDir\_attachlink_baseline.csv"

Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $attDir | Out-Null
Set-Content "$attDir\a.pdf" -Value 'original-content-A'
Set-Content "$attDir\b.pdf" -Value 'original-content-B'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$null = $xl.Workbooks.Open($xlam)
# v1.6.4：加载项不再带 Workbook_Open，打开不隐式初始化——菜单、快捷键、应用事件都由显式初始化建立
$null = $xl.Run('AttachLink_Initialize')

# 底稿：相对链接 a.pdf / b.pdf / broken.pdf（断链行）
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$names = @('a', 'b', 'broken')
foreach ($i in 0..2) {
    $addr = $attach + '\' + $names[$i] + '.pdf'
    $null = $ws.Hyperlinks.Add($ws.Cells.Item($i + 1, 1), $addr, '', ('L' + ($i + 1)))
}
$wb.SaveAs("$wbDir\main.xlsx", 51)

function ColF($xl, $path) {
    # 打开底稿跑清单核心，读取 F 列（内容基线）状态与汇总文本
    $w = $xl.Workbooks.Open($path)
    $summary = $xl.Run('AttachLink_BuildIndexCore', $w)
    $sheet = $w.Worksheets.Item($idxName)
    $vals = @{}
    $row = 1
    while ($row -le 40) {
        $c = $sheet.Cells.Item($row, 3).Value2      # 相对路径列定位行
        if ($c -and ($c -like '*.pdf')) { $vals["$c"] = [string]$sheet.Cells.Item($row, 6).Value2 }
        $row++
    }
    $w.Save()
    $w.Close($false)
    @{ vals = $vals; summary = $summary }
}

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }

'--- 1) 首次体检：登记基线 ---'
$r1 = ColF $xl "$wbDir\main.xlsx"
$r1.summary
Check (Test-Path $baseline)                                '基线文件已生成'
Check ($r1.vals["$attach\a.pdf"] -eq '已登记 v1')           'a.pdf 登记基线'
Check ($r1.vals["$attach\b.pdf"] -eq '已登记 v1')          'b.pdf 登记基线'
Check ($r1.vals["$attach\broken.pdf"] -eq '—')             '断链行不参与基线（—）'
$h1 = (Get-Content $baseline) -join '|'
Check ($h1 -match 'a\.pdf\|[0-9a-f]{64}')                  '基线含 a.pdf 的64位哈希'

'--- 2) 模拟同步端同名静默覆盖（篡改 a.pdf）---'
Set-Content "$attDir\a.pdf" -Value 'CLOBBERED-by-sync'
$r2 = ColF $xl "$wbDir\main.xlsx"
Check ($r2.vals["$attach\a.pdf"] -eq '■ 内容变更')          '篡改被红字警示'
Check ($r2.vals["$attach\b.pdf"] -eq '一致 v1')                '未动文件仍一致'
$h2 = (Get-Content $baseline) -join '|'
Check ($h1 -eq $h2)                                        '基线未被篡改洗白（哈希保持登记时值）'
Check ($r2.summary -match '内容变更 1')                     '汇总计入内容变更 1'

'--- 3) 有意替换：删基线重建 ---'
Remove-Item $baseline -Force
$r3 = ColF $xl "$wbDir\main.xlsx"
Check ($r3.vals["$attach\a.pdf"] -eq '已登记 v1')          '重建基线后重新登记'

'--- 4) 二次体检全部一致 ---'
$r4 = ColF $xl "$wbDir\main.xlsx"
Check ($r4.vals["$attach\a.pdf"] -eq '一致 v1')                'a.pdf 一致'
Check ($r4.vals["$attach\b.pdf"] -eq '一致 v1')                'b.pdf 一致'

'--- 5) 哈希引擎矩阵（R4：CAPI 主路径 / certutil / 纯 VBA 回退 / 生产分派） ---'
$engDir = "$base\eng"
$engCnDir = "$engDir\哈希夹具目录"
New-Item -ItemType Directory -Force -Path $engCnDir | Out-Null
function Write-RandBytes([int]$n, [string]$path, [int]$seed) {
    $bytes = New-Object byte[] $n
    $rnd = New-Object System.Random $seed
    $rnd.NextBytes($bytes)
    [System.IO.File]::WriteAllBytes($path, $bytes)
}
$fixtures = @()
# 0 字节 / 1 字节 / abc(3) / 63 / 64 / 65 字节
New-Item -ItemType File -Force -Path "$engDir\empty.bin" | Out-Null
[IO.File]::WriteAllBytes("$engDir\one.bin", [byte[]](0x41))
[IO.File]::WriteAllText("$engDir\abc.bin", 'abc')
Write-RandBytes 63  "$engDir\b63.bin"  63
Write-RandBytes 64  "$engDir\b64.bin"  64
Write-RandBytes 65  "$engDir\b65.bin"  65
# 跨 4MB 分块边界：4MB-1 / 4MB / 4MB+1
Write-RandBytes 4194303 "$engDir\m4mb_minus1.bin" 4194303
Write-RandBytes 4194304 "$engDir\m4mb.bin"        4194304
Write-RandBytes 4194305 "$engDir\m4mb_plus1.bin"  4194305
# 约 10MB 随机内容
$bigBytes = New-Object byte[] (10MB)
$rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$rng.GetBytes($bigBytes)
[IO.File]::WriteAllBytes("$engDir\rand10mb.bin", $bigBytes)
$rng.Dispose()
# 中文文件名 + 中文子目录
[IO.File]::WriteAllText("$engCnDir\附件甲·乙#.pdf", 'chinese-name-content')
foreach ($f in (Get-ChildItem $engDir -File)) { $fixtures += $f.FullName }
foreach ($f in (Get-ChildItem $engCnDir -File)) { $fixtures += $f.FullName }

$smallForVba = @{}   # engine 3（纯 VBA）仅对小夹具断言（慢）
foreach ($p in @("$engDir\empty.bin", "$engDir\one.bin", "$engDir\abc.bin", "$engDir\b63.bin", "$engDir\b64.bin", "$engDir\b65.bin")) { $smallForVba[$p.ToLower()] = $true }

foreach ($fp in $fixtures) {
    $name = Split-Path $fp -Leaf
    $expected = ([string](Get-FileHash $fp -Algorithm SHA256).Hash).ToLower()
    $h1 = [string]$xl.Run('AttachLink_HashEngine', $fp, 1)
    Check ($h1 -eq $expected) "engine1(CAPI) $name 与 Get-FileHash 一致"
    $h2 = [string]$xl.Run('AttachLink_HashEngine', $fp, 2)
    Check ($h2 -eq $expected) "engine2(certutil) $name 与 Get-FileHash 一致"
    $h0 = [string]$xl.Run('AttachLink_HashEngine', $fp, 0)
    Check ($h0 -eq $expected) "engine0(生产分派) $name 与 Get-FileHash 一致"
    if ($smallForVba.ContainsKey($fp.ToLower())) {
        $h3 = [string]$xl.Run('AttachLink_HashEngine', $fp, 3)
        Check ($h3 -eq $expected) "engine3(纯VBA) $name 与 Get-FileHash 一致"
    }
}

'--- 6) 标准向量 ---'
$hEmpty = [string]$xl.Run('AttachLink_HashEngine', "$engDir\empty.bin", 1)
Check ($hEmpty -eq 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855') '空文件 SHA-256 标准向量（CAPI）'
$hAbc = [string]$xl.Run('AttachLink_HashEngine', "$engDir\abc.bin", 1)
Check ($hAbc -eq 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad') 'abc 文件 SHA-256 标准向量（CAPI）'

'--- 7) 引擎状态与计数语义 ---'
$state = [string]$xl.Run('AttachLink_HashEngineState')
"    HashEngineState: $state"
Check ($state -match 'capi=1') "HashEngineState 报 capi=1（实际 $state）"
Check ($state -match 'win64=1') "HashEngineState 报 win64=1（实际 $state）"
$null = $xl.Run('AttachLink_ResetHashCount')
[void][string]$xl.Run('AttachLink_HashEngine', "$engDir\abc.bin", 0)
$hc1 = [long]$xl.Run('AttachLink_GetHashCount')
Check ($hc1 -eq 1) "engine0 计 1 次真实计算（实际 $hc1）"
[void][string]$xl.Run('AttachLink_HashEngine', "$engDir\abc.bin", 1)
$hc2 = [long]$xl.Run('AttachLink_GetHashCount')
Check ($hc2 -eq 1) "engine1 直调不计数（实际 $hc2）"

'--- 8) CAPI 不可用时，生产分派经回退链仍得到正确哈希（测试钩子把本会话的 CAPI 标成「自检失败」） ---'
$prevState = [long]$xl.Run('AttachLink_TestCapiState', -1)
Check ([string]$xl.Run('AttachLink_HashEngineState') -match 'capi=-1') '钩子生效：HashEngineState 报 capi=-1'
$null = $xl.Run('AttachLink_TestCapiState', 1)
Check ([string]$xl.Run('AttachLink_HashEngineState') -match 'capi=-1') '钩子拒绝把状态直接置成 1（那会绕过自检；状态仍为 -1）'
$null = $xl.Run('AttachLink_ResetHashCount')
$fallbackFixtures = @("$engDir\empty.bin", "$engDir\one.bin", "$engDir\abc.bin", "$engDir\b64.bin", "$engDir\b65.bin", "$engDir\m4mb_minus1.bin", "$engDir\m4mb.bin", "$engDir\m4mb_plus1.bin", "$engDir\rand10mb.bin")
foreach ($f in (Get-ChildItem $engCnDir -File)) { $fallbackFixtures += $f.FullName }
$fallbackChecked = 0
foreach ($fp in $fallbackFixtures) {
    $expected = ([string](Get-FileHash $fp -Algorithm SHA256).Hash).ToLower()
    $hf = [string]$xl.Run('AttachLink_HashEngine', $fp, 0)
    Check ($hf -eq $expected) ("回退链（生产分派，CAPI 不可用）{0} 与 Get-FileHash 一致" -f (Split-Path $fp -Leaf))
    $fallbackChecked++
}
$hcF = [long]$xl.Run('AttachLink_GetHashCount')
Check ($hcF -eq $fallbackChecked) "回退路径每次生产分派仍只计 1 次真实计算（实际 $hcF / 应为 $fallbackChecked）"
$capiDirect = [string]$xl.Run('AttachLink_HashEngine', "$engDir\abc.bin", 1)
Check ($capiDirect -eq '') 'CAPI 标成不可用后，直调 CAPI（engine 1）返回空串（本会话不重试）'
$null = $xl.Run('AttachLink_TestCapiState', 0)      # 还原为「尚未自检」：下次使用重新自检
$hBack = [string]$xl.Run('AttachLink_HashEngine', "$engDir\abc.bin", 0)
Check ($hBack -eq 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad') '还原后重新自检通过，CAPI 恢复为主路径'
Check ([string]$xl.Run('AttachLink_HashEngineState') -match 'capi=1') '还原并重新自检后 HashEngineState 报 capi=1'

try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
Start-Sleep 2
if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
