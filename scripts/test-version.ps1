# =====================================================================
# AttachLink 附件版本更新功能测试
# ---------------------------------------------------------------------
# 生命周期：登记 v1 → 受控版本更替（v1→v2，走真实 Core 入口）→ 体检「一致 v2」
#           + 更替记录列留痕 → 未走流程的替换（另一个附件）仍红字警示
#           → 基线 v2 格式与 v1 兼容（手工写 v1 格式基线可被读取）
#           → N-07 失败注入：新文件锁定 / 旧附件锁定 → 失败且原附件与基线不变
# 用法：powershell -ExecutionPolicy Bypass -File test-version.ps1
# 注意：含中文，必须 UTF-8 with BOM。
# =====================================================================
$ErrorActionPreference = 'Stop'
$attach = [string]([char]0x9644 + [char]0x4EF6)
$base = Join-Path $env:TEMP 'opencode\vertest'
$wbDir = "$base\proj"
$attDir = "$wbDir\$attach"
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
$baseline = "$attDir\_attachlink_baseline.csv"

Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $attDir | Out-Null
Set-Content "$attDir\ok.pdf" -Value 'good-version-1'
Set-Content "$attDir\tamper.pdf" -Value 'tamper-original'
Set-Content "$base\ok_v2.pdf" -Value 'good-version-2-fixed'
Set-Content "$base\ok_v3.pdf" -Value 'good-version-3'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$null = $xl.Workbooks.Open($xlam)
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$null = $ws.Hyperlinks.Add($ws.Cells.Item(1, 1), ($attach + '\ok.pdf'), '', 'OK')
$null = $ws.Hyperlinks.Add($ws.Cells.Item(2, 1), ($attach + '\tamper.pdf'), '', 'TP')
$wb.SaveAs("$wbDir\main.xlsx", 51)

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }
function Rows($xl, $wb) {
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)
    $idxName = $attach + [string]([char]0x6E05) + [string]([char]0x5355)
    $sheet = $wb.Worksheets.Item($idxName)
    $out = @{}
    $row = 1
    while ($row -le 40) {
        $c3 = $sheet.Cells.Item($row, 3).Value2
        if ($c3 -and "$c3" -like '*.pdf') {
            $out["$c3"] = @{ F = [string]$sheet.Cells.Item($row, 6).Value2; G = [string]$sheet.Cells.Item($row, 7).Value2 }
        }
        $row++
    }
    $out
}

'--- 1) 首次体检：登记 v1 ---'
$r1 = Rows $xl $wb
Check ($r1["$attach\ok.pdf"].F -eq '已登记 v1')        'ok.pdf 登记v1'
Check ($r1["$attach\tamper.pdf"].F -eq '已登记 v1')    'tamper.pdf 登记v1'
Check (Test-Path $baseline)                            '基线文件生成'

'--- 2) 受控版本更替（N-07：走真实 Core 入口 AttachLink_UpdateVersionCore）---'
$res = [string]$xl.Run('AttachLink_UpdateVersionCore', "$attDir\ok.pdf", "$base\ok_v2.pdf")
Check ($res -like '附件版本更新完成*')                              'Core 返回成功说明'
Check ((Get-FileHash "$attDir\ok.pdf" -Algorithm SHA256).Hash -eq `
       (Get-FileHash "$base\ok_v2.pdf" -Algorithm SHA256).Hash)      '更替后字节与新版一致（备份替换）'
$blNow2 = (Get-Content $baseline -Encoding UTF8) -join ';'
Check ($blNow2 -match 'ok\.pdf\|[0-9a-f]{64}\|2\|')                 '基线升级为 v2'

'--- 2b) N-07 失败注入①：新文件被独占锁定 → 失败且原附件/基线不变 ---'
$bytesBefore = [IO.File]::ReadAllBytes("$attDir\ok.pdf")
$blBefore = (Get-Content $baseline -Encoding UTF8) -join ';'
$lock1 = [IO.File]::Open("$base\ok_v3.pdf", 'Open', 'Read', 'None')
try {
    $res2 = [string]$xl.Run('AttachLink_UpdateVersionCore', "$attDir\ok.pdf", "$base\ok_v3.pdf")
} finally { $lock1.Close() }
Check ($res2 -like '失败*')                                         '新文件锁定 → Core 返回失败'
Check ($res2 -like '*原附件与基线均未改动*')                        '失败说明明示未改动'
Check (([IO.File]::ReadAllBytes("$attDir\ok.pdf") -join ',') -eq ($bytesBefore -join ',')) '原附件字节不变'
Check (((Get-Content $baseline -Encoding UTF8) -join ';') -eq $blBefore)                   '基线不变'

'--- 2c) N-07 失败注入②：旧附件被独占锁定 → 替换失败（备份移动被拒）且原附件/基线不变 ---'
$lock2 = [IO.File]::Open("$attDir\ok.pdf", 'Open', 'Read', 'None')
try {
    $res3 = [string]$xl.Run('AttachLink_UpdateVersionCore', "$attDir\ok.pdf", "$base\ok_v3.pdf")
} finally { $lock2.Close() }
Check ($res3 -like '失败：替换旧版本失败*')                         '旧附件锁定 → 替换失败'
Check ($res3 -like '*原附件与基线均未改动*')                        '失败说明明示未改动'
Check (([IO.File]::ReadAllBytes("$attDir\ok.pdf") -join ',') -eq ($bytesBefore -join ',')) '原附件字节不变'
Check (((Get-Content $baseline -Encoding UTF8) -join ';') -eq $blBefore)                   '基线不变'
Check (-not (Test-Path "$attDir\ok.pdf.bak"))                       '无 .bak 残留'
Check (-not (Test-Path "$attDir\ok_v3.pdf.updtmp"))                 '无 .updtmp 残留'
'--- 3) 体检：v2 一致 + 更替记录列 ---'
$r3 = Rows $xl $wb
Check ($r3["$attach\ok.pdf"].F -eq '一致 v2')                     'ok.pdf 一致 v2'
Check ($r3["$attach\ok.pdf"].G -like 'v1→v2 20*')                   '更替记录列留痕（v1→v2 + 时间）'
Check ($r3["$attach\tamper.pdf"].G -eq '')                        'v1 附件无更替记录'

'--- 4) 未走流程的替换仍红字 ---'
Set-Content "$attDir\tamper.pdf" -Value 'CLOBBERED'
$r4 = Rows $xl $wb
Check ($r4["$attach\tamper.pdf"].F -eq '■ 内容变更')              '未走流程替换仍红字'
Check ($r4["$attach\ok.pdf"].F -eq '一致 v2')                     '受控更替不受影响'

'--- 5) v1 旧格式基线兼容 ---'
Remove-Item $baseline -Force
Remove-Item "$attDir\tamper.pdf" -Force
Set-Content "$attDir\tamper.pdf" -Value 'tamper-original'
$h = (Get-FileHash "$attDir\tamper.pdf" -Algorithm SHA256).Hash.ToLower()
Set-Content $baseline -Value "tamper.pdf|$h" -Encoding UTF8       # 旧 v1 两字段格式
$r5 = Rows $xl $wb
Check ($r5["$attach\tamper.pdf"].F -eq '一致 v1')                 'v1 旧格式基线读作版本1且一致'
$blNow = (Get-Content $baseline -Encoding UTF8) -join ';'
Check ($blNow -match 'tamper\.pdf\|[0-9a-f]{64}\|1\|')            '落盘自动升级为 v2 格式（含版本字段）'

try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
Start-Sleep 2
if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
