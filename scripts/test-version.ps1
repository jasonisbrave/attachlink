# =====================================================================
# AttachLink 附件版本更新功能测试
# ---------------------------------------------------------------------
# 生命周期：登记 v1 → 受控版本更替（v1→v2，旧文件删除）→ 体检「一致 v2」
#           + 更替记录列留痕 → 未走流程的替换（另一个附件）仍红字警示
#           → 基线 v2 格式与 v1 兼容（手工写 v1 格式基线可被读取）
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

'--- 2) 受控版本更替（模拟右键流程核心：直接调 Core 不可行，含对话框；改为验证端态）---'
'    （AttachLink_UpdateVersion 含 FileDialog/MsgBox 交互，自动化走端态断言：手工执行等价文件与基线操作）'
# 端态等价：新文件覆盖 + 基线 v2 —— 由 VBA 侧完成需交互，此处先验证「体检端」对 v2 基线的呈现
$bl = Get-Content $baseline -Encoding UTF8
$bl2 = $bl | ForEach-Object {
    $p = $_ -split '\|'
    if ($p[0] -eq 'ok.pdf') { "ok.pdf|$((Get-FileHash "$base\ok_v2.pdf" -Algorithm SHA256).Hash.ToLower())|2|2026-11-18 22:00|2026-11-20 10:30" }
    else { $_ }
}
Set-Content $baseline -Value $bl2 -Encoding UTF8
Copy-Item "$base\ok_v2.pdf" "$attDir\ok.pdf" -Force
'--- 3) 体检：v2 一致 + 更替记录列 ---'
$r3 = Rows $xl $wb
Check ($r3["$attach\ok.pdf"].F -eq '一致 v2')                     'ok.pdf 一致 v2'
Check ($r3["$attach\ok.pdf"].G -like 'v1→v2 2026-11-20*')         '更替记录列留痕（v1→v2 + 时间）'
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
