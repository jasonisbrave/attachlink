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
Check ($r1.vals["$attach\a.pdf"] -eq '已登记基线')          'a.pdf 登记基线'
Check ($r1.vals["$attach\b.pdf"] -eq '已登记基线')          'b.pdf 登记基线'
Check ($r1.vals["$attach\broken.pdf"] -eq '—')             '断链行不参与基线（—）'
$h1 = (Get-Content $baseline) -join '|'
Check ($h1 -match 'a\.pdf\|[0-9a-f]{64}')                  '基线含 a.pdf 的64位哈希'

'--- 2) 模拟同步端同名静默覆盖（篡改 a.pdf）---'
Set-Content "$attDir\a.pdf" -Value 'CLOBBERED-by-sync'
$r2 = ColF $xl "$wbDir\main.xlsx"
Check ($r2.vals["$attach\a.pdf"] -eq '■ 内容变更')          '篡改被红字警示'
Check ($r2.vals["$attach\b.pdf"] -eq '一致')                '未动文件仍一致'
$h2 = (Get-Content $baseline) -join '|'
Check ($h1 -eq $h2)                                        '基线未被篡改洗白（哈希保持登记时值）'
Check ($r2.summary -match '内容变更 1')                     '汇总计入内容变更 1'

'--- 3) 有意替换：删基线重建 ---'
Remove-Item $baseline -Force
$r3 = ColF $xl "$wbDir\main.xlsx"
Check ($r3.vals["$attach\a.pdf"] -eq '已登记基线')          '重建基线后重新登记'

'--- 4) 二次体检全部一致 ---'
$r4 = ColF $xl "$wbDir\main.xlsx"
Check ($r4.vals["$attach\a.pdf"] -eq '一致')                'a.pdf 一致'
Check ($r4.vals["$attach\b.pdf"] -eq '一致')                'b.pdf 一致'

try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
Start-Sleep 2
if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
