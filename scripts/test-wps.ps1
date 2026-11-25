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
    $r = $wps.Run('AttachLink_SelfTest')
    Check ($r -eq $true) 'SelfTest（晚绑定改造后可编译）'
    $v = $wps.Run('AttachLink_GetVersion')
    Check ("$v" -eq '1.5.1') "GetVersion=1.5.1（实际 $v）"
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

try { foreach ($w in @($wps.Workbooks)) { $w.Saved = $true }; $wps.Quit() } catch {}
Start-Sleep 3
if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
