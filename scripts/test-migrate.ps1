# =====================================================================
# AttachLink 归档迁移功能测试（自动化夹具）
# ---------------------------------------------------------------------
# 复刻真实场景：账套附件夹（含未引用文件）+ 归档目录两个底稿（链接被
# Excel 另存重基为 ../ 相对路径指回账套）+ 归档目录残留旧手工副本。
# 验证移动语义与链接改写：源夹移除、账套根保留、同名覆盖以源为准、
# 未引用文件一并移动、多出的旧文件保留、全部链接改写为「附件\文件名」。
# V5：文件被独占占用导致移动失败时，该文件链接保持原地址不改写、汇总单列。
# 用法：powershell -ExecutionPolicy Bypass -File test-migrate.ps1
# 期望末行：RESULT: PASS（夹具自动清理；FAIL 时保留现场供检查）
# 注意：本文件含中文，必须保存为 UTF-8 with BOM（PS5.1 编码陷阱见 WORKLOG #11）。
# =====================================================================
$ErrorActionPreference = 'Stop'
$attach = [string]([char]0x9644 + [char]0x4EF6)   # 附件
$base = Join-Path $env:TEMP 'opencode\alm'
$zht  = "$base\zht\14370003000"                   # 模拟账套（编制阶段）
$srcA = "$zht\$attach"                            # 源附件文件夹
$arch = "$base\arch"                              # 模拟归档目录（如 财报底稿）
$dstA = "$arch\$attach"
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
if (-not (Test-Path $xlam)) { throw "未找到 $xlam，请先运行 build\build.ps1。" }

# ---- 构造夹具 ----
Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $srcA, $arch | Out-Null
foreach ($n in 'a.pdf','b.pdf','c.pdf') { Set-Content "$srcA\$n" -Value "dummy-$n" }   # c.pdf 未被引用
New-Item -ItemType Directory -Force -Path $dstA | Out-Null
Set-Content "$dstA\old-extra.pdf" -Value 'manual-copy-leftover'                        # 旧手工复制多出的文件应保留
Set-Content "$dstA\a.pdf" -Value 'STALE-manual-copy'                                   # 旧手工副本应被账套原件覆盖

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
# WP-A 两条绝对链接 / WP-B 一条 —— 保存时会被 Excel 重基为 ../ 相对路径（即真实归档件形态）
foreach ($pair in @(@('WP-A.xlsx', @("$srcA\a.pdf","$srcA\b.pdf")), @('WP-B.xlsx', @("$srcA\a.pdf")))) {
    $wb = $xl.Workbooks.Add()
    $i = 1
    foreach ($lnk in $pair[1]) {
        $c = $wb.Worksheets.Item(1).Cells.Item($i, 1)
        $null = $wb.Worksheets.Item(1).Hyperlinks.Add($c, $lnk, '', 'attach' + [char]0x20 + $i)
        $i++
    }
    $wb.SaveAs("$arch\$($pair[0])", 51)
    $wb.Close($false)
}

# ---- 以归档目录中的 WP-A 为宿主运行迁移核心 ----
$null = $xl.Workbooks.Open($xlam)
# v1.6.4：加载项不再带 Workbook_Open，打开不隐式初始化——菜单、快捷键、应用事件都由显式初始化建立
$null = $xl.Run('AttachLink_Initialize')
$null = $xl.Workbooks.Open("$arch\WP-A.xlsx")
"--- 迁移汇总 ---"
$xl.Run('AttachLink_ArchiveMigrateCore', $srcA)

# ---- 断言 ----
$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }

Check (-not (Test-Path $srcA))                          '源附件文件夹移动后移除'
Check (Test-Path $zht)                                  '账套根目录保留'
Check (Test-Path "$dstA\a.pdf")                         '目标 a.pdf 存在'
Check ((Get-Content "$dstA\a.pdf") -eq 'dummy-a.pdf')   '目标 a.pdf 为账套原件（旧手工副本被覆盖）'
Check (Test-Path "$dstA\c.pdf")                         '未引用的 c.pdf 一并移动'
Check (Test-Path "$dstA\old-extra.pdf")                 '目标多出的旧文件保留'
Check (-not (Test-Path "$zht\$attach"))                 '账套中不再残留附件夹'

foreach ($wp in 'WP-A.xlsx','WP-B.xlsx') {
    $wb = $xl.Workbooks.Open("$arch\$wp")   # 核心已保存关闭，重开验证落盘状态
    foreach ($hl in $wb.Worksheets.Item(1).Hyperlinks) { "  $wp link -> $($hl.Address)" }
    $wb.Close($false)
}
$wbA = $xl.Workbooks.Open("$arch\WP-A.xlsx")
$addr = @($wbA.Worksheets.Item(1).Hyperlinks | ForEach-Object { $_.Address })
Check ($addr -contains "$attach\a.pdf")                 'WP-A 链接1 已改写为相对地址'
Check ($addr -contains "$attach\b.pdf")                 'WP-A 链接2 已改写为相对地址'
$wbA.Close($false)
$wbB = $xl.Workbooks.Open("$arch\WP-B.xlsx")
$addrB = @($wbB.Worksheets.Item(1).Hyperlinks | ForEach-Object { $_.Address })
Check ($addrB -contains "$attach\a.pdf")                'WP-B 链接已改写为相对地址'
$wbB.Close($false)
Check (Test-Path "$arch\$attach\b.pdf")                 '改写后相对地址可解析到文件'

'--- V5：文件被独占占用 → 移动失败，链接不改写、汇总单列 ---'
$base5 = Join-Path $env:TEMP 'opencode\alm5'
$zht5 = "$base5\zht\14370003000"; $srcA5 = "$zht5\$attach"
$arch5 = "$base5\arch"; $dstA5 = "$arch5\$attach"
Remove-Item $base5 -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $srcA5, $arch5 | Out-Null
Set-Content "$srcA5\x.pdf" -Value 'X-content'
Set-Content "$srcA5\y.pdf" -Value 'Y-content'
$wbC = $xl.Workbooks.Add()
$i = 1
foreach ($lnk in @("$srcA5\x.pdf", "$srcA5\y.pdf")) {
    $null = $wbC.Worksheets.Item(1).Hyperlinks.Add($wbC.Worksheets.Item(1).Cells.Item($i, 1), $lnk, '', 'L' + $i)
    $i++
}
$wbC.SaveAs("$arch5\WP-C.xlsx", 51)
$wbC.Close($false)
# 迁移前链接地址基线（保存后已被 Excel 重基为相对路径）
$wbPre = $xl.Workbooks.Open("$arch5\WP-C.xlsx")
$addrXpre = [string]$wbPre.Worksheets.Item(1).Hyperlinks.Item(1).Address
$wbPre.Close($false)
# 独占占用 x.pdf 后运行迁移
$lock5 = [IO.File]::Open("$srcA5\x.pdf", 'Open', 'Read', 'None')
try {
    $null = $xl.Workbooks.Open("$arch5\WP-C.xlsx")
    $s5 = [string]$xl.Run('AttachLink_ArchiveMigrateCore', $srcA5)
    ($s5 -split "`n") | ForEach-Object { "    $_" }
    Check ($s5 -like '*链接保持原样*' -and $s5 -like '*x.pdf*') 'V5: 汇总单列移动失败文件'
    Check (Test-Path "$srcA5\x.pdf")                          'V5: 占用文件仍在源处（移动失败）'
    Check (-not (Test-Path "$dstA5\x.pdf"))                   'V5: 归档附件夹无该文件'
    Check ((Test-Path "$dstA5\y.pdf") -and -not (Test-Path "$srcA5\y.pdf")) 'V5: 其余文件正常移动'
    $wbC2 = $xl.Workbooks.Open("$arch5\WP-C.xlsx")
    $addrXpost = [string]$wbC2.Worksheets.Item(1).Hyperlinks.Item(1).Address
    $addrYpost = [string]$wbC2.Worksheets.Item(1).Hyperlinks.Item(2).Address
    $wbC2.Close($false)
    Check ($addrXpost -eq $addrXpre)                          'V5: 失败文件链接地址保持原样'
    Check ($addrYpost -eq "$attach\y.pdf")                    'V5: 成功文件链接正常改写'
} finally { $lock5.Close() }
Remove-Item $base5 -Recurse -Force -ErrorAction SilentlyContinue

# ---- 清理 ----
try { foreach ($wb in @($xl.Workbooks)) { $wb.Saved = $true }; $xl.Quit() } catch {}
Start-Sleep 2
if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
