# =====================================================================
# AttachLink v1.2.0 功能测试：文件夹引用 / 菜单二级化 / 文件夹行体检
# + 修复批 V1：同名文件夹绝不覆盖（复用 / 整夹另存 名(2) 五场景）
# 用法：powershell -ExecutionPolicy Bypass -File test-folder.ps1
# 注意：含中文，必须 UTF-8 with BOM。
# =====================================================================
$ErrorActionPreference = 'Stop'
$attach = [string]([char]0x9644 + [char]0x4EF6)
$base = Join-Path $env:TEMP 'opencode\foldertest'
$wbDir = "$base\proj"
$attDir = "$wbDir\$attach"
$srcRoot = "$base\src\合同扫描件"
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
$baseline = "$attDir\_attachlink_baseline.csv"

Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path "$srcRoot\2023", "$wbDir" | Out-Null
Set-Content "$srcRoot\购销合同.pdf" -Value 'contract-A'
Set-Content "$srcRoot\2023\发票.pdf" -Value 'invoice-2023'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$null = $xl.Workbooks.Open($xlam)
$wb = $xl.Workbooks.Add()
$wb.SaveAs("$wbDir\main.xlsx", 51)

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }

'--- 1) 插入文件夹引用核心（复制 + 基线登记） ---'
$summary = $xl.Run('AttachLink_InsertFolderCore', $srcRoot)
$summary1 = $summary -replace "`r`n", ' | '
"    $summary1"
Check (Test-Path "$attDir\合同扫描件\购销合同.pdf")            '文件已复制入附件夹'
Check (Test-Path "$attDir\合同扫描件\2023\发票.pdf")           '子目录结构保留'
Check (Test-Path $baseline)                                    '基线文件生成'
$bl = Get-Content $baseline -Encoding UTF8
Check (($bl -match [regex]::Escape('合同扫描件\购销合同.pdf')).Count -ge 1)  '根文件已登记基线'
Check (($bl -match [regex]::Escape('合同扫描件\2023\发票.pdf')).Count -ge 1) '子目录文件已登记基线（相对子路径键）'

'--- 2) 清单体检测文件夹行 ---'
$null = $wb.Worksheets.Item(1).Hyperlinks.Add($wb.Worksheets.Item(1).Cells.Item(1,1), ($attach + '\合同扫描件'), '', '合同扫描件')
$summary2 = $xl.Run('AttachLink_BuildIndexCore', $wb)
$idxName = $attach + [string]([char]0x6E05) + [string]([char]0x5355)
$sheet = $wb.Worksheets.Item($idxName)
$rowE = $null; $rowF = $null
$row = 1
while ($row -le 40) {
    $c3 = $sheet.Cells.Item($row, 3).Value2
    if ("$c3" -like '*合同扫描件') { $rowE = [string]$sheet.Cells.Item($row, 5).Value2; $rowF = [string]$sheet.Cells.Item($row, 6).Value2; break }
    $row++
}
"    E=[$rowE] F=[$rowF]"
Check ($rowE -eq '文件夹')                        '状态列显示「文件夹」'
Check ($rowF -eq '2 文件一致')                    '基线列显示「2 文件一致」'

'--- 3) 篡改夹内文件 → 红字 ---'
Set-Content "$attDir\合同扫描件\2023\发票.pdf" -Value 'TAMPERED'
$null = $xl.Run('AttachLink_BuildIndexCore', $wb)
$rowF2 = $null; $row = 1
while ($row -le 40) {
    $c3 = $sheet.Cells.Item($row, 3).Value2
    if ("$c3" -like '*合同扫描件') { $rowF2 = [string]$sheet.Cells.Item($row, 6).Value2; break }
    $row++
}
Check ($rowF2 -eq '■ 内容变更')                   '夹内篡改被检测（红字）'

'--- 4) 右键菜单结构（1.4.x：AttachLink 5 顶项 + 归档子菜单 5 项；原生菜单项共存，仅做存在性断言） ---'
$ctrls = @($xl.CommandBars.Item('Cell').Controls)
$names = $ctrls | ForEach-Object { $_.Caption }
"    顶层: $($names -join ' | ')"
Check (@($ctrls | Where-Object { $_.Caption -eq '插入附件链接（可多选）' }).Count -eq 1) '顶项1 插入附件链接'
Check (@($ctrls | Where-Object { $_.Caption -eq '插入文件夹引用' }).Count -eq 1)       '顶项2 插入文件夹引用'
Check (@($ctrls | Where-Object { $_.Caption -eq '重命名链接显示名' }).Count -eq 1)     '顶项3 重命名链接显示名'
Check (@($ctrls | Where-Object { $_.Caption -eq '插入底稿引用' }).Count -eq 1)        '顶项4 插入底稿引用'
$popGuidang = @($ctrls | Where-Object { $_.Caption -eq '归档' })[0]
Check ($popGuidang.Type -eq 10 -and $popGuidang.Controls.Count -eq 5)          '子菜单「归档」=5 项'
Check (@($popGuidang.Controls | Where-Object { $_.Caption -eq '归档迁移（批量）' }).Count -eq 1) '  含 归档迁移'
Check (@($popGuidang.Controls | Where-Object { $_.Caption -eq '附件清单 / 体检' }).Count -eq 1) '  含 体检'
Check (@($popGuidang.Controls | Where-Object { $_.Caption -eq '附件断链修复' }).Count -eq 1)   '  含 断链修复'
Check (@($popGuidang.Controls | Where-Object { $_.Caption -eq '附件版本更新' }).Count -eq 1)    '  含 版本更新'
Check (@($popGuidang.Controls | Where-Object { $_.Caption -eq '附件删除（删链接+文件）' }).Count -eq 1) '  含 附件删除'
Check (@($ctrls | Where-Object { $_.Caption -eq '附件清单' }).Count -eq 0)     '旧「附件清单」父项已并入归档'

'--- 5) oracle 复核回归：G 列串行 / 路径防护 ---'
# 构造：v2 文件 → 文件夹 → 断链文件 三连行，断言 G 列不串行、断链行 F=—
$wb2 = $xl.Workbooks.Add()
$null = $xl.Run('AttachLink_BuildIndexCore', $wb2)  # 先生成清单表（避免同名表归属问题，先登记再构造）
Set-Content "$attDir\v2file.pdf" -Value 'v2-content'
$h2 = (Get-FileHash "$attDir\v2file.pdf" -Algorithm SHA256).Hash.ToLower()
Add-Content $baseline -Value "v2file.pdf|$h2|2|2026-11-18 22:00|2026-11-20 09:00" -Encoding UTF8
$ws2 = $wb2.Worksheets.Item(1)
$null = $ws2.Hyperlinks.Add($ws2.Cells.Item(1,1), ($attach + '\v2file.pdf'), '', 'V2')
$null = $ws2.Hyperlinks.Add($ws2.Cells.Item(2,1), ($attach + '\合同扫描件'), '', 'F')
$null = $ws2.Hyperlinks.Add($ws2.Cells.Item(3,1), ($attach + '\missing.pdf'), '', 'BROKEN')
$wb2.SaveAs("$wbDir\main2.xlsx", 51)
$null = $xl.Run('AttachLink_BuildIndexCore', $wb2)
$sheet2 = $wb2.Worksheets.Item($idxName)
$g1 = [string]$sheet2.Cells.Item(5, 7).Value2   # v2 文件行（表头3+空？按行序：4起数据）
$row = 4; $rows = @{}
while ($row -le 20) {
    $c3 = [string]$sheet2.Cells.Item($row, 3).Value2
    if ($c3) { $rows[$c3] = @{ F = [string]$sheet2.Cells.Item($row, 6).Value2; G = [string]$sheet2.Cells.Item($row, 7).Value2 } }
    $row++
}
Check ($rows["$attach\v2file.pdf"].G -like 'v1→v2*')                    'v2 文件行 G 列有更替记录'
Check ($rows["$attach\合同扫描件"].G -eq '')                            '文件夹行 G 列为空（不串行）'
Check ($rows["$attach\missing.pdf"].G -eq '')                           '断链行 G 列为空（不串行）'
Check ($rows["$attach\missing.pdf"].F -eq '—')                          '断链行 F=—'
# 外部绝对文件夹链接（父目录名恰为「附件」的外部路径）：显示「外部，未校验」
$extAttach = "$base\otherproj\附件"
New-Item -ItemType Directory -Force -Path "$extAttach\extfol" | Out-Null
Set-Content "$extAttach\extfol\x.pdf" -Value 'external'
$null = $ws2.Hyperlinks.Add($ws2.Cells.Item(4,1), "$extAttach\extfol", '', 'EXT')
$null = $xl.Run('AttachLink_BuildIndexCore', $wb2)
$sheet3 = $wb2.Worksheets.Item($idxName)
$row = 4; $extRow = $null
while ($row -le 20) {
    $c3 = [string]$sheet3.Cells.Item($row, 3).Value2
    if ($c3 -like '*extfol') { $extRow = @{ F = [string]$sheet3.Cells.Item($row, 6).Value2 }; break }
    $row++
}
Check ($extRow.F -eq '外部，未校验')                                   '外部绝对文件夹显示「外部，未校验」'
# 锁定文件（failCnt）+ 篡改并存：红字优先，不被「校验不可用」遮蔽（oracle #2）
$lockDir = "$attDir\锁定夹"
New-Item -ItemType Directory -Force -Path $lockDir | Out-Null
Set-Content "$lockDir\a.pdf" -Value 'orig'
Set-Content "$lockDir\b.pdf" -Value 'will-lock'
$null = $ws2.Hyperlinks.Add($ws2.Cells.Item(5,1), ($attach + '\锁定夹'), '', 'LK')
$null = $xl.Run('AttachLink_BuildIndexCore', $wb2)   # 先登记基线
Set-Content "$lockDir\a.pdf" -Value 'TAMPERED'
$lockStream = [IO.File]::Open("$lockDir\b.pdf", 'Open', 'Read', 'None')   # 独占锁 → VBA 哈希失败
try {
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb2)
    $sheet4 = $wb2.Worksheets.Item($idxName)
    $row = 4; $lkRow = $null
    while ($row -le 22) {
        $c3 = [string]$sheet4.Cells.Item($row, 3).Value2
        if ($c3 -like '*锁定夹') { $lkRow = @{ F = [string]$sheet4.Cells.Item($row, 6).Value2 }; break }
        $row++
    }
    "    锁定夹 F=[$($lkRow.F)]"
    Check ($lkRow.F -eq '■ 内容变更（含不可用）')                       '篡改红字优先于校验不可用（oracle #2）'
} finally {
    $lockStream.Close()
}
# 路径防护
$r1 = $xl.Run('AttachLink_InsertFolderCore', 'C:\')
Check ($r1 -like '失败：*根目录*')                                        '拒绝磁盘根目录'
$r2 = $xl.Run('AttachLink_InsertFolderCore', $attDir)
Check (($r2 -like '失败：*源文件夹内部*') -or ($r2 -like '失败：*已位于*'))  '拒绝附件夹内自引用（命中源内/自引防护）'
try { $wb2.Close($false) } catch {}

'--- 6) V1：同名文件夹绝不覆盖——逐文件三方比对，全一致复用、否则整夹另存 名(2) ---'
# 组 3 曾篡改 2023\发票.pdf——先恢复副本与源一致，场景②才有「全一致复用」前提
Set-Content "$attDir\合同扫描件\2023\发票.pdf" -Value 'invoice-2023'
# ② 源未变再次引用 → 直接复用、不复制
$s6b = [string]$xl.Run('AttachLink_InsertFolderCore', $srcRoot)
"    $($s6b -replace "`r`n", ' | ')"
Check ($s6b -like '*已直接复用*')                                       'V1-②: 内容一致直接复用'
Check (-not (Test-Path "$attDir\合同扫描件(2)"))                          'V1-②: 未产生 (2) 副本'
Check ($s6b -like ('*链接指向*' + $attach + '*合同扫描件*'))                  'V1-②: 汇总含链接指向现有副本'
# ③ 副本内某文件被用户改过 → 整夹另存，原副本字节不变
Set-Content "$attDir\合同扫描件\购销合同.pdf" -Value 'contract-A-USER-EDITED'
$hCopyBefore = (Get-FileHash "$attDir\合同扫描件\购销合同.pdf" -Algorithm SHA256).Hash
$s6c = [string]$xl.Run('AttachLink_InsertFolderCore', $srcRoot)
"    $($s6c -replace "`r`n", ' | ')"
Check ((Get-FileHash "$attDir\合同扫描件\购销合同.pdf" -Algorithm SHA256).Hash -eq $hCopyBefore) 'V1-③: 原副本字节不变（未被覆盖）'
Check (Test-Path "$attDir\合同扫描件(2)\购销合同.pdf")                      'V1-③: 出现 名(2)'
Check ((Get-FileHash "$attDir\合同扫描件(2)\购销合同.pdf" -Algorithm SHA256).Hash -eq (Get-FileHash "$srcRoot\购销合同.pdf" -Algorithm SHA256).Hash) 'V1-③: (2) 内容=源'
Check ($s6c -like '*未覆盖*' -and $s6c -like '*已另存为*合同扫描件(2)*')     'V1-③: 汇总明示未覆盖/已另存 名(2)'
$bl6c = Get-Content $baseline -Encoding UTF8
Check (($bl6c -match [regex]::Escape('合同扫描件(2)\购销合同.pdf')).Count -ge 1) 'V1-③: (2) 有自己的基线'
Check (($bl6c -match [regex]::Escape('合同扫描件\2023\发票.pdf')).Count -ge 1) 'V1-③: 旧键保留不被改写'
# ④ 副本与基线不一致但源==副本内容 → 仍另存（三方一致规则：copy==baseline==source）
$srcB = "$base\srcB\资料"
New-Item -ItemType Directory -Force -Path $srcB | Out-Null
Set-Content "$srcB\b1.pdf" -Value 'B1'
$s6d1 = [string]$xl.Run('AttachLink_InsertFolderCore', $srcB)             # 首次引用：复制+登记
Check ($s6d1 -like '*已复制到*')                                          'V1-④: 首次引用已复制'
Set-Content "$attDir\资料\b1.pdf" -Value 'B2'                              # 副本被改动（基线仍 B1）
Set-Content "$srcB\b1.pdf" -Value 'B2'                                     # 源同步为 B2（源==副本≠基线）
$s6d2 = [string]$xl.Run('AttachLink_InsertFolderCore', $srcB)
"    $($s6d2 -replace "`r`n", ' | ')"
Check (Test-Path "$attDir\资料(2)\b1.pdf")                                  'V1-④: 基线不符仍另存 名(2)'
Check ((Get-FileHash "$attDir\资料\b1.pdf" -Algorithm SHA256).Hash -eq (Get-FileHash "$srcB\b1.pdf" -Algorithm SHA256).Hash) 'V1-④: 原副本未被覆盖（仍=改动后内容 B2）'
# ⑤ 源新增文件 → 另存
$srcC = "$base\srcC\手册"
New-Item -ItemType Directory -Force -Path $srcC | Out-Null
Set-Content "$srcC\h1.pdf" -Value 'H1'
$null = $xl.Run('AttachLink_InsertFolderCore', $srcC)
Set-Content "$srcC\h2.pdf" -Value 'H2'                                     # 源多了 h2.pdf
$s6e = [string]$xl.Run('AttachLink_InsertFolderCore', $srcC)
Check (Test-Path "$attDir\手册(2)\h2.pdf")                                  'V1-⑤: 源新增文件 → 另存 名(2)'
# 同名路径其实是文件 → 整夹另存
$srcD = "$base\srcD\单文件占用"
New-Item -ItemType Directory -Force -Path $srcD | Out-Null
Set-Content "$srcD\d1.pdf" -Value 'D1'
Set-Content "$attDir\单文件占用" -Value 'iam-a-file'                       # 同名路径被文件占用
$s6f = [string]$xl.Run('AttachLink_InsertFolderCore', $srcD)
Check (Test-Path "$attDir\单文件占用(2)\d1.pdf")                            'V1-⑥: 同名路径是文件 → 另存 名(2)'

try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
Start-Sleep 2
if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
