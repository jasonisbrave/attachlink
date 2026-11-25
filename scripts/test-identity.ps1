# =====================================================================
# AttachLink v1.5.0 T2/T3 测试：附件身份（插入即登记/同内容复用/异内容
# 自动另存）+ 可信断链重连 + 基线安全读写 + 哈希去重
# ---------------------------------------------------------------------
# 覆盖（对应任务书 13 组 + 修复批 V）：
#   1  插入新文件 → 归档 + 基线当场登记 + 相对链接
#   2  同名同内容 → 静默复用（不复制、无 (2)、汇总含「复用」）
#   3  同名异内容 → 自动另存 名(2)（原文件不动、基线两条、汇总含「另存」）
#   4  与 (2) 同内容 → 复用 (2)，不产生 (3)
#   5  体检全部「一致 v1」、无红字
#   6  手写 v1（2 字段）基线 → 读作「一致 v1」，落盘升级 5 字段
#   7  畸形基线 → 体检后字节不变、汇总含「基线损坏」、内容列红字
#   8  6 字段（未来格式）基线 → 可比对出「一致」，字节不变（只读不覆盖）
#   9  基线被独占锁定 → 插入照常但汇总如实报「未持久化」，原文件完整
#   10 断链修复七场景（同名同内容/按内容找回/同名异内容拒绝/候选不唯一/
#      无基线未校验/仅改指/文件夹三种）
#   11 显示名在重连前后不变
#   12 底稿引用同名异内容 → 自动另存（不弹窗）、基线已登记
#   13 哈希去重：同一文件被多条断链命中只真实计算一次
#   14 O-04 基线损坏 = 失败关闭（插入/修复在任何写入前拒绝）
#   15 O-05 副本与基线不一致 → 另存新名并明示（插入/底稿引用）
#   16 O-02 子目录断链不回退借根目录同名键
#   17 V7 文件名含 #：单文件插入/文件夹引用/底稿引用三类入口复制前拒绝
#   18 V11b 复制后源大小变化 → 基线以目标实际哈希登记
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File test-identity.ps1
# 期望末行：RESULT: PASS（夹具自动清理；FAIL 保留现场）
# 注意：含中文，必须 UTF-8 with BOM。
# =====================================================================
$ErrorActionPreference = 'Stop'
$attach = [string]([char]0x9644 + [char]0x4EF6)   # 附件
$idxName = $attach + [string]([char]0x6E05) + [string]([char]0x5355)   # 附件清单
$base = Join-Path $env:TEMP 'opencode\identitytest'
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
if (-not (Test-Path $xlam)) { throw "未找到 $xlam，请先运行 build\build.ps1。" }

Add-Type -Namespace WI -Name U32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }
function Sha($p) { (Get-FileHash $p -Algorithm SHA256).Hash.ToLower() }

# ---- 夹具 ----
Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
$wbDir = "$base\proj";  $attDir = "$wbDir\$attach";  $baseline = "$attDir\_attachlink_baseline.csv"
$srcA = "$base\srcA"; $srcB = "$base\srcB"; $srcC = "$base\srcC"; $srcD = "$base\srcD"
New-Item -ItemType Directory -Force -Path $attDir, $srcA, $srcB, $srcC, $srcD | Out-Null
Set-Content "$srcA\new1.pdf" -Value 'N1-original'
Set-Content "$srcB\new1.pdf" -Value 'N1-original'          # 与 A 同内容
Set-Content "$srcC\new1.pdf" -Value 'N1-DIFFERENT-content' # 异内容
Set-Content "$srcD\new1.pdf" -Value 'N1-DIFFERENT-content' # 与 C 同内容

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$procId = 0
[void][WI.U32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$procId)
"Excel PID: $procId"
$null = $xl.Workbooks.Open($xlam)     # 第一时间装载 dist 副本（防常驻看门狗抢装）

# 超时兜底：15 分钟后脚本仍未结束（COM 卡死/意外模态）→ 按 PID 结束本实例
$sentinel = "$base\done.flag"
$killer = Start-Process powershell -PassThru -WindowStyle Hidden -ArgumentList '-NoProfile', '-Command', `
    "Start-Sleep -Seconds 900; if (-not (Test-Path '$sentinel')) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }"

# 体检行读取：相对路径 → F 列文本 + 字体色
function Get-IndexRows($wb) {
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)
    $sheet = $wb.Worksheets.Item($idxName)
    $out = @{}
    $row = 4
    while ($row -le 60) {
        $c3 = [string]$sheet.Cells.Item($row, 3).Value2
        if ($c3) {
            $out[$c3] = @{ F = [string]$sheet.Cells.Item($row, 6).Value2; Color = [int]$sheet.Cells.Item($row, 6).Font.Color }
        }
        $row++
    }
    $out
}

try {
    $wb1 = $xl.Workbooks.Add()
    $ws1 = $wb1.Worksheets.Item(1)
    $wb1.SaveAs("$wbDir\main.xlsx", 51)

    '--- 1) 插入新文件：归档 + 基线当场登记 + 相对链接 ---'
    $s1 = [string]$xl.Run('AttachLink_InsertCore', "$srcA\new1.pdf", $ws1.Range('A1'))
    "    $s1"
    Check (Test-Path "$attDir\new1.pdf")                              '1: 文件已归档'
    Check (Test-Path $baseline)                                       '1: 基线文件当场生成'
    $bl = [string]([IO.File]::ReadAllText($baseline))
    Check ($bl -match ('(?m)^new1\.pdf\|' + (Sha "$attDir\new1.pdf") + '\|1\|')) '1: 基线含该文件哈希（v1）'
    Check ($ws1.Cells.Item(1, 1).Hyperlinks.Item(1).Address -eq "$attach\new1.pdf") '1: 链接地址为相对路径'
    Check ($s1 -like '*已插入 1*')                                     '1: 汇总已插入 1'

    '--- 2) 同名同内容：静默复用 ---'
    $s2 = [string]$xl.Run('AttachLink_InsertCore', "$srcB\new1.pdf", $ws1.Range('A2'))
    "    $s2"
    $files2 = @(Get-ChildItem $attDir -File | Where-Object { $_.Name -ne '_attachlink_baseline.csv' })
    Check ($files2.Count -eq 1)                                        '2: 文件数不变（未重复复制）'
    Check (-not (Test-Path "$attDir\new1(2).pdf"))                     '2: 未产生 (2)'
    Check ($ws1.Cells.Item(2, 1).Hyperlinks.Item(1).Address -eq "$attach\new1.pdf") '2: 链接指向同一文件'
    Check ($s2 -like '*复用*')                                          '2: 汇总含「复用」'

    '--- 3) 同名异内容：自动另存 名(2) ---'
    $hashOrig = Sha "$attDir\new1.pdf"
    $s3 = [string]$xl.Run('AttachLink_InsertCore', "$srcC\new1.pdf", $ws1.Range('A3'))
    "    $s3"
    Check (Test-Path "$attDir\new1(2).pdf")                            '3: 生成 new1(2).pdf'
    Check ((Sha "$attDir\new1.pdf") -eq $hashOrig)                     '3: 原文件哈希不变（未被覆盖）'
    $bl3 = [string]([IO.File]::ReadAllText($baseline))
    Check (($bl3 -match '(?m)^new1\.pdf\|') -and ($bl3 -match '(?m)^new1\(2\)\.pdf\|')) '3: 基线含两条'
    Check ($ws1.Cells.Item(3, 1).Hyperlinks.Item(1).Address -eq "$attach\new1(2).pdf") '3: 链接指向 (2)'
    Check ($s3 -like '*另存*' -and $s3 -like '*new1(2).pdf*')           '3: 汇总含「另存」与文件名'

    '--- 4) 与 (2) 同内容：复用 (2)，不产生 (3) ---'
    $s4 = [string]$xl.Run('AttachLink_InsertCore', "$srcD\new1.pdf", $ws1.Range('A4'))
    "    $s4"
    Check (-not (Test-Path "$attDir\new1(3).pdf"))                     '4: 未产生 (3)'
    Check ($ws1.Cells.Item(4, 1).Hyperlinks.Item(1).Address -eq "$attach\new1(2).pdf") '4: 链接复用 (2)'
    Check ($s4 -like '*复用*')                                          '4: 汇总含「复用」'

    '--- 5) 体检：全部「一致 v1」、无红字 ---'
    $rows5 = Get-IndexRows $wb1
    Check ($rows5["$attach\new1.pdf"].F -eq '一致 v1')                  '5: new1.pdf 一致 v1（插入即登记后不再显示已登记）'
    Check ($rows5["$attach\new1(2).pdf"].F -eq '一致 v1')               '5: new1(2).pdf 一致 v1'
    $redCnt = @($rows5.GetEnumerator() | Where-Object { $_.Value.Color -eq 192 }).Count
    Check ($redCnt -eq 0)                                              '5: 无红字'

    '--- 6) 手写 v1（2 字段）基线：读作一致 v1，落盘升级 5 字段 ---'
    $h1 = Sha "$attDir\new1.pdf"; $h2 = Sha "$attDir\new1(2).pdf"
    Set-Content $baseline -Value @("new1.pdf|$h1", "new1(2).pdf|$h2") -Encoding UTF8
    Set-Content "$srcA\new2.pdf" -Value 'N2-content'
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\new2.pdf", $ws1.Range('A5'))   # 触发新登记 → 落盘
    $bl6 = [string]([IO.File]::ReadAllText($baseline))
    Check ($bl6 -match '(?m)^new1\.pdf\|[0-9a-f]{64}\|1\|')            '6: 落盘升级为 5 字段（含版本）'
    $rows6 = Get-IndexRows $wb1
    Check ($rows6["$attach\new1.pdf"].F -eq '一致 v1')                  '6: v1 基线读作「一致 v1」'

    '--- 7) 畸形基线：字节不变 + 汇总「基线损坏」+ 内容列红字 ---'
    $snap7 = [IO.File]::ReadAllBytes($baseline)
    Add-Content $baseline -Value 'bad|three|fields' -Encoding UTF8
    $snap7b = [IO.File]::ReadAllBytes($baseline)
    $s7 = [string]$xl.Run('AttachLink_BuildIndexCore', $wb1)
    "    $($s7 -split "`n" | Select-Object -First 1)"
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes($baseline)) -eq [Convert]::ToBase64String($snap7b)) '7: 基线文件字节不变'
    Check ($s7 -like '*基线损坏*')                                      '7: 汇总含「基线损坏」'
    $sheet7 = $wb1.Worksheets.Item($idxName)
    $f7 = $null; $c7 = 0
    $row = 4
    while ($row -le 60) {
        $c3 = [string]$sheet7.Cells.Item($row, 3).Value2
        if ($c3 -like '*new1.pdf') { $f7 = [string]$sheet7.Cells.Item($row, 6).Value2; $c7 = [int]$sheet7.Cells.Item($row, 6).Font.Color; break }
        $row++
    }
    Check ($f7 -eq '■ 基线损坏')                                        '7: 内容基线列显示「■ 基线损坏」（非一致/已登记）'
    Check ($c7 -eq 192)                                                '7: 内容基线列为红字'
    [IO.File]::WriteAllBytes($baseline, $snap7)                        # 还原供后续组使用

    '--- 8) 6 字段（未来格式）基线：可比对，字节不变 ---'
    $h1 = Sha "$attDir\new1.pdf"; $h2 = Sha "$attDir\new1(2).pdf"; $h3 = Sha "$attDir\new2.pdf"
    Set-Content $baseline -Value @(
        "new1.pdf|$h1|1|2026-11-24 10:00|2026-11-24 10:00|future-field",
        "new1(2).pdf|$h2|1|2026-11-24 10:00|2026-11-24 10:00|future-field",
        "new2.pdf|$h3|1|2026-11-24 10:00|2026-11-24 10:00|future-field") -Encoding UTF8
    $snap8 = [IO.File]::ReadAllBytes($baseline)
    $s8 = [string]$xl.Run('AttachLink_BuildIndexCore', $wb1)
    $rows8 = Get-IndexRows $wb1
    Check ($rows8["$attach\new1.pdf"].F -eq '一致 v1')                  '8: 未来格式基线可比对出「一致」'
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes($baseline)) -eq [Convert]::ToBase64String($snap8)) '8: 基线文件字节不变（只读不覆盖）'
    Check ($s8 -notlike '*基线损坏*')                                   '8: 不误报基线损坏'

    '--- 9) 基线被独占锁定：O-04 失败关闭（如实报「无法读取」，不做任何写入） ---'
    Set-Content "$srcA\new3.pdf" -Value 'N3-content'
    $lock = [IO.File]::Open($baseline, 'Open', 'Read', 'None')
    try {
        $s9 = [string]$xl.Run('AttachLink_InsertCore', "$srcA\new3.pdf", $ws1.Range('A6'))
        "    $($s9 -split "`n" | Select-Object -First 1)"
        Check ($s9 -like '失败：基线文件损坏或无法读取*')                        '9: 失败关闭并明示「无法读取」'
        Check ($s9 -like '*读取失败*')                                       '9: 原因如实（读取失败：文件被占用）'
        Check (-not (Test-Path "$attDir\new3.pdf"))                          '9: 文件未复制（任何写入之前拒绝）'
        Check ([int]$ws1.Cells.Item(6, 1).Hyperlinks.Count -eq 0)            '9: 链接未插入'
    } finally { $lock.Close() }
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes($baseline)) -eq [Convert]::ToBase64String($snap8)) '9: 原基线完整未动'
    Remove-Item $baseline -Force
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb1)                 # 重建基线（含 new3）
    $wb1.Close($true)

    # ================= 断链修复（proj2） =================
    $wbDir2 = "$base\proj2"; $attDir2 = "$wbDir2\$attach"; $baseline2 = "$attDir2\_attachlink_baseline.csv"
    $repDir = "$base\repair"
    New-Item -ItemType Directory -Force -Path $attDir2, $repDir, "$base\foldersrc\foldC", "$base\foldersrc\foldA" | Out-Null
    $wb2 = $xl.Workbooks.Add()
    $ws2 = $wb2.Worksheets.Item(1)
    $wb2.SaveAs("$wbDir2\main2.xlsx", 51)

    foreach ($n in 'ra', 'rb', 'rc', 'rd', 'dup') { Set-Content "$srcA\$n.pdf" -Value "$n-original-content" }
    Set-Content "$srcA\re.pdf" -Value 're-original-content'
    Set-Content "$srcA\rf.pdf" -Value 'rf-original-content'
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\ra.pdf", $ws2.Range('A1'))
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\rb.pdf", $ws2.Range('A2'))
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\rc.pdf", $ws2.Range('A3'))
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\rd.pdf", $ws2.Range('A4'))
    Copy-Item "$srcA\re.pdf" "$attDir2\re.pdf"                          # 手工放置：不登记基线
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(5, 1), ($attach + '\re.pdf'), '', 'RE')
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\rf.pdf", $ws2.Range('A6'))
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\dup.pdf", $ws2.Range('A7'))
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(8, 1), ($attach + '\dup.pdf'), '', 'DUP2')
    # (f) 指回旧位置的链接：改为绝对路径（父目录名为「附件」的旧位置，文件不存在）
    $ws2.Cells.Item(6, 1).Hyperlinks.Item(1).Address = "$base\oldproj\$attach\rf.pdf"
    # 文件夹：foldC（将一致）、foldA（将不符）走 InsertFolderCore 登记；foldB 无基线
    Set-Content "$base\foldersrc\foldC\c1.pdf" -Value 'FC-content'
    Set-Content "$base\foldersrc\foldA\f1.pdf" -Value 'FA-orig'
    $null = $xl.Run('AttachLink_InsertFolderCore', "$base\foldersrc\foldC")
    $null = $xl.Run('AttachLink_InsertFolderCore', "$base\foldersrc\foldA")
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(9, 1), ($attach + '\foldC'), '', 'FC')
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(10, 1), ($attach + '\foldA'), '', 'FA')
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(11, 1), ($attach + '\foldB'), '', 'FB')
    $ws2.Cells.Item(1, 1).Hyperlinks.Item(1).TextToDisplay = 'RA-显示名'
    $dispBefore = $ws2.Cells.Item(1, 1).Hyperlinks.Item(1).TextToDisplay

    # 修复源目录：同名同内容 ra/dup；rb 同名异内容 + 同内容异名候选；rc 只有同名异内容；
    # rd 两个同内容异名候选；re 无基线按名；foldC 一致 / foldA 不符 / foldB 无基线
    Set-Content "$repDir\ra.pdf" -Value 'ra-original-content'
    Set-Content "$repDir\rb.pdf" -Value 'RB-DIFFERENT'
    Set-Content "$repDir\renamed-rb.pdf" -Value 'rb-original-content'
    Set-Content "$repDir\rc.pdf" -Value 'RC-DIFFERENT'
    Set-Content "$repDir\x1.pdf" -Value 'rd-original-content'
    Set-Content "$repDir\x2.pdf" -Value 'rd-original-content'
    Set-Content "$repDir\re.pdf" -Value 're-original-content'
    Set-Content "$repDir\dup.pdf" -Value 'dup-original-content'
    New-Item -ItemType Directory -Force -Path "$repDir\foldC", "$repDir\foldA", "$repDir\foldB" | Out-Null
    Set-Content "$repDir\foldC\c1.pdf" -Value 'FC-content'
    Set-Content "$repDir\foldA\f1.pdf" -Value 'FA-TAMPERED'
    Set-Content "$repDir\foldB\f2.pdf" -Value 'FB-content'

    # 制造断链：移走附件
    foreach ($n in 'ra', 'rb', 'rc', 'rd', 're', 'dup') { Remove-Item "$attDir2\$n.pdf" -Force }
    Remove-Item "$attDir2\foldA", "$attDir2\foldB", "$attDir2\foldC" -Recurse -Force -ErrorAction SilentlyContinue

    '--- 10) 断链修复七场景 ---'
    $s10 = [string]$xl.Run('AttachLink_RepairCore', $wb2, $repDir)
    ($s10 -split "`n") | ForEach-Object { "    $_" }
    # (a) 同名同内容
    Check ((Test-Path "$attDir2\ra.pdf") -and ((Sha "$attDir2\ra.pdf") -eq (Sha "$srcA\ra.pdf"))) '10a: ra.pdf 已落回'
    Check ($ws2.Cells.Item(1, 1).Hyperlinks.Item(1).Address -eq "$attach\ra.pdf") '10a: 链接地址不变'
    Check ($s10 -like '*内容一致（同名）*')                              '10a: 汇总含「内容一致（同名）」'
    # (b) 同名异内容 + 同内容异名候选 → 按内容找回，以原名落回
    Check ((Test-Path "$attDir2\rb.pdf") -and ((Sha "$attDir2\rb.pdf") -eq (Sha "$srcA\rb.pdf"))) '10b: rb.pdf 以原名落回（内容=原件）'
    Check ($ws2.Cells.Item(2, 1).Hyperlinks.Item(1).Address -eq "$attach\rb.pdf") '10b: 链接地址不变'
    Check ($s10 -like '*按内容找回*' -and $s10 -like '*rb.pdf*renamed-rb.pdf*') '10b: 汇总注明源文件名'
    # (c) 只有同名异内容 → 不重连不复制
    Check (-not (Test-Path "$attDir2\rc.pdf"))                          '10c: rc.pdf 未复制'
    Check ($ws2.Cells.Item(3, 1).Hyperlinks.Item(1).Address -eq "$attach\rc.pdf") '10c: 链接保持断链（地址未动）'
    Check ($s10 -like '*同名但内容不同*')                                '10c: 汇总含「同名但内容不同」'
    # (d) 两个同内容异名候选 → 候选不唯一
    Check (-not (Test-Path "$attDir2\rd.pdf"))                          '10d: rd.pdf 未复制'
    Check ($s10 -like '*候选不唯一*' -and $s10 -like '*x1.pdf*' -and $s10 -like '*x2.pdf*') '10d: 汇总列出候选'
    # (e) 无基线 → 按名重连 + 登记 v1
    Check (Test-Path "$attDir2\re.pdf")                                 '10e: re.pdf 已落回'
    Check ($ws2.Cells.Item(5, 1).Hyperlinks.Item(1).Address -eq "$attach\re.pdf") '10e: 链接已重连'
    Check ($s10 -like '*未校验*')                                        '10e: 汇总标注未校验'
    $blR = [string]([IO.File]::ReadAllText($baseline2))
    Check ($blR -match '(?m)^re\.pdf\|[0-9a-f]{64}\|1\|')               '10e: 基线登记 v1（建立身份）'
    # (f) 指回旧位置 + 本附件夹已有同内容同名文件 → 只改指不复制
    Check ($ws2.Cells.Item(6, 1).Hyperlinks.Item(1).Address -eq "$attach\rf.pdf") '10f: 链接已改指本附件夹'
    Check ($s10 -like '*仅改指*')                                        '10f: 汇总含「仅改指」'
    # (g) 文件夹：一致合并 / 不符不合并 / 无基线按名
    Check (Test-Path "$attDir2\foldC\c1.pdf")                           '10g: foldC 内容一致已合并'
    Check (-not (Test-Path "$attDir2\foldA"))                           '10g: foldA 与基线不符未合并'
    Check (Test-Path "$attDir2\foldB\f2.pdf")                           '10g: foldB 无基线按名重连'
    Check ($s10 -like '*与基线不符未合并 1*')                            '10g: 汇总计数（不符 1）'
    $blR2 = [string]([IO.File]::ReadAllText($baseline2))
    Check ($blR2 -match ('(?m)^' + [regex]::Escape('foldb\f2.pdf') + '\|')) '10g: foldB 文件已登记基线'

    '--- 11) 显示名在重连前后不变 ---'
    Check ($ws2.Cells.Item(1, 1).Hyperlinks.Item(1).TextToDisplay -eq $dispBefore) '11: 显示名不变'
    $wb2.Close($true)

    '--- 12) 底稿引用：同名异内容自动另存（不弹窗）+ 基线登记 ---'
    $wbDir3 = "$base\proj3"; $attDir3 = "$wbDir3\$attach"; $baseline3 = "$attDir3\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir3, "$base\srcwb" | Out-Null
    $rwb = $xl.Workbooks.Add()
    $rwb.SaveAs("$base\srcwb\ref.xlsx", 51)
    $rwb.Close($false)
    Set-Content "$attDir3\ref.xlsx" -Value 'OLD-COPY-DIFFERENT'         # 预置同名异内容副本
    $wb3 = $xl.Workbooks.Add()
    $ws3 = $wb3.Worksheets.Item(1)
    $wb3.SaveAs("$wbDir3\main3.xlsx", 51)
    $s12 = [string]$xl.Run('AttachLink_InsertCrossRefCore', "$base\srcwb\ref.xlsx", 'Sheet1', $ws3.Range('H1'))
    ($s12 -split "`n") | ForEach-Object { "    $_" }
    Check (Test-Path "$attDir3\ref(2).xlsx")                            '12: 自动另存 ref(2).xlsx'
    Check ((Get-Content "$attDir3\ref.xlsx") -eq 'OLD-COPY-DIFFERENT')  '12: 原副本未被覆盖'
    $hl12 = $ws3.Cells.Item(1, 8).Hyperlinks.Item(1)
    Check ($hl12.Address -eq "$attach\ref(2).xlsx")                     '12: 链接指向 (2)'
    Check ($hl12.SubAddress -eq "'Sheet1'!A1")                          '12: SubAddress 正确'
    $bl12 = [string]([IO.File]::ReadAllText($baseline3))
    Check ($bl12 -match '(?m)^ref\(2\)\.xlsx\|')                        '12: 基线已登记'
    Check ($s12 -like '*另存*')                                          '12: 汇总含「另存」'
    $wb3.Close($true)

    '--- 13) 哈希去重：同一文件被多条断链命中只真实计算一次 ---'
    $wbDir4 = "$base\proj4"; $attDir4 = "$wbDir4\$attach"; $repDir4 = "$base\repair4"
    New-Item -ItemType Directory -Force -Path $attDir4, $repDir4 | Out-Null
    Set-Content "$srcA\dup2.pdf" -Value 'dup2-content'
    $wb4 = $xl.Workbooks.Add()
    $ws4 = $wb4.Worksheets.Item(1)
    $wb4.SaveAs("$wbDir4\main4.xlsx", 51)
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\dup2.pdf", $ws4.Range('A1'))
    $null = $ws4.Hyperlinks.Add($ws4.Cells.Item(2, 1), ($attach + '\dup2.pdf'), '', 'D2')
    Remove-Item "$attDir4\dup2.pdf" -Force
    Set-Content "$repDir4\dup2.pdf" -Value 'dup2-content'
    $null = $xl.Run('AttachLink_ResetHashCount')
    $s13 = [string]$xl.Run('AttachLink_RepairCore', $wb4, $repDir4)
    ($s13 -split "`n") | ForEach-Object { "    $_" }
    $hashCnt = [int]$xl.Run('AttachLink_GetHashCount')
    Check ($hashCnt -eq 1)                                              "13: 同一文件只真实计算一次（实际 $hashCnt）"
    Check ((Test-Path "$attDir4\dup2.pdf") -and ($ws4.Cells.Item(2, 1).Hyperlinks.Item(1).Address -eq "$attach\dup2.pdf")) '13: 两条断链均已重连'
    $wb4.Close($true)

    '--- 14) O-04：基线损坏 = 失败关闭（插入/修复在任何写入前拒绝） ---'
    $wbDir5 = "$base\proj5"; $attDir5 = "$wbDir5\$attach"; $baseline5 = "$attDir5\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir5 | Out-Null
    Set-Content "$srcA\m1.pdf" -Value 'M1-content'
    $wb5 = $xl.Workbooks.Add()
    $ws5 = $wb5.Worksheets.Item(1)
    $wb5.SaveAs("$wbDir5\main5.xlsx", 51)
    Add-Content $baseline5 -Value 'bad|three|fields' -Encoding UTF8       # 制造损坏基线
    $snap14 = [IO.File]::ReadAllBytes($baseline5)
    $s14 = [string]$xl.Run('AttachLink_InsertCore', "$srcA\m1.pdf", $ws5.Range('A1'))
    "    $($s14 -split "`n" | Select-Object -First 1)"
    Check ($s14 -like '失败：基线文件损坏或无法读取*')                        '14a: 插入被失败关闭'
    Check ($s14 -like '*未插入任何附件*')                                  '14a: 说明本次未插入'
    Check (-not (Test-Path "$attDir5\m1.pdf"))                             '14a: 文件未复制'
    Check ([int]$ws5.Cells.Item(1, 1).Hyperlinks.Count -eq 0)              '14a: 链接未插入'
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes($baseline5)) -eq [Convert]::ToBase64String($snap14)) '14a: 损坏基线未被覆盖'
    $s14b = [string]$xl.Run('AttachLink_RepairCore', $wb5, $srcA)
    "    $($s14b -split "`n" | Select-Object -First 1)"
    Check ($s14b -like '失败：基线文件损坏或无法读取*')                       '14b: 修复被失败关闭'
    $s14c = [string]$xl.Run('AttachLink_InsertFolderCore', $srcA)
    Check ($s14c -like '失败：基线文件损坏或无法读取*')                       '14c: 文件夹引用被失败关闭'
    $s14d = [string]$xl.Run('AttachLink_InsertCrossRefCore', "$srcA\m1.pdf", 'Sheet1', $ws5.Range('A2'))
    Check ($s14d -like '失败：基线文件损坏或无法读取*')                       '14d: 底稿引用被失败关闭'
    $wb5.Close($true)

    '--- 15) O-05：副本与基线不一致时不得复用扩散 → 另存新名并明示 ---'
    $wbDir6 = "$base\proj6"; $attDir6 = "$wbDir6\$attach"; $baseline6 = "$attDir6\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir6 | Out-Null
    Set-Content "$srcA\an1.pdf" -Value 'AN1-original'
    $wb6 = $xl.Workbooks.Add()
    $ws6 = $wb6.Worksheets.Item(1)
    $wb6.SaveAs("$wbDir6\main6.xlsx", 51)
    $null = $xl.Run('AttachLink_InsertCore', "$srcA\an1.pdf", $ws6.Range('A1'))   # 登记基线 v1
    $hashOrig15 = Sha "$attDir6\an1.pdf"
    Set-Content "$attDir6\an1.pdf" -Value 'AN1-TAMPERED-BY-OTHERS'               # 副本被改动（体检会标红）
    Set-Content "$srcA\an1.pdf" -Value 'AN1-TAMPERED-BY-OTHERS'                  # 来源内容与改动后副本一致
    $s15 = [string]$xl.Run('AttachLink_InsertCore', "$srcA\an1.pdf", $ws6.Range('A2'))
    "    $($s15 -split "`n" | Where-Object { $_ -like '*基线不一致*' } | Select-Object -First 1)"
    Check (Test-Path "$attDir6\an1(2).pdf")                                  '15a: 另存 an1(2).pdf（未复用篡改副本）'
    Check ((Sha "$attDir6\an1.pdf") -ne $hashOrig15)                         '15a: 原副本字节未动（仍为篡改内容）'
    Check ((Sha "$attDir6\an1(2).pdf") -eq (Sha "$srcA\an1.pdf"))            '15a: 另存副本内容=来源'
    Check ($ws6.Cells.Item(2, 1).Hyperlinks.Item(1).Address -eq "$attach\an1(2).pdf") '15a: 新链接指向另存副本'
    Check ($s15 -like '*与其内容基线不一致*' -and $s15 -like '*未复用*' -and $s15 -like '*已另存为*') '15a: 汇总明示不一致/未复用/已另存'
    $bl15 = [string]([IO.File]::ReadAllText($baseline6))
    Check ($bl15 -match ([regex]::Escape("an1.pdf|$hashOrig15")))            '15a: 原键保持原哈希（篡改副本持续标红）'
    Check ($bl15 -match '(?m)^an1\(2\)\.pdf\|')                             '15a: 另存副本登记自身基线'
    # 15b) 底稿引用继承同一三方哈希闸
    New-Item -ItemType Directory -Force -Path "$base\srcwb15" | Out-Null
    $rwb15 = $xl.Workbooks.Add()
    $rwb15.SaveAs("$base\srcwb15\an2.xlsx", 51)
    $rwb15.Close($false)
    Set-Content "$attDir6\an2.xlsx" -Value 'X' -Force
    Set-Content "$base\srcwb15\an2.xlsx" -Value 'X' -Force
    $null = $xl.Run('AttachLink_InsertCrossRefCore', "$base\srcwb15\an2.xlsx", 'Sheet1', $ws6.Range('A3'))  # 登记基线（hash X）
    Set-Content "$attDir6\an2.xlsx" -Value 'Y' -Force                        # 副本被改动
    Set-Content "$base\srcwb15\an2.xlsx" -Value 'Y' -Force                   # 来源=改动后内容
    $s15b = [string]$xl.Run('AttachLink_InsertCrossRefCore', "$base\srcwb15\an2.xlsx", 'Sheet1', $ws6.Range('A4'))
    Check (Test-Path "$attDir6\an2(2).xlsx")                                 '15b: 底稿引用另存 an2(2).xlsx'
    Check ((Get-Content "$attDir6\an2.xlsx") -eq 'Y')                        '15b: 原副本未被覆盖'
    Check ($s15b -like '*基线不一致*' -and $s15b -like '*未复用*' -and $s15b -like '*已另存为*') '15b: 底稿引用汇总同样明示'
    $wb6.Close($true)

    '--- 16) O-02：子目录断链不回退借根目录同名键（防张冠李戴） ---'
    $wbDir7 = "$base\proj7"; $attDir7 = "$wbDir7\$attach"; $baseline7 = "$attDir7\_attachlink_baseline.csv"; $repDir16 = "$base\repair16"
    New-Item -ItemType Directory -Force -Path "$attDir7\sub", $repDir16 | Out-Null
    Set-Content "$attDir7\a.pdf" -Value 'ROOT-content'
    Set-Content "$attDir7\sub\a.pdf" -Value 'SUB-original-content'
    $wb7 = $xl.Workbooks.Add()
    $ws7 = $wb7.Worksheets.Item(1)
    $null = $ws7.Hyperlinks.Add($ws7.Cells.Item(1, 1), ($attach + '\a.pdf'), '', 'ROOT')
    $null = $ws7.Hyperlinks.Add($ws7.Cells.Item(2, 1), ($attach + '\sub\a.pdf'), '', 'SUB')
    $wb7.SaveAs("$wbDir7\main7.xlsx", 51)
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb7)                       # 登记 root 与 sub 两条键
    # 手工删除 sub 键，只留根目录键（模拟「子目录文件从未登记」场景）
    $bl16 = Get-Content $baseline7 -Encoding UTF8 | Where-Object { $_ -notlike 'sub\*' }
    Set-Content $baseline7 -Value $bl16 -Encoding UTF8
    Remove-Item "$attDir7\sub\a.pdf" -Force                                  # 子目录断链
    Set-Content "$repDir16\a.pdf" -Value 'ROOT-content'                      # 修复源只有「根目录内容」同名文件
    $s16 = [string]$xl.Run('AttachLink_RepairCore', $wb7, $repDir16)
    ($s16 -split "`n") | ForEach-Object { "    $_" }
    Check (-not (Test-Path "$attDir7\sub\a.pdf"))                            '16: 子目录未被张冠李戴落回'
    Check ($ws7.Cells.Item(2, 1).Hyperlinks.Item(1).Address -eq "$attach\sub\a.pdf") '16: 链接保持断链（地址未动）'
    Check ($s16 -like '*同名但内容不同*')                                     '16: 如实归类「同名但内容不同」'
    $wb7.Close($true)

    '--- 17) V7：文件名含 # → 三类入口复制前拒绝、附件夹零复制 ---'
    $wbDir8 = "$base\proj8"; $attDir8 = "$wbDir8\$attach"
    New-Item -ItemType Directory -Force -Path $attDir8, "$base\srcHash", "$base\foldHash", "$base\srcwbHash" | Out-Null
    Set-Content "$base\srcHash\bad#name.pdf" -Value 'HASH-V7-content'
    Set-Content "$base\foldHash\ok.pdf" -Value 'OK-content'
    Set-Content "$base\foldHash\bad#too.pdf" -Value 'BAD-V7-content'
    Set-Content "$base\srcwbHash\re#f.xlsx" -Value 'dummy-wb'                # 文件系统允许 #；Excel SaveAs 会截断，故直接造占位文件（闸在复制前触发）
    $wb8 = $xl.Workbooks.Add()
    $ws8 = $wb8.Worksheets.Item(1)
    $wb8.SaveAs("$wbDir8\main8.xlsx", 51)
    # 17a 单文件插入
    $s17a = [string]$xl.Run('AttachLink_InsertCore', "$base\srcHash\bad#name.pdf", $ws8.Range('A1'))
    "    $($s17a -split "`n" | Select-Object -First 1)"
    Check ($s17a -like '失败：以下文件名含有 #*' -and $s17a -like '*bad#name.pdf*') '17a: 单文件插入拒绝并列出清单'
    Check (-not (Test-Path "$attDir8\bad#name.pdf"))                          '17a: 附件夹内无复制'
    # 17b 文件夹引用（夹内任一文件名含 # → 整批拒绝）
    $s17b = [string]$xl.Run('AttachLink_InsertFolderCore', "$base\foldHash")
    "    $($s17b -split "`n" | Select-Object -First 1)"
    Check ($s17b -like '失败：以下文件名含有 #*' -and $s17b -like '*bad#too.pdf*') '17b: 文件夹引用整批拒绝并列出清单'
    Check (-not (Test-Path "$attDir8\ok.pdf") -and -not (Test-Path "$attDir8\bad#too.pdf")) '17b: 整批未复制（含正常文件）'
    # 17c 底稿引用（目标底稿文件名含 #）
    $s17c = [string]$xl.Run('AttachLink_InsertCrossRefCore', "$base\srcwbHash\re#f.xlsx", 'Sheet1', $ws8.Range('A2'))
    "    $($s17c -split "`n" | Select-Object -First 1)"
    Check ($s17c -like '失败：以下文件名含有 #*' -and $s17c -like '*re#f.xlsx*') '17c: 底稿引用拒绝并列出文件名'
    Check (-not (Test-Path "$attDir8\re#f.xlsx"))                             '17c: 附件夹内无复制'

    '--- 18) V11b：复制后源大小变化 → 基线以目标实际哈希登记 ---'
    Set-Content "$srcA\grow.pdf" -Value 'GROW-original'
    $null = $xl.Run('AttachLink_TestGrowAfterCopy', 1)
    $s18 = [string]$xl.Run('AttachLink_InsertCore', "$srcA\grow.pdf", $ws8.Range('A3'))
    $null = $xl.Run('AttachLink_TestGrowAfterCopy', 0)
    "    $($s18 -split "`n" | Select-Object -First 1)"
    $hashTgt = Sha "$attDir8\grow.pdf"
    $hashSrcGrown = Sha "$srcA\grow.pdf"
    Check ($hashTgt -ne $hashSrcGrown)                                        '18: 夹具有效（复制后源确实变大）'
    $bl18 = [string]([IO.File]::ReadAllText("$attDir8\_attachlink_baseline.csv"))
    Check ($bl18 -match ('(?m)^grow\.pdf\|' + $hashTgt + '\|'))                '18: 基线登记目标实际哈希（非源哈希）'
    Check ($bl18 -notmatch ('(?m)^grow\.pdf\|' + $hashSrcGrown + '\|'))       '18: 未登记变化后的源哈希'
    $wb8.Close($true)
}
finally {
    New-Item -ItemType File -Force -Path $sentinel | Out-Null
    try { Stop-Process -Id $killer.Id -Force -ErrorAction SilentlyContinue } catch {}
    try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
    Start-Sleep 2
}

if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
