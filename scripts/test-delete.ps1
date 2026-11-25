# =====================================================================
# AttachLink 删除附件功能测试（直调 AttachLink_DeleteAttachmentCore）
# ---------------------------------------------------------------------
# 覆盖：文件/文件夹/断链删除 → 链接删除 + 样式清除 + 边框填充保留
#       → 特殊表名（含 | 与单引号）重定位 → 未选项隔离
#       → 基线键清理（含文件夹后代键）→ 失败真实性（锁定文件 / 保护工作表）
#       → 共享文件二次确认（V4：范围外链接 → 需确认 / 确认后删除 / 全范围内不确认）
# 用法：powershell -ExecutionPolicy Bypass -File test-delete.ps1
# 注意：含中文，必须 UTF-8 with BOM。
# =====================================================================
$ErrorActionPreference = 'Stop'
$attach = [string]([char]0x9644 + [char]0x4EF6)
$base = Join-Path $env:TEMP 'opencode\deletetest'
$wbDir = "$base\proj"
$attDir = "$wbDir\$attach"
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
$baseline = "$attDir\_attachlink_baseline.csv"

Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path "$attDir\f1\sub" | Out-Null
Set-Content "$attDir\a.pdf" -Value 'A'
Set-Content "$attDir\b.pdf" -Value 'B'
Set-Content "$attDir\c.pdf" -Value 'C'
Set-Content "$attDir\f1\in1.pdf" -Value 'I1'
Set-Content "$attDir\f1\sub\in2.pdf" -Value 'I2'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }

try {
    $null = $xl.Workbooks.Open($xlam)
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Sheet1'
    $ws2 = $wb.Worksheets.Add()
    $ws2.Name = "S|2 'x"          # 表名含 | 与单引号：验证 Split(...,2) 重定位
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(1, 1), "$attach\a.pdf", '', 'AA')
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(2, 1), "$attach\b.pdf", '', 'BB')    # 未选中，不删
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(3, 1), "$attach\f1", '', 'FF')       # 文件夹
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(4, 1), "$attach\gone.pdf", '', 'GG') # 断链
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(1, 1), "$attach\c.pdf", '', 'CC')
    # A1 预置超链接样式 + 边框/填充（验证样式清除与格式保留）
    $ws.Cells.Item(1, 1).Font.Underline = 2                     # xlUnderlineStyleSingle
    $ws.Cells.Item(1, 1).Font.Color = 255
    $ws.Cells.Item(1, 1).Font.Bold = $true
    $ws.Cells.Item(1, 1).Borders.Item(7).LineStyle = 1          # xlEdgeLeft xlContinuous
    $ws.Cells.Item(1, 1).Borders.Item(7).Color = 16711680
    $ws.Cells.Item(1, 1).Interior.Color = 65535
    $wb.SaveAs("$wbDir\main.xlsx", 51)

    '--- 1) 基线登记（BuildIndexCore）---'
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)
    Check (Test-Path $baseline) '基线文件生成'
    $baseTxt = [string]([IO.File]::ReadAllText($baseline))
    Check ($baseTxt -match ('(?m)^' + [regex]::Escape('f1\in1.pdf') + '\|')) '后代键 f1\in1.pdf 已登记'
    Check ($baseTxt -match ('(?m)^' + [regex]::Escape('f1\sub\in2.pdf') + '\|')) '后代键 f1\sub\in2.pdf 已登记'
    Check ($baseTxt -match ('(?m)^' + [regex]::Escape('b.pdf') + '\|')) 'b.pdf 已登记'

    '--- 2) 直调 DeleteAttachmentCore：a.pdf + f1 文件夹 + 断链 gone + c.pdf（特殊表名）---'
    $lf = [string][char]10
    $linksJ = "A1|Sheet1" + $lf + "A3|Sheet1" + $lf + "A4|Sheet1" + $lf + "A1|S|2 'x"
    $pathsJ = "$attDir\a.pdf" + $lf + "$attDir\f1" + $lf + "$attDir\gone.pdf" + $lf + "$attDir\c.pdf"
    $summary = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, $linksJ, $pathsJ)
    "    summary = $summary"
    Check (-not (Test-Path "$attDir\a.pdf")) 'a.pdf 已删除'
    Check (-not (Test-Path "$attDir\f1")) '文件夹 f1 及内容已删除'
    Check (-not (Test-Path "$attDir\c.pdf")) 'c.pdf 已删除（特殊表名重定位成功）'
    Check (Test-Path "$attDir\b.pdf") '未选中的 b.pdf 保留'
    $a1 = $ws.Cells.Item(1, 1)
    Check ($a1.Hyperlinks.Count -eq 0) 'A1 链接已删除'
    Check ("$($a1.Value2)" -eq '') 'A1 文本已清空'
    Check ([int]$a1.Font.Underline -eq -4142) 'A1 下划线已移除'
    Check ([int]$a1.Font.ColorIndex -eq -4105) 'A1 字体色恢复自动'
    Check (-not $a1.Font.Bold) 'A1 粗体已移除'
    Check ([int]$a1.Borders.Item(7).LineStyle -eq 1 -and [int]$a1.Borders.Item(7).Color -eq 16711680) 'A1 边框保留'
    Check ([int]$a1.Interior.Color -eq 65535) 'A1 填充保留'
    Check ($ws.Cells.Item(2, 1).Hyperlinks.Count -eq 1 -and "$($ws.Cells.Item(2, 1).Value2)" -ne '') 'A2 链接与文本未受影响'
    Check ($ws.Cells.Item(3, 1).Hyperlinks.Count -eq 0) 'A3 文件夹链接已删除'
    Check ($ws2.Cells.Item(1, 1).Hyperlinks.Count -eq 0) '特殊表名页 A1 链接已删除'
    $baseTxt2 = [string]([IO.File]::ReadAllText($baseline))
    Check ($baseTxt2 -notmatch '(?m)^a\.pdf\|') '基线键 a.pdf 已清'
    Check ($baseTxt2 -notmatch '(?m)^f1\\') '基线键 f1 全部后代键已清'
    Check ($baseTxt2 -notmatch '(?m)^c\.pdf\|') '基线键 c.pdf 已清'
    Check ($baseTxt2 -match ('(?m)^' + [regex]::Escape('b.pdf') + '\|')) '基线键 b.pdf 保留'
    Check ($summary -like '*已删除 3 个文件/文件夹、4 个链接*' -and $summary -like '*1 个链接目标本就不存在*' -and $summary -notlike '*部分失败*') '汇总如实（3 文件/4 链接/1 断链，无部分失败）'

    '--- 3) 失败真实性：锁定文件 ---'
    Set-Content "$attDir\lock.pdf" -Value 'L'
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(5, 1), "$attach\lock.pdf", '', 'LK')
    $fs = [IO.File]::Open("$attDir\lock.pdf", 'Open', 'Read', 'None')   # 独占锁定
    try {
        $s3 = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "A5|Sheet1", "$attDir\lock.pdf")
        "    summary = $s3"
        Check ($s3 -like '*1 个文件/文件夹删除失败*' -and $s3 -like '*部分失败*') '锁定文件：delFail 如实 + 部分失败'
    } finally { $fs.Close() }

    '--- 4) 失败真实性：保护工作表（V8b：链接删不掉 → 文件保留，先删链接后删文件） ---'
    Set-Content "$attDir\p.pdf" -Value 'P'
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(6, 1), "$attach\p.pdf", '', 'PP')
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)                        # 补登记 p.pdf 基线（受保护之前）
    $ws.Protect()
    $s4 = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "A6|Sheet1", "$attDir\p.pdf")
    $ws.Unprotect()
    "    summary = $s4"
    Check ($s4 -like '*1 个链接删除失败*' -and $s4 -like '*部分失败*') '保护工作表：linkDelFail 如实 + 部分失败'
    Check (Test-Path "$attDir\p.pdf") 'V8b: p.pdf 文件保留（链接未删成不动文件）'
    Check ($s4 -like '*文件已保留*') 'V8b: 汇总明示「链接未能删除，文件已保留」'
    Check ($ws.Cells.Item(6, 1).Hyperlinks.Count -eq 1) '被保护单元格链接保留（如实失败而非假成功）'
    # 确认后同参数带 confirmShared 重试不适用本场景（无共享链接），直接验证受保护时基线键也保留
    $bt4 = [string]([IO.File]::ReadAllText($baseline))
    Check ($bt4 -match ('(?m)^' + [regex]::Escape('p.pdf') + '\|')) 'V8b: 链接未删成 → 基线键保留'

    '--- 5) Core 安全契约（oracle 复核 #1）：外部路径 / 附件夹根 / 首合法后越界 → 零删除 ---'
    Set-Content "$attDir\z1.pdf" -Value 'Z1'
    Set-Content "$base\outside.pdf" -Value 'OUT'                    # 附件夹外的目标
    $s5a = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "B1|Sheet1", "$base\outside.pdf")
    "    s5a = $s5a"
    Check ($s5a -like '*失败：*越界*') '外部路径被拒绝（预检拦截）'
    Check (Test-Path "$base\outside.pdf") '外部文件未被删除'
    $s5b = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "B1|Sheet1", $attDir)
    Check ($s5b -like '*失败：*') '附件夹根本身被拒绝'
    $s5c = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "B1|Sheet1" + $lf + "B2|Sheet1", "$attDir\z1.pdf" + $lf + "$base\outside.pdf")
    Check ($s5c -like '*失败：*越界*') '首项合法、后项越界：整体拒绝'
    Check (Test-Path "$attDir\z1.pdf") '首项合法目标也未被删除（零删除契约）'

    '--- 6) 基线严格加载（oracle 复核 #3）：锁定 → 如实失败且不回写 ---'
    Set-Content "$attDir\bl.pdf" -Value 'BL'
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(7, 1), "$attach\bl.pdf", '', 'BL')
    $fs2 = [IO.File]::Open($baseline, 'Open', 'Read', 'None')      # 独占基线文件
    try {
        $s6 = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "A7|Sheet1", "$attDir\bl.pdf")
        "    s6 = $s6"
        Check ($s6 -like '*部分失败*' -and $s6 -like '*基线读取失败*') '基线锁定：如实报基线读取失败'
        Check (-not (Test-Path "$attDir\bl.pdf")) 'bl.pdf 文件本体仍正常删除'
    } finally { $fs2.Close() }
    $btxt = [string]([IO.File]::ReadAllText($baseline))
    Check ($btxt -match ('(?m)^' + [regex]::Escape('b.pdf') + '\|')) '基线未被部分回写（b.pdf 键仍在）'

    '--- 7) 畸形基线（oracle 二轮 #3）：解析失败 → 如实报错且不回写 ---'
    Set-Content "$attDir\bad.pdf" -Value 'BAD'
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(8, 1), "$attach\bad.pdf", '', 'BD')
    $before7 = [string]([IO.File]::ReadAllText($baseline))
    Add-Content $baseline -Value 'garbage-line-without-separator' -Encoding UTF8
    $s7 = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "A8|Sheet1", "$attDir\bad.pdf")
    "    s7 = $s7"
    Check ($s7 -like '*部分失败*' -and $s7 -like '*基线读取失败*') '畸形基线：如实报读取失败'
    Check (-not (Test-Path "$attDir\bad.pdf")) 'bad.pdf 文件本体仍正常删除'
    $after7 = [string]([IO.File]::ReadAllText($baseline))
    Check ($after7.Contains('garbage-line-without-separator')) '畸形行原样保留（未被静默丢弃）'
    Check ($after7 -match ('(?m)^' + [regex]::Escape('b.pdf') + '\|')) '原基线键未被覆盖（严格模式不回写部分字典）'

    '--- 8) V4 共享文件二次确认：同一文件被范围外链接引用 ---'
    Set-Content "$attDir\s.pdf" -Value 'S'
    Set-Content "$attDir\t.pdf" -Value 'T'
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(10, 1), "$attach\s.pdf", '', 'S1')    # 本次选中
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(2, 2), "$attach\s.pdf", '', 'S2')   # 范围外，另一张表
    $null = $ws.Hyperlinks.Add($ws.Cells.Item(11, 1), "$attach\t.pdf", '', 'T1')    # 边界：两处都在范围内
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(3, 2), "$attach\t.pdf", '', 'T2')
    # 第 7 组故意留下了畸形基线行（严格模式不会覆盖损坏基线），先重建干净基线
    Remove-Item $baseline -Force -ErrorAction SilentlyContinue
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)
    $btxt8 = [string]([IO.File]::ReadAllText($baseline))
    Check ($btxt8 -match '(?m)^s\.pdf\|') 'V4: s.pdf 基线键已登记'
    Check ($btxt8 -match '(?m)^t\.pdf\|') 'V4: t.pdf 基线键已登记'

    # 8a) 未带 confirmShared：只选 A10，s.pdf 仍被另一表 B2 引用 → 不删除、返回需确认
    $fsoObj = New-Object -ComObject Scripting.FileSystemObject
    $s8a = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "A10|Sheet1", "$attDir\s.pdf", $fsoObj)
    "    s8a = $s8a"
    Check ($s8a.StartsWith('需确认：') -and $s8a -like '*1 处*') 'V4: 共享链接未确认 → 返回「需确认：」且计 1 处'
    Check ($s8a -like "*$($ws2.Name)!B2*") 'V4: 需确认文本点名范围外链接位置（表名!地址）'
    Check ($s8a -like "*需确认：要删除的文件还被其它 1 处链接引用*" -and $s8a -like '*请点「取消」*') 'V4: 需确认文案说明后果与取消方式'
    Check (Test-Path "$attDir\s.pdf") 'V4: 未确认时 s.pdf 文件保留'
    Check ($ws.Cells.Item(10, 1).Hyperlinks.Count -eq 1 -and $ws2.Cells.Item(2, 2).Hyperlinks.Count -eq 1) 'V4: 未确认时选中处与范围外链接均保留'
    $btxt8a = [string]([IO.File]::ReadAllText($baseline))
    Check ($btxt8a -match '(?m)^s\.pdf\|') 'V4: 未确认时基线键 s.pdf 保留'

    # 8b) 同参数带 confirmShared = True：照常删除，范围外链接保留成断链
    $addrB2 = [string]$ws2.Cells.Item(2, 2).Hyperlinks.Item(1).Address
    $s8b = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "A10|Sheet1", "$attDir\s.pdf", $fsoObj, $true)
    "    s8b = $s8b"
    Check (-not $s8b.StartsWith('需确认：')) 'V4: 确认后不再返回需确认'
    Check (-not (Test-Path "$attDir\s.pdf")) 'V4: 确认后 s.pdf 文件已删除'
    Check ($ws.Cells.Item(10, 1).Hyperlinks.Count -eq 0) 'V4: 确认后选中处链接已删除'
    Check ($ws2.Cells.Item(2, 2).Hyperlinks.Count -eq 1 -and [string]$ws2.Cells.Item(2, 2).Hyperlinks.Item(1).Address -eq $addrB2) 'V4: 范围外链接保留且地址不变'
    $btxt8b = [string]([IO.File]::ReadAllText($baseline))
    Check ($btxt8b -notmatch '(?m)^s\.pdf\|') 'V4: 确认删除后基线键 s.pdf 已清'
    Check ($btxt8b -match ('(?m)^' + [regex]::Escape('b.pdf') + '\|')) 'V4: 其余基线键未受影响（b.pdf 仍在）'

    # 8c) 边界：删除范围已包含全部指向该文件的链接 → 无需确认，直接删除
    $s8c = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $wb, "A11|Sheet1" + $lf + "B3|" + $ws2.Name, "$attDir\t.pdf" + $lf + "$attDir\t.pdf", $fsoObj)
    "    s8c = $s8c"
    Check ($s8c -notlike '*需确认*') 'V4: 全部链接均在范围内 → 不要求确认'
    Check (-not (Test-Path "$attDir\t.pdf")) 'V4: t.pdf 文件已删除'
    Check ($ws.Cells.Item(11, 1).Hyperlinks.Count -eq 0 -and $ws2.Cells.Item(3, 2).Hyperlinks.Count -eq 0) 'V4: t.pdf 两处链接均已删除'
    Check ($s8c -like '*只删除一次*' -and $s8c -notlike '*本就不存在*') 'V4: 多处链接指向同一文件 → 汇总说明「文件只删除一次」，不误报「目标本就不存在」'

    $wb.Close($false)
}
finally {
    if ($null -ne $xl) {
        $xl.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
}

if ($ok) { 'RESULT: PASS（删除附件全部断言通过）'; exit 0 } else { 'RESULT: FAIL'; exit 1 }
