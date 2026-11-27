# =====================================================================
# AttachLink v1.5.0 T14（体检哈希去重/进度/取消）+ T8（保存归一化失败
# 可见）+ V1 补修（安全替换/损坏措辞/子目录断链/逐链接隔离/扩展名护栏）
# + oracle A 修复批（O-06 取消透传/重入防护/N-01 只读基线措辞）
# + 修复批 V（council C4/C5）：V2 受保护体检提示+包装器兜底、V3 版本更新
#   忙碌互斥、V6 正斜杠链接识别、V8a 受保护表插入逐条隔离、V9 损坏基线尾句、
#   V10 网页路径闸、V11a 合法文件名 a..b.pdf
# ---------------------------------------------------------------------
# 覆盖（对应任务书 D 节 + oracle A 修复规格 §三）：
#   1  哈希去重：3 链接同文件 + 2 文件各 1 链接（跨 2 表）→ 体检只算 3 次
#   2  进度：GetProgressCount ≥ 链接数
#   3  取消：BuildIndexCore(wb,2) → 「已取消」、基线字节不变、A1 红字、
#      StatusBar/ScreenUpdating/EnableCancelKey 全部还原
#   4  正常体检后 EnableCancelKey=1、StatusBar=False
#   5  T8：受保护表上绝对路径链接 → 归一化返回 1、体检 E=「正常（■ 绝对
#      路径）」、汇总计数；解除保护后归一化返回 0、提示消失
#   6  NotifyNormalizeFailure 有/无失败记录均无错调用（有记录时状态栏出现提醒）
#   7  补修：a 备份替换失败注入（.bak 被目录占位→False 且原基线不变）
#      b 子目录断链不被平铺同名文件截断
#      c 受保护表下 RepairCore 逐链接隔离（钩子真实失败文案上汇总，p1 失败
#         不计入 p2；工作表顺序显式固定，断言精确计数）
#      d 损坏基线下插入附注含「基线文件损坏」
#      e 按内容查找跳过不同扩展名的同内容文件
#      f O-06 嵌套取消透传：testCancelInFolder 在两层子文件夹扫描中触发 Esc
#         → 「已取消」、基线不变、A1 取消说明、三环境值还原
#      g O-06 重入防护：InsertCore 运行期间再调 InsertCrossRefCore
#         → 「失败：已有操作正在进行」进汇总且经 TestReentrySeen 可读
#      h N-01 只读（未来格式）基线：新增键登记未持久化 → 「本次未写入新登记」，
#         不写「体检将补登记」、不提示「检查写权限」
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File test-progress.ps1
# 期望末行：RESULT: PASS（夹具自动清理；FAIL 保留现场）
# 注意：含中文，必须 UTF-8 with BOM。
# =====================================================================
$ErrorActionPreference = 'Stop'
$attach = [string]([char]0x9644 + [char]0x4EF6)   # 附件
$idxName = $attach + [string]([char]0x6E05) + [string]([char]0x5355)   # 附件清单
$base = Join-Path $env:TEMP 'opencode\progresstest'
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
if (-not (Test-Path $xlam)) { throw "未找到 $xlam，请先运行 build\build.ps1。" }

Add-Type -Namespace WG -Name U32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }
function Sha($p) { (Get-FileHash $p -Algorithm SHA256).Hash.ToLower() }

Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path "$base\src" | Out-Null

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$procId = 0
[void][WG.U32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$procId)
"Excel PID: $procId"
$null = $xl.Workbooks.Open($xlam)     # 第一时间装载 dist 副本（防常驻看门狗抢装）
# v1.6.4：加载项不再带 Workbook_Open，打开不隐式初始化——菜单、快捷键、应用事件都由显式初始化建立
$null = $xl.Run('AttachLink_Initialize')

# 超时兜底：15 分钟后脚本仍未结束（COM 卡死/意外模态）→ 按 PID 结束本实例
$sentinel = "$base\done.flag"
$killer = Start-Process powershell -PassThru -WindowStyle Hidden -ArgumentList '-NoProfile', '-Command', `
    "Start-Sleep -Seconds 900; if (-not (Test-Path '$sentinel')) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }"

# 体检行读取：相对/绝对路径 → E、F 列文本
function Get-IndexRows($wb) {
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)
    $sheet = $wb.Worksheets.Item($idxName)
    $out = @{}
    $row = 4
    while ($row -le 80) {
        $c3 = [string]$sheet.Cells.Item($row, 3).Value2
        if ($c3) {
            $out[$c3] = @{ E = [string]$sheet.Cells.Item($row, 5).Value2; F = [string]$sheet.Cells.Item($row, 6).Value2 }
        }
        $row++
    }
    $out
}

try {
    '--- 1) 哈希去重：5 链接 3 文件跨 2 表，体检只真实计算 3 次 ---'
    $wbDir = "$base\proj"; $attDir = "$wbDir\$attach"; $baseline = "$attDir\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir | Out-Null
    Set-Content "$base\src\fa.pdf" -Value 'FA-content'
    Set-Content "$base\src\fb.pdf" -Value 'FB-content'
    Set-Content "$base\src\fc.pdf" -Value 'FC-content'
    $wb = $xl.Workbooks.Add()
    $ws1 = $wb.Worksheets.Item(1)
    $wb.SaveAs("$wbDir\main.xlsx", 51)
    $ws2 = $wb.Worksheets.Add()
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\fa.pdf", $ws1.Range('A1'))
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\fb.pdf", $ws1.Range('A3'))
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\fc.pdf", $ws1.Range('A4'))
    # 再补两条指向 fa.pdf 的链接（共 3 条同文件链接，跨 2 个工作表）
    $null = $ws1.Hyperlinks.Add($ws1.Cells.Item(2, 1), ($attach + '\fa.pdf'), '', 'FA2')
    $null = $ws2.Hyperlinks.Add($ws2.Cells.Item(1, 1), ($attach + '\fa.pdf'), '', 'FA3')
    $null = $xl.Run('AttachLink_ResetHashCount')
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)
    $hashCnt = [int]$xl.Run('AttachLink_GetHashCount')
    Check ($hashCnt -eq 3)                                              "1: 5 条链接 3 个文件，体检只真实计算 3 次（实际 $hashCnt）"

    '--- 2) 进度：GetProgressCount ≥ 链接数 ---'
    $null = $xl.Run('AttachLink_ResetProgressCount')
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)
    $progCnt = [int]$xl.Run('AttachLink_GetProgressCount')
    Check ($progCnt -ge 5)                                              "2: 进度刷新次数 ≥ 链接数 5（实际 $progCnt）"

    '--- 3) 取消：cancelAfter=2 → 已取消 + 基线不变 + 环境还原 ---'
    # 先制造一个「未登记」的链接（extra.pdf 无基线键），取消时若有登记会被丢弃
    Set-Content "$attDir\extra.pdf" -Value 'EXTRA-content'
    $null = $ws1.Hyperlinks.Add($ws1.Cells.Item(5, 1), ($attach + '\extra.pdf'), '', 'EX')
    $snap3 = [IO.File]::ReadAllBytes($baseline)
    $r3 = [string]$xl.Run('AttachLink_BuildIndexCore', $wb, 2)
    "    $r3"
    Check ($r3 -like '已取消*')                                          '3: 返回以「已取消」开头'
    Check ($r3 -like '*2/6*')                                            '3: 已处理 2/6'
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes($baseline)) -eq [Convert]::ToBase64String($snap3)) '3: 基线文件字节不变（登记被丢弃）'
    $a1 = [string]$wb.Worksheets.Item($idxName).Range('A1').Value2
    Check ($a1 -like '*体检已取消*' -and $a1 -like '*2/6*')               '3: 清单 A1 红字说明已处理 N/M'
    Check ($xl.StatusBar -eq $false)                                     '3: StatusBar 已还原 False'
    Check ($xl.ScreenUpdating -eq $true)                                 '3: ScreenUpdating 已还原 True'
    Check ([int]$xl.EnableCancelKey -eq 1)                               '3: EnableCancelKey 已还原 1（xlInterrupt）'

    '--- 4) 正常体检后环境还原 ---'
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb)
    Check ([int]$xl.EnableCancelKey -eq 1)                               '4: 正常体检后 EnableCancelKey=1'
    Check ($xl.StatusBar -eq $false)                                     '4: 正常体检后 StatusBar=False'
    $bl4 = [string]([IO.File]::ReadAllText($baseline))
    Check ($bl4 -match '(?m)^extra\.pdf\|')                              '4: 正常体检补登记 extra.pdf（取消时丢弃的登记可重建）'
    $wb.Close($true)

    '--- 5) T8：受保护表上的绝对路径链接 ---'
    $wbDir2 = "$base\proj2"; $attDir2 = "$wbDir2\$attach"
    New-Item -ItemType Directory -Force -Path $attDir2 | Out-Null
    Set-Content "$base\src\fd.pdf" -Value 'FD-content'
    $wb2 = $xl.Workbooks.Add()
    $wsB = $wb2.Worksheets.Item(1)
    $wb2.SaveAs("$wbDir2\main2.xlsx", 51)
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\fd.pdf", $wsB.Range('A1'))
    # 改成指向本附件夹的绝对路径链接，再保护工作表
    $wsB.Cells.Item(1, 1).Hyperlinks.Item(1).Address = "$attDir2\fd.pdf"
    $wsB.Protect()
    # 本机 Excel 实测：受保护工作表上 Hyperlink.Address 赋值并不抛错（与 Hyperlinks.Add 不同），
    # 仅靠保护无法让归一化失败——用测试钩子注入 1 次改指失败（与 BuildIndexCore 的 cancelAfter 同类；
    # 保护保留以贴近真实场景；钩子被本次归一化消耗后归零，不影响后续组）
    $null = $xl.Run('AttachLink_TestRepointFail', 1)
    $n1 = [int]$xl.Run('NormalizeAttachmentLinks', $wb2)
    Check ($n1 -eq 1)                                                    '5: 受保护表归一化返回失败数 1'
    $rows5 = Get-IndexRows $wb2
    Check ($rows5["$attDir2\fd.pdf"].E -eq '正常（■ 绝对路径）')           '5: 体检 E 列「正常（■ 绝对路径）」'
    $sum5 = [string]$xl.Run('AttachLink_BuildIndexCore', $wb2)
    Check ($sum5 -like '*绝对路径未归一化 1*')                            '5: 汇总含「绝对路径未归一化 1」'
    Check ($sum5 -like '*解除保护后重新保存*')                            '5: 汇总附处置说明'

    '--- 6) NotifyNormalizeFailure：有/无记录均无错调用 ---'
    $xl.StatusBar = $false
    $null = $xl.Run('AttachLink_NotifyNormalizeFailure')                 # 无参：汇总全部记录
    Check ([string]$xl.StatusBar -like '*未能归一化为相对路径*')           '6: 有记录时状态栏出现提醒'
    $null = $xl.Run('AttachLink_NotifyNormalizeFailure', $wb2)           # 带 WB：只看该工作簿
    Check ([string]$xl.StatusBar -like '*未能归一化为相对路径*')           '6: 带 WB 同样提醒'
    # 解除保护 → 归一化成功 → 记录清除、地址还原、体检提示消失
    $wsB.Unprotect()
    $n2 = [int]$xl.Run('NormalizeAttachmentLinks', $wb2)
    Check ($n2 -eq 0)                                                    '6: 解除保护后归一化返回 0'
    Check ($wsB.Cells.Item(1, 1).Hyperlinks.Item(1).Address -eq "$attach\fd.pdf") '6: 地址已归一化为相对路径'
    $xl.StatusBar = $false
    $null = $xl.Run('AttachLink_NotifyNormalizeFailure', $wb2)           # 记录已清除：无错且不置状态栏
    Check ($xl.StatusBar -eq $false)                                     '6: 无记录时不改状态栏'
    $rows6 = Get-IndexRows $wb2
    Check ($rows6["$attach\fd.pdf"].E -eq '正常')                         '6: 体检 E 列恢复「正常」'
    $sum6 = [string]$xl.Run('AttachLink_BuildIndexCore', $wb2)
    Check ($sum6 -notlike '*绝对路径未归一化*')                           '6: 汇总不再含绝对路径计数（为 0 时逐字不变）'
    $wb2.Close($true)

    '--- 7a) 补修：备份替换失败注入（.bak 被同名目录占位） ---'
    $wbDir3 = "$base\proj3"; $attDir3 = "$wbDir3\$attach"; $baseline3 = "$attDir3\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir3 | Out-Null
    Set-Content "$base\src\g1.pdf" -Value 'G1-content'
    Set-Content "$base\src\g2.pdf" -Value 'G2-content'
    $wb3 = $xl.Workbooks.Add()
    $wsC = $wb3.Worksheets.Item(1)
    $wb3.SaveAs("$wbDir3\main3.xlsx", 51)
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\g1.pdf", $wsC.Range('A1'))
    $snap7a = [IO.File]::ReadAllBytes($baseline3)
    New-Item -ItemType Directory -Force -Path "$baseline3.bak" | Out-Null   # 占位：MoveFile 正式→.bak 必失败
    $s7a = [string]$xl.Run('AttachLink_InsertCore', "$base\src\g2.pdf", $wsC.Range('A2'))
    "    $($s7a -split "`n" | Select-Object -Last 1)"
    Check ($s7a -like '*未持久化*' -and $s7a -like '*备份原基线失败*')      '7a: 汇总如实报「备份原基线失败」'
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes($baseline3)) -eq [Convert]::ToBase64String($snap7a)) '7a: 正式基线字节不变（未两失）'
    Remove-Item "$baseline3.bak" -Recurse -Force
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb3)
    $bl7a = [string]([IO.File]::ReadAllText($baseline3))
    Check ($bl7a -match '(?m)^g2\.pdf\|')                                '7a: 清除占位后体检补登记 g2（可恢复）'
    $wb3.Close($true)

    '--- 7b) 补修：子目录断链不被平铺同名文件截断 ---'
    $wbDir4 = "$base\proj4"; $attDir4 = "$wbDir4\$attach"; $repDir4 = "$base\repair4"
    New-Item -ItemType Directory -Force -Path "$attDir4\sub", $repDir4 | Out-Null
    Set-Content "$attDir4\a.pdf" -Value 'FLAT-content'
    Set-Content "$attDir4\sub\a.pdf" -Value 'SUB-content'
    $wb4 = $xl.Workbooks.Add()
    $wsD = $wb4.Worksheets.Item(1)
    $wb4.SaveAs("$wbDir4\main4.xlsx", 51)
    $null = $wsD.Hyperlinks.Add($wsD.Cells.Item(1, 1), ($attach + '\a.pdf'), '', 'FLAT')
    $null = $wsD.Hyperlinks.Add($wsD.Cells.Item(2, 1), ($attach + '\sub\a.pdf'), '', 'SUB')
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb4)                   # 登记两条基线
    Remove-Item "$attDir4\sub\a.pdf" -Force                              # 子目录断链（平铺同名仍在）
    Set-Content "$repDir4\a.pdf" -Value 'SUB-content'                    # 修复源：同内容
    $s7b = [string]$xl.Run('AttachLink_RepairCore', $wb4, $repDir4)
    ($s7b -split "`n") | ForEach-Object { "    $_" }
    Check ((Test-Path "$attDir4\sub\a.pdf") -and ((Get-Content "$attDir4\sub\a.pdf") -eq 'SUB-content')) '7b: 子目录断链以原内容落回 sub\a.pdf'
    Check ((Get-Content "$attDir4\a.pdf") -eq 'FLAT-content')            '7b: 平铺同名文件未被触碰'
    Check ($wsD.Cells.Item(2, 1).Hyperlinks.Item(1).Address -eq "$attach\sub\a.pdf") '7b: 链接地址保留子目录'
    Check ($s7b -like '*内容一致（同名）*')                               '7b: 按同名同内容重连（未截断为冲突）'
    $wb4.Close($true)

    '--- 7c) 补修：受保护工作表下 RepairCore 逐链接隔离（钩子真实文案 + 精确计数） ---'
    $wbDir5 = "$base\proj5"; $attDir5 = "$wbDir5\$attach"; $baseline5 = "$attDir5\_attachlink_baseline.csv"; $repDir5 = "$base\repair5"
    New-Item -ItemType Directory -Force -Path $attDir5, $repDir5 | Out-Null
    Set-Content "$base\src\p1.pdf" -Value 'P1-content'
    Set-Content "$base\src\p2.pdf" -Value 'P2-content'
    $wb5 = $xl.Workbooks.Add()
    $wsE1 = $wb5.Worksheets.Item(1)
    $wb5.SaveAs("$wbDir5\main5.xlsx", 51)
    # 显式把工作表 2 排在表 1 之后：RepairCore 按标签顺序处理，p1（表1 A1）必须先于
    # p2（表2 B2）被处理，唯一一次注入才确定落在 p1 上
    $wsE2 = $wb5.Worksheets.Add([System.Reflection.Missing]::Value, $wsE1)
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\p1.pdf", $wsE1.Range('A1'))
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\p2.pdf", $wsE2.Range('B2'))
    Remove-Item "$attDir5\p1.pdf", "$attDir5\p2.pdf" -Force
    Set-Content "$repDir5\p1.pdf" -Value 'P1-content'
    Set-Content "$repDir5\p2.pdf" -Value 'P2-content'
    $wsE1.Protect()                                                       # p1 所在表受保护
    # 同 5)：保护表实际不阻断改指，用钩子让 p1 的改指失败（p1 先于 p2 处理，钩子=1 恰好被 p1 消耗）
    $null = $xl.Run('AttachLink_TestRepointFail', 1)
    $s7c = [string]$xl.Run('AttachLink_RepairCore', $wb5, $repDir5)
    ($s7c -split "`n") | ForEach-Object { "    $_" }
    Check ($wsE2.Cells.Item(2, 2).Hyperlinks.Item(1).Address -eq "$attach\p2.pdf") '7c: 未保护表的 p2 正常重连'
    Check ((Test-Path "$attDir5\p2.pdf"))                                '7c: p2 文件已落回'
    Check ($s7c -like '*重连·内容一致（同名）：1*')                          '7c: 恰 1 个同名同内容重连（p2）'
    Check ($s7c -like '*测试注入的改指失败*')                               '7c: p1 失败文案含钩子真实原因（repErr 已上汇总）'
    Check ($s7c -like '*重连失败：1*')                                     '7c: 恰 1 个重连失败（p1，逐链接隔离未扩散）'
    Check ($s7c -notlike '*重连失败：2*')                                  '7c: 失败未蔓延到 p2'
    Check (Test-Path "$attDir5\p1.pdf")                                  '7c: p1 文件已落回（仅链接未改指，如实写明）'
    Check ($wsE1.Cells.Item(1, 1).Hyperlinks.Count -ge 1)                '7c: p1 链接保留（改指失败不删链接）'
    $bl7c = [string]([IO.File]::ReadAllText($baseline5))
    Check ($bl7c -match '(?m)^p1\.pdf\|' -and $bl7c -match '(?m)^p2\.pdf\|') '7c: 部分失败下重连登记仍整批落盘'
    $wsE1.Unprotect()
    $wb5.Close($true)

    '--- 7d) O-04：损坏基线下插入失败关闭（不再走到落盘附注路径） ---'
    $wbDir6 = "$base\proj6"; $attDir6 = "$wbDir6\$attach"; $baseline6 = "$attDir6\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir6 | Out-Null
    Set-Content "$base\src\q1.pdf" -Value 'Q1-content'
    Set-Content "$base\src\q2.pdf" -Value 'Q2-content'
    $wb6 = $xl.Workbooks.Add()
    $wsF = $wb6.Worksheets.Item(1)
    $wb6.SaveAs("$wbDir6\main6.xlsx", 51)
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\q1.pdf", $wsF.Range('A1'))
    Add-Content $baseline6 -Value 'bad|three|fields' -Encoding UTF8       # 制造损坏基线
    $snap7d = [IO.File]::ReadAllBytes($baseline6)
    $s7d = [string]$xl.Run('AttachLink_InsertCore', "$base\src\q2.pdf", $wsF.Range('A2'))
    "    $($s7d -split "`n" | Select-Object -First 1)"
    Check ($s7d -like '失败：基线文件损坏或无法读取*')                        '7d: 失败关闭并带损坏原因'
    Check ($s7d -like '*未插入任何附件*')                                  '7d: 说明本次未插入任何附件'
    Check (-not (Test-Path "$attDir6\q2.pdf"))                             '7d: 文件未复制（任何写入之前拒绝）'
    Check ($s7d -notlike '*体检将补登记*')                                 '7d: 不写「体检将补登记」'
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes($baseline6)) -eq [Convert]::ToBase64String($snap7d)) '7d: 损坏基线未被覆盖'
    $wb6.Close($true)

    '--- 7e) 补修：按内容查找跳过不同扩展名的同内容文件 ---'
    $wbDir7 = "$base\proj7"; $attDir7 = "$wbDir7\$attach"; $repDir7 = "$base\repair7"
    New-Item -ItemType Directory -Force -Path $attDir7, $repDir7 | Out-Null
    Set-Content "$base\src\h1.pdf" -Value 'H1-content'
    $wb7 = $xl.Workbooks.Add()
    $wsG = $wb7.Worksheets.Item(1)
    $wb7.SaveAs("$wbDir7\main7.xlsx", 51)
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\h1.pdf", $wsG.Range('A1'))
    Remove-Item "$attDir7\h1.pdf" -Force
    Set-Content "$repDir7\h1.txt" -Value 'H1-content'                     # 同内容、不同扩展名
    $s7e = [string]$xl.Run('AttachLink_RepairCore', $wb7, $repDir7)
    ($s7e -split "`n") | ForEach-Object { "    $_" }
    Check (-not (Test-Path "$attDir7\h1.pdf"))                           '7e: 不同扩展名不自动找回（未落回）'
    Check (-not (Test-Path "$attDir7\h1.txt"))                           '7e: 候选 txt 未被复制'
    Check ($s7e -like '*未找到*' -and $s7e -like '*扩展名*')               '7e: 汇总「未找到」并提示扩展名护栏'
    $wb7.Close($true)

    '--- 7f) O-06：嵌套子文件夹扫描中取消（钩子）→ 透传为「已取消」，环境与基线还原 ---'
    $wbDir8 = "$base\proj8"; $attDir8 = "$wbDir8\$attach"; $baseline8 = "$attDir8\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path "$attDir8\fd\sub1\sub2" | Out-Null
    Set-Content "$attDir8\fd\root.pdf" -Value 'R-content'
    Set-Content "$attDir8\fd\sub1\s1.pdf" -Value 'S1-content'
    Set-Content "$attDir8\fd\sub1\sub2\s2.pdf" -Value 'S2-content'
    $wb8 = $xl.Workbooks.Add()
    $wsH = $wb8.Worksheets.Item(1)
    $null = $wsH.Hyperlinks.Add($wsH.Cells.Item(1, 1), ($attach + '\fd'), '', 'FOLDER')
    $wb8.SaveAs("$wbDir8\main8.xlsx", 51)
    $su8 = [int]$xl.ScreenUpdating
    $ck8 = [int]$xl.EnableCancelKey
    # 钩子=1：CheckFolderBaseline 首次进入（fd 根目录）即按用户 Esc 处理——
    # 若 O-06 透传链路失效，取消会被父级 Resume Next 吞掉、体检假完成
    $null = $xl.Run('AttachLink_TestCancelInFolder', 1)
    $r8 = [string]$xl.Run('AttachLink_BuildIndexCore', $wb8)
    "    $r8"
    Check ($r8 -like '已取消*')                                          '7f: 嵌套扫描中取消透传为「已取消」'
    Check ($r8 -notlike '*文件一致*')                                     '7f: 取消未被吞掉假完成（无一致结论）'
    Check (-not (Test-Path $baseline8))                                  '7f: 基线未生成（登记全部丢弃）'
    $a8 = [string]$wb8.Worksheets.Item($idxName).Range('A1').Value2
    Check ($a8 -like '*体检已取消*')                                      '7f: 清单 A1 只有取消说明'
    Check ($xl.StatusBar -eq $false)                                     '7f: StatusBar 还原 False'
    Check ([int]$xl.ScreenUpdating -eq $su8)                             '7f: ScreenUpdating 还原为进入前原值'
    Check ([int]$xl.EnableCancelKey -eq $ck8)                            '7f: EnableCancelKey 还原为进入前原值'
    $wb8.Close($true)

    '--- 7g) O-06：重入防护——InsertCore 运行期间再调 InsertCrossRefCore 必被拦截 ---'
    $wbDir9 = "$base\proj9"; $attDir9 = "$wbDir9\$attach"
    New-Item -ItemType Directory -Force -Path $attDir9 | Out-Null
    Set-Content "$base\src\r1.pdf" -Value 'R1-content'
    $wb9 = $xl.Workbooks.Add()
    $wsI = $wb9.Worksheets.Item(1)
    $wb9.SaveAs("$wbDir9\main9.xlsx", 51)
    $null = $xl.Run('AttachLink_TestReentryProbe')
    $s9 = [string]$xl.Run('AttachLink_InsertCore', "$base\src\r1.pdf", $wsI.Range('A1'))
    $seen9 = [string]$xl.Run('AttachLink_TestReentrySeen')   # 一次性读取（读取即清空）
    Check ($s9 -like '*重入探测*' -and $s9 -like '*已有操作正在进行*')        '7g: 重入拦截结果进本次汇总'
    Check ($seen9 -like '*已有操作正在进行*')                             '7g: 拦截结果经 TestReentrySeen 可读'
    Check (Test-Path "$attDir9\r1.pdf")                                  '7g: 拦截后本次插入自身仍正常完成'
    $after9 = [string]$xl.Run('AttachLink_InsertCrossRefCore', "$base\no-such.xlsx", 'Sheet1', $wsI.Range('A2'))
    Check ($after9 -like '*目标底稿不存在*')                              '7g: 拦截一次性消费后忙标志已复位（后续调用正常入参校验）'
    $wb9.Close($true)

    '--- 7h) N-01：只读（未来格式）基线下新增键登记未持久化的措辞 ---'
    $wbDir10 = "$base\proj10"; $attDir10 = "$wbDir10\$attach"; $baseline10 = "$attDir10\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir10 | Out-Null
    Set-Content "$base\src\q1.pdf" -Value 'Q1-content'
    Set-Content "$base\src\q2.pdf" -Value 'Q2-content'
    $wb10 = $xl.Workbooks.Add()
    $wsJ = $wb10.Worksheets.Item(1)
    $wb10.SaveAs("$wbDir10\main10.xlsx", 51)
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\q1.pdf", $wsJ.Range('A1'))
    # 把基线改写成未来格式（>4 字段 → 本版只读）：q1 键仍可比对，新登记无法持久化
    $h10 = Sha "$attDir10\q1.pdf"
    Set-Content $baseline10 -Value "q1.pdf|$h10|1|2026-01-01 00:00|2026-01-01 00:00|future" -Encoding UTF8
    $snap10 = [IO.File]::ReadAllBytes($baseline10)
    $s10 = [string]$xl.Run('AttachLink_InsertCore', "$base\src\q2.pdf", $wsJ.Range('A2'))
    "    $($s10 -split "`n" | Select-Object -Last 1)"
    Check ($s10 -like '*未持久化*' -and $s10 -like '*本次未写入新登记*')      '7h: 附注「本次未写入新登记」'
    Check ($s10 -notlike '*体检将补登记*')                                '7h: 只读基线不写「体检将补登记」'
    Check ($s10 -notlike '*检查写权限*')                                  '7h: 只读基线不提示「检查写权限」'
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes($baseline10)) -eq [Convert]::ToBase64String($snap10)) '7h: 只读基线字节不变（未被覆盖）'
    $wb10.Close($true)

    '--- 8) V2：受保护工作簿体检的明确提示 + 忙碌标志出错兜底 ---'
    $wbDir11 = "$base\proj11"; $attDir11 = "$wbDir11\$attach"; $baseline11 = "$attDir11\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir11 | Out-Null
    Set-Content "$base\src\w1.pdf" -Value 'W1-content'
    $wb11 = $xl.Workbooks.Add()
    $wsK = $wb11.Worksheets.Item(1)
    $wb11.SaveAs("$wbDir11\main11.xlsx", 51)
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\w1.pdf", $wsK.Range('A1'))
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb11)                    # 建清单表
    $sheet11 = $wb11.Worksheets.Item($idxName)
    $dataRow11 = 4                                                          # 首个数据行（相对路径列非空）
    while ($dataRow11 -le 60 -and -not $sheet11.Cells.Item($dataRow11, 3).Value2) { $dataRow11++ }
    $sheet11.Cells.Item($dataRow11, 3).Value2 = 'MARK-KEEP'                # 内容标记（验证体检不改它）
    $sheet11.Protect()
    $s8a = [string]$xl.Run('AttachLink_BuildIndexCore', $wb11)
    "    $s8a"
    Check ($s8a -like '失败：*' -and $s8a -like '*受保护*')                  'V2-①: 受保护清单表 → 「失败：…受保护…」'
    Check ([string]$sheet11.Cells.Item($dataRow11, 3).Value2 -eq 'MARK-KEEP') 'V2-①: 清单表内容未被改动'
    $sheet11.Unprotect()
    $wbDir12 = "$base\proj12"; $attDir12 = "$wbDir12\$attach"
    New-Item -ItemType Directory -Force -Path $attDir12 | Out-Null
    $wb12 = $xl.Workbooks.Add()
    $wsL = $wb12.Worksheets.Item(1)
    $wb12.SaveAs("$wbDir12\main12.xlsx", 51)
    $wb12.Protect('', $true)                                                 # 无清单表 + 结构受保护（参数1=密码、参数2=Structure）
    $cnt12 = $wb12.Worksheets.Count
    $s8b = [string]$xl.Run('AttachLink_BuildIndexCore', $wb12)
    "    $s8b"
    Check ($s8b -like '失败：*' -and $s8b -like '*受保护*')                  'V2-②: 结构受保护且无清单表 → 「失败：…受保护…」'
    Check ($wb12.Worksheets.Count -eq $cnt12)                               'V2-②: 未新建工作表'
    $wb12.Unprotect('')
    [void]$wb11.Activate()                                                   # COM 操作可能切走活动工作簿，重锚定后再测
    # ③ 两种拒绝之后忙碌标志必须已复位（否则全部功能被锁死，重启 Excel 才恢复）
    $s8c = [string]$xl.Run('AttachLink_InsertCore', "$base\src\w1.pdf", $wsK.Range('A2'))
    Check ($s8c -notlike '*已有操作正在进行*')                               'V2-③: 体检被拒后插入不被忙标志锁死'
    $s8d = [string]$xl.Run('AttachLink_RepairCore', $wb11, "$base\src")
    Check ($s8d -notlike '*已有操作正在进行*')                               'V2-③: 体检被拒后修复不被忙标志锁死'
    # ④ 注入未处理错误 → 包装器兜底：返回「失败：<过程名> - <原因>」且忙标志复位
    [void]$wb11.Activate()                                                   # 修复 Core 可能切走活动工作簿，重锚定
    $null = $xl.Run('AttachLink_TestThrowAtInsert', 1)
    $s8e = [string]$xl.Run('AttachLink_InsertCore', "$base\src\w1.pdf", $wsK.Range('A3'))
    "    $s8e"
    Check ($s8e -like '失败：AttachLink_InsertCore*' -and $s8e -like '*测试注入的未处理错误*') 'V2-④: Inner 未处理错误被包装器兜住（含过程名与原因）'
    $s8f = [string]$xl.Run('AttachLink_InsertCore', "$base\src\w1.pdf", $wsK.Range('A4'))
    "    $($s8f -split "`n" | Select-Object -First 1)"
    Check ($s8f -notlike '失败：*' -and $s8f -notlike '*已有操作正在进行*')      'V2-④: 兜底后忙标志已复位（再次插入正常完成）'
    $wb11.Close($true); $wb12.Close($true)

    '--- 9) V8a：受保护表插入——副本登记基线、原因写明、解除保护后重试直接复用 ---'
    $wbDir13 = "$base\proj13"; $attDir13 = "$wbDir13\$attach"; $baseline13 = "$attDir13\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir13 | Out-Null
    Set-Content "$base\src\v1.pdf" -Value 'V1-content'
    $wb13 = $xl.Workbooks.Add()
    $wsM = $wb13.Worksheets.Item(1)
    $wb13.SaveAs("$wbDir13\main13.xlsx", 51)
    $wsM.Protect()
    $s9 = [string]$xl.Run('AttachLink_InsertCore', "$base\src\v1.pdf", $wsM.Range('A1'))
    ($s9 -split "`n") | ForEach-Object { "    $_" }
    Check ($s9 -like '*失败 1 个文件*')                                     'V8a: 该条按文件计入失败（不再整批中断、不再无原因）'
    Check ($s9 -like '*受保护*' -and $s9 -like '*已复制到附件夹并登记基线*' -and $s9 -like '*直接复用*') 'V8a: 原因写明受保护+已登记+可复用重试'
    Check (Test-Path "$attDir13\v1.pdf")                                   'V8a: 副本已复制'
    $bl9 = [string]([IO.File]::ReadAllText($baseline13))
    Check ($bl9 -match '(?m)^v1\.pdf\|')                                    'V8a: 副本已登记基线'
    Check ($wsM.Cells.Item(1, 1).Hyperlinks.Count -eq 0)                    'V8a: 链接未插入（失败不消耗单元格）'
    $wsM.Unprotect()
    $s9b = [string]$xl.Run('AttachLink_InsertCore', "$base\src\v1.pdf", $wsM.Range('A1'))
    Check ($s9b -like '*复用*' -and $wsM.Cells.Item(1, 1).Hyperlinks.Count -eq 1) 'V8a: 解除保护后重试直接复用并插入链接'
    $wb13.Close($true)

    '--- 10) V9：基线损坏时汇总尾句如实（不再「内容基线全部一致」自相矛盾） ---'
    $wbDir14 = "$base\proj14"; $attDir14 = "$wbDir14\$attach"; $baseline14 = "$attDir14\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir14 | Out-Null
    Set-Content "$attDir14\e1.pdf" -Value 'E1-content'
    $wb14 = $xl.Workbooks.Add()
    $wsN = $wb14.Worksheets.Item(1)
    $wb14.SaveAs("$wbDir14\main14.xlsx", 51)
    $null = $wsN.Hyperlinks.Add($wsN.Cells.Item(1, 1), ($attach + '\e1.pdf'), '', 'E1')
    $null = $xl.Run('AttachLink_BuildIndexCore', $wb14)                     # 正常登记
    Add-Content $baseline14 -Value 'bad|three|fields' -Encoding UTF8        # 制造损坏
    $sV9 = [string]$xl.Run('AttachLink_BuildIndexCore', $wb14)
    "    $($sV9 -split "`n" | Select-Object -Last 1)"
    Check ($sV9 -like '*基线文件损坏，本次未做内容比对、未写入任何登记*')        'V9: 损坏基线尾句如实说明'
    Check ($sV9 -notlike '*内容基线全部一致*')                               'V9: 不再以「全部一致」收尾（不矛盾）'
    $wb14.Close($true)

    '--- 11) V6：附件/a.pdf 正斜杠识别 + V11a：合法文件名 a..b.pdf 不再误判 ---'
    $wbDir15 = "$base\proj15"; $attDir15 = "$wbDir15\$attach"
    New-Item -ItemType Directory -Force -Path "$attDir15\sub" | Out-Null
    Set-Content "$attDir15\a.pdf" -Value 'SLASH-A'
    Set-Content "$attDir15\sub\b.pdf" -Value 'SLASH-B'
    Set-Content "$attDir15\a..b.pdf" -Value 'DOTDOT'
    $wb15 = $xl.Workbooks.Add()
    $wsO = $wb15.Worksheets.Item(1)
    $wb15.SaveAs("$wbDir15\main15.xlsx", 51)
    $null = $wsO.Hyperlinks.Add($wsO.Cells.Item(1, 1), ($attach + '/a.pdf'), '', 'SLA')
    $null = $wsO.Hyperlinks.Add($wsO.Cells.Item(2, 1), ($attach + '/sub/b.pdf'), '', 'SLB')
    $null = $wsO.Hyperlinks.Add($wsO.Cells.Item(3, 1), ($attach + '\a..b.pdf'), '', 'DDT')
    $sumV6 = [string]$xl.Run('AttachLink_BuildIndexCore', $wb15)
    "    $($sumV6 -split "`n" | Select-Object -First 1)"
    Check ($sumV6 -notlike '*未发现任何附件链接*')                           'V6: 正斜杠链接被识别（不再整体漏检）'
    $rowsV6 = Get-IndexRows $wb15
    $leafE = { param($leaf) $k = @($rowsV6.Keys | Where-Object { $_ -like "*$leaf" }); if ($k) { $rowsV6[$k[0]].E } else { '<无此行>' } }
    Check ((& $leafE 'a.pdf') -eq '正常')                                    'V6: 附件/a.pdf 体检状态「正常」'
    Check ((& $leafE 'b.pdf') -eq '正常')                                    'V6: 附件/sub/b.pdf 体检状态「正常」'
    Check ((& $leafE 'a..b.pdf') -eq '正常')                                 'V11a: 文件名含连续两点不再误判「链接路径异常」'
    $wb15.Close($true)

    '--- 12) V10：网页路径底稿（SharePoint/OneDrive）:// 闸 ---'
    $gateTxt = [string]$xl.Run('AttachLink_TestWebPathGate', 'https://contoso.sharepoint.com/sites/audit/Shared Documents/main.xlsx')
    Check ($gateTxt -like '*网页路径*' -and $gateTxt -like '*另存到本地*')      'V10: 闸文案命中（https 路径）'
    $gateLocal = [string]$xl.Run('AttachLink_TestWebPathGate', 'C:\audit\main.xlsx')
    Check ($gateLocal -eq '')                                                'V10: 本地路径放行（""）'
    # 参数化 Core 用假对象（wb.Path=https://...）直测；ActiveWorkbook 型 Core 经源码接线断言
    $srcBas = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\src\AttachLink.bas'), [Text.Encoding]::UTF8)
    $fakeOk = $false
    try {
        Add-Type -TypeDefinition 'public class FakeWebWb { public string Path { get; set; } }' -ReferencedAssemblies System.dll
        $fakeWb = New-Object FakeWebWb
        $fakeWb.Path = 'https://contoso.sharepoint.com/sites/audit/Shared Documents/main.xlsx'
        $s10a = [string]$xl.Run('AttachLink_BuildIndexCore', $fakeWb)
        "    BuildIndex: $s10a"
        Check ($s10a -like '失败：*网页路径*')                                'V10: 体检 Core 拒绝网页路径底稿'
        $s10b = [string]$xl.Run('AttachLink_DeleteAttachmentCore', $fakeWb, 'A1|Sheet1', 'C:\x\c.pdf')
        "    Delete: $s10b"
        Check ($s10b -like '失败：*网页路径*')                                'V10: 删除 Core 拒绝网页路径底稿'
        $fakeOk = $true
    } catch {
        "    假对象编组失败（$($_.Exception.Message)），改源码接线断言"
    }
    if (-not $fakeOk) {
        Check ($srcBas -match '(?s)Function BuildIndexInner.*?IsWebPathWorkbook') 'V10: 体检 Core 网页闸已接线（源码）'
        Check ($srcBas -match '(?s)Function DeleteAttachmentCoreInner.*?IsWebPathWorkbook') 'V10: 删除 Core 网页闸已接线（源码）'
    }
    Check ($srcBas -match '(?s)Function InsertCrossRefCoreInner.*?IsWebPathWorkbook') 'V10: 底稿引用 Core 网页闸已接线'
    Check ($srcBas -match '(?s)Function UpdateVersionCoreInner.*?IsWebPathWorkbook') 'V10: 版本更新 Core 网页闸已接线'
    $wbDir16 = "$base\proj16"; $attDir16 = "$wbDir16\$attach"; $baseline16 = "$attDir16\_attachlink_baseline.csv"
    New-Item -ItemType Directory -Force -Path $attDir16 | Out-Null
    Set-Content "$base\src\n1.pdf" -Value 'N1-content'
    Set-Content "$base\src\n2.pdf" -Value 'N2-content'
    $wb16 = $xl.Workbooks.Add()
    $wsP = $wb16.Worksheets.Item(1)
    $wb16.SaveAs("$wbDir16\main16.xlsx", 51)
    $hN1 = Sha "$base\src\n1.pdf"
    # V3：版本更新参与忙碌互斥——借重入探测在 InsertCore 运行期间直调 UpdateVersionCore，
    # 必须被忙标志拦截（单线程 COM 无法在脚本层真并发，探测钩子是唯一真实入口）
    $null = $xl.Run('AttachLink_InsertCore', "$base\src\n1.pdf", $wsP.Range('A1'))  # 先归档 n1 供后续版本更新
    Set-Content "$base\src\n1v2.pdf" -Value 'N1-NEW-content'
    $snapN1 = [IO.File]::ReadAllBytes("$attDir16\n1.pdf")
    $null = $xl.Run('AttachLink_TestReentryProbe')
    $sV3 = [string]$xl.Run('AttachLink_InsertCore', "$base\src\n2.pdf", $wsP.Range('A2'))
    $seenV3 = [string]$xl.Run('AttachLink_TestReentrySeen')
    Check ($seenV3 -like '*版本更新探测*已有操作正在进行*')                  'V3: Core 运行期间版本更新被忙碌互斥拒绝'
    Check ([Convert]::ToBase64String([IO.File]::ReadAllBytes("$attDir16\n1.pdf")) -eq [Convert]::ToBase64String($snapN1)) 'V3: 被拒后文件未变'
    # 空闲时版本更新正常执行（包装器接入后功能不回归）
    $sV3a = [string]$xl.Run('AttachLink_UpdateVersionCore', "$attDir16\n1.pdf", "$base\src\n1v2.pdf")
    "    $($sV3a -split "`n" | Select-Object -First 1)"
    Check ($sV3a -like '*版本更新完成*')                                     'V3: 空闲时版本更新正常执行'
    Check ((Get-Content "$attDir16\n1.pdf" -Raw) -like '*N1-NEW-content*')    'V3: 版本更新确实生效（文件已更替）'
    $sV3b = [string]$xl.Run('AttachLink_UpdateVersionCore', "$attDir16\n1.pdf", "$base\no-such.pdf")
    Check ($sV3b -like '失败：*')                                            'V3: 忙标志复位后入参校验正常（拒绝不存在的源文件）'
    $wb16.Close($true)
}
finally {
    New-Item -ItemType File -Force -Path $sentinel | Out-Null
    try { Stop-Process -Id $killer.Id -Force -ErrorAction SilentlyContinue } catch {}
    try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
    Start-Sleep 2
}

if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
