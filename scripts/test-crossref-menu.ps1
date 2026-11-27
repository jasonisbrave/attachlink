# =====================================================================
# AttachLink v1.6.0 R5/R6 回归：右键菜单结构（附件管理子菜单）+
# 附件夹前置提示去重（按入口 × 文件夹）+ 前置检查位置静态守卫
# ---------------------------------------------------------------------
# 覆盖（对应开发交接.md 第十节 R5/R6/10.3）：
#   1  一级菜单顺序：插入附件链接（可多选）/ 插入文件夹引用 / 插入底稿引用 /
#      归档迁移（批量）/ 附件管理（弹出式 Type 10）
#   2  附件管理子项顺序与各 OnAction（清单/体检 → 断链修复 → 版本更新 →
#      重命名链接显示名 → 附件删除）
#   3  AttachLink_EnsureMenu 自愈调用后结构不变；无「归档」/ 一级「重命名」残留
#   4  提示记录：同入口二次不提示、异入口各自提示一次、路径大小写不敏感、
#      Reset 后恢复提示
#   5  静态守卫：AttachLink_InsertCrossRef 过程体内 EnsureAttachDirPrompted
#      恰好一次且位于 Application.FileDialog 之前（选区校验之后）；
#      另两个入口各恰好一次；旧单变量记录 lastNotifiedAttachDir 已移除
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File test-crossref-menu.ps1
# 期望末行：RESULT: PASS
# 注意：含中文，必须 UTF-8 with BOM。
# =====================================================================
$ErrorActionPreference = 'Stop'
$base = Join-Path $env:TEMP 'opencode\crossrefmenutest'
$xlam = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
$basSrc = Join-Path $PSScriptRoot '..\src\AttachLink.bas'
if (-not (Test-Path $xlam)) { throw "未找到 $xlam，请先运行 build\build.ps1。" }

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }

Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $base | Out-Null

# ---- 静态守卫（先于 Excel 起实例，守卫失败也不拉 Excel） ----
'--- 0) 静态守卫：EnsureAttachDirPrompted 位置与次数 ---'
$src = [IO.File]::ReadAllText($basSrc, [Text.Encoding]::UTF8)
function Get-ProcBody($name) {
    $m = [regex]::Match($src, '(?ms)^Public Sub ' + $name + '\(\).*?^End Sub')
    if (-not $m.Success) { throw "源码中未找到过程 $name" }
    return $m.Value
}
$bodyInsert = Get-ProcBody 'AttachLink_Insert'
$bodyFolder = Get-ProcBody 'AttachLink_InsertFolder'
$bodyCross  = Get-ProcBody 'AttachLink_InsertCrossRef'
Check ([regex]::Matches($bodyCross, 'EnsureAttachDirPrompted').Count -eq 1)  '底稿引用：EnsureAttachDirPrompted 恰好一次'
Check ([regex]::Matches($bodyInsert, 'EnsureAttachDirPrompted').Count -eq 1) '插入附件：EnsureAttachDirPrompted 恰好一次'
Check ([regex]::Matches($bodyFolder, 'EnsureAttachDirPrompted').Count -eq 1) '插入文件夹：EnsureAttachDirPrompted 恰好一次'
Check ($bodyCross.IndexOf('EnsureAttachDirPrompted') -ge 0 -and `
       $bodyCross.IndexOf('EnsureAttachDirPrompted') -lt $bodyCross.IndexOf('Application.FileDialog')) `
      '底稿引用：检查位于 Application.FileDialog 之前'
Check ($bodyCross.IndexOf('EnsureAttachDirPrompted') -gt $bodyCross.IndexOf('TypeName(Selection)')) `
      '底稿引用：检查位于选区校验之后'
Check ($bodyCross.Contains('"crossref"'))               '底稿引用：入口标识为 crossref'
Check (-not $src.Contains('lastNotifiedAttachDir'))     '全模块：旧单变量记录 lastNotifiedAttachDir 已移除'

# ---- Excel 实例（隐藏）。打开加载项之前没有 EXCEL7 窗口，常驻看门狗不会附着；打开之后加载项自带一个隐藏
#      文档窗口，看门狗会看到并按「隐藏窗口」跳过——下面还第一时间装载 dist 副本，它更不会抢装 ----
Add-Type -Namespace WM -Name U32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$procId = 0
[void][WM.U32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$procId)
"Excel PID: $procId"
$null = $xl.Workbooks.Open($xlam)     # 第一时间装载 dist 副本（防常驻看门狗抢装）
# v1.6.4：加载项不再带 Workbook_Open，打开不隐式初始化——菜单、快捷键、应用事件都由显式初始化建立
$null = $xl.Run('AttachLink_Initialize')
$xl.DisplayAlerts = $false             # 打开加载项会触发 Workbook_Open，Excel 在其结束时把 DisplayAlerts 复位为 True——按红线设在打开之后（开发交接 3.5）；v1.6.4 起加载项自身不再复位；保留此写法对别的加载项仍稳妥

# 超时兜底：10 分钟后脚本仍未结束（COM 卡死）→ 按 PID 结束本实例
$sentinel = "$base\done.flag"
$killer = Start-Process powershell -PassThru -WindowStyle Hidden -ArgumentList '-NoProfile', '-Command', `
    "Start-Sleep -Seconds 600; if (-not (Test-Path '$sentinel')) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }"

# 菜单快照：取 Cell 命令栏中全部 MENU_TAG 控件（值已取好再释放 COM 对象）
function Get-MenuSnapshot {
    $cb = $xl.CommandBars.Item('Cell')
    $top = @()
    for ($i = 1; $i -le $cb.Controls.Count; $i++) {
        $c = $cb.Controls.Item($i)
        if ($c.Tag -eq 'AttachLinkCtx') {
            # 临时控件的 OnAction 读回时带加载项前缀（AttachLink.xlam!XXX），取 ! 后的过程名
            $item = @{ Caption = [string]$c.Caption; Type = [int]$c.Type; OnAction = ([string]$c.OnAction -split '!')[-1]; Sub = @() }
            if ([int]$c.Type -eq 10) {
                for ($j = 1; $j -le $c.Controls.Count; $j++) {
                    $s = $c.Controls.Item($j)
                    $item.Sub += @{ Caption = [string]$s.Caption; OnAction = ([string]$s.OnAction -split '!')[-1]; Tag = [string]$s.Tag }
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($s)
                }
            }
            $top += $item
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($c)
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cb)
    return $top
}

try {
    '--- 1) 右键菜单一级结构（R6） ---'
    $menu = Get-MenuSnapshot
    $caps = @($menu | ForEach-Object { $_.Caption })
    "    一级: $($caps -join ' | ')"
    $expTop = @('插入附件链接（可多选）', '插入文件夹引用', '插入底稿引用', '归档迁移（批量）', '附件管理')
    Check ($menu.Count -eq 5) '一级恰好 5 个 AttachLinkCtx 控件'
    $orderOk = ($menu.Count -eq 5)
    if ($orderOk) { for ($i = 0; $i -lt 5; $i++) { if ($menu[$i].Caption -ne $expTop[$i]) { $orderOk = $false } } }
    Check $orderOk '一级标题顺序与 R6 定稿一致'
    $expOA = [ordered]@{
        '插入附件链接（可多选）' = 'AttachLink_Insert'
        '插入文件夹引用'         = 'AttachLink_InsertFolder'
        '插入底稿引用'           = 'AttachLink_InsertCrossRef'
        '归档迁移（批量）'       = 'AttachLink_ArchiveMigrate'
    }
    foreach ($k in $expOA.Keys) {
        $hits = @($menu | Where-Object { $_.Caption -eq $k })
        Check ($hits.Count -eq 1 -and $hits[0].OnAction -eq $expOA[$k]) "一级 OnAction：$k → $($expOA[$k])"
    }
    Check (($menu | Where-Object { $_.Caption -eq '归档' }).Count -eq 0)           '无旧「归档」子菜单残留'
    Check (($menu | Where-Object { $_.Caption -eq '重命名链接显示名' }).Count -eq 0) '一级无「重命名链接显示名」残留'
    Check (($menu | Where-Object { $_.Caption -eq '附件清单 / 体检' }).Count -eq 0)  '一级无「附件清单 / 体检」残留'

    '--- 2) 「附件管理」弹出式与子项 ---'
    $pop = @($menu | Where-Object { $_.Caption -eq '附件管理' })[0]
    Check ($pop -ne $null -and [int]$pop.Type -eq 10) '第 5 项「附件管理」为弹出式（Type 10）'
    $expSub = @(
        @{ c = '附件清单 / 体检';       oa = 'AttachLink_BuildIndex' },
        @{ c = '附件断链修复';           oa = 'AttachLink_Repair' },
        @{ c = '附件版本更新';           oa = 'AttachLink_UpdateVersion' },
        @{ c = '重命名链接显示名';       oa = 'AttachLink_RenameLink' },
        @{ c = '附件删除（删链接+文件）'; oa = 'AttachLink_DeleteAttachment' }
    )
    if ($pop -ne $null) {
        $subCaps = @($pop.Sub | ForEach-Object { $_.Caption })
        "    子项: $($subCaps -join ' | ')"
        Check ($pop.Sub.Count -eq 5) '子菜单恰好 5 项'
        $subOk = ($pop.Sub.Count -eq 5)
        if ($subOk) {
            for ($i = 0; $i -lt 5; $i++) {
                if ($pop.Sub[$i].Caption -ne $expSub[$i].c -or $pop.Sub[$i].OnAction -ne $expSub[$i].oa) { $subOk = $false }
            }
        }
        Check $subOk '子项标题顺序与 OnAction 全部一致'
        Check (($pop.Sub | Where-Object { $_.Tag -ne 'AttachLinkCtx' }).Count -eq 0) '子项全部带 MENU_TAG'
    } else {
        Check $false '「附件管理」弹出项存在（前置断言已记录）'
    }

    '--- 3) AttachLink_EnsureMenu 自愈调用后结构不变 ---'
    $null = $xl.Run('AttachLink_EnsureMenu')
    $menu2 = Get-MenuSnapshot
    Check ($menu2.Count -eq 5) 'EnsureMenu 后一级仍 5 项'
    $same = ($menu2.Count -eq 5 -and $menu.Count -eq 5)
    if ($same) {
        for ($i = 0; $i -lt 5; $i++) {
            if ($menu2[$i].Caption -ne $menu[$i].Caption -or $menu2[$i].OnAction -ne $menu[$i].OnAction) { $same = $false }
        }
    }
    Check $same 'EnsureMenu 后一级标题与 OnAction 不变'
    $pop2 = @($menu2 | Where-Object { $_.Caption -eq '附件管理' })[0]
    if ($pop2 -ne $null -and $pop -ne $null) {
        $sameSub = ($pop2.Sub.Count -eq $pop.Sub.Count)
        if ($sameSub) { for ($i = 0; $i -lt $pop.Sub.Count; $i++) { if ($pop2.Sub[$i].Caption -ne $pop.Sub[$i].Caption) { $sameSub = $false } } }
        Check $sameSub 'EnsureMenu 后子菜单不变'
    } else {
        Check $false 'EnsureMenu 后「附件管理」仍在'
    }

    '--- 4) 前置检查提示记录（入口 × 文件夹，会话级） ---'
    $attach = [string]([char]0x9644 + [char]0x4EF6)   # 附件
    $dirA = "$base\projA\$attach"
    $dirB = "$base\projB\$attach"
    New-Item -ItemType Directory -Force -Path $dirA, $dirB | Out-Null
    $null = $xl.Run('AttachLink_FolderNoticeReset')
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA, 'insert') -eq $true)  '初始：insert×A 需要提示'
    $null = $xl.Run('AttachLink_FolderNoticeMark', $dirA, 'insert')
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA, 'insert') -eq $false) '同入口二次：不再提示'
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA, 'crossref') -eq $true) '异入口 crossref×A：各自提示一次'
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA, 'folder') -eq $true)   '异入口 folder×A：各自提示一次'
    $null = $xl.Run('AttachLink_FolderNoticeMark', $dirA, 'folder')
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA, 'folder') -eq $false)  'folder×A 标记后不提示'
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA, 'crossref') -eq $true) 'folder 标记不影响 crossref'
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirB, 'insert') -eq $true)   '同入口不同文件夹：需要提示'
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA.ToUpperInvariant(), 'insert') -eq $false) '路径大小写不敏感（大写变体不重复提示）'
    $null = $xl.Run('AttachLink_FolderNoticeReset')
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA, 'insert') -eq $true)  'Reset 后 insert×A 恢复提示'
    Check ($xl.Run('AttachLink_FolderNoticeNeeded', $dirA, 'folder') -eq $true)   'Reset 后 folder×A 恢复提示'
} finally {
    New-Item -ItemType File -Force -Path $sentinel | Out-Null
    try { Stop-Process -Id $killer.Id -Force -ErrorAction SilentlyContinue } catch {}
    try { foreach ($w in @($xl.Workbooks)) { $w.Saved = $true }; $xl.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
    Start-Sleep 2
}

if ($ok) { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue; 'RESULT: PASS（夹具已清理）' } else { 'RESULT: FAIL（夹具保留在 ' + $base + '）' }
