# =====================================================================
# AttachLink 初始化契约回归（v1.6.4）
# ---------------------------------------------------------------------
# 对象：src\AttachLink.ThisWorkbook.vba、src\AttachLink.bas 与 dist\AttachLink.xlam
#      （dist 必须是这份源码构建出来的：版本号从源码读出，再与产物对账）
# 契约（开发交接 3.3 第 15 条）：
#   1. 加载项没有 Workbook_Open，也不处理「打开工作簿」事件；应用事件只有右键与保存前两条
#   2. Workbooks.Open 打开加载项不隐式初始化：菜单不在位，宿主的 DisplayAlerts / ScreenUpdating 不被复位
#   3. 显式 AttachLink_Initialize 之后：菜单在位、SelfTest 为真、版本可读、应用事件已绑定
#   4. 以上与装载时的事件开关无关；此后新建、打开别的工作簿，宿主两项设置不再被改回 True
#   5. 关闭全部工作簿再开一本：菜单仍在；被外部删除后，右键自愈用的 EnsureMenu 能把它补回
# 段落：
#   A. 静态断言（读源码，不启动 Excel）
#   B. 隐藏实例，装载时事件开启
#   C. 隐藏实例，装载与初始化时事件关闭，之后打开事件
# 真实的右键接线与快捷键无法经对象模型触发，由可见实例上的验证覆盖（维护者操作）。
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test-init-contract.ps1
# 期望末行 RESULT: PASS；退出码 0 = 通过，1 = 有失败
# 红线：只结束自己拉起的实例；只删除带 AttachLinkCtx 标签的控件（产品自建的控件是临时的，不进命令栏缓存）。
# 编码：含中文，必须保存为 UTF-8 with BOM（WORKLOG 根因 #11）。
# =====================================================================
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$repo   = Split-Path -Parent $PSScriptRoot
$srcTw  = Join-Path $repo 'src\AttachLink.ThisWorkbook.vba'
$srcBas = Join-Path $repo 'src\AttachLink.bas'
$xlam   = Join-Path $repo 'dist\AttachLink.xlam'
$attach = [string]([char]0x9644 + [char]0x4EF6)
$base   = Join-Path $env:TEMP ('opencode\initcontract_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $base | Out-Null

$script:ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; Write-Host "FAIL: $label" } else { Write-Host "pass: $label" } }
function Release-Com($obj) {
    if ($null -eq $obj) { return }
    try { while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) -gt 0) { } } catch {}
}
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class AttachLinkInitPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid); }'

function Test-Menu($app) {
    # 右键菜单 Tag 是否在位。集合与条目逐个释放，不留给 GC
    $found = $false
    $bars = $null; $bar = $null; $ctrls = $null
    try {
        $bars = $app.CommandBars
        $bar = $bars.Item('Cell')
        $ctrls = $bar.Controls
        $n = [int]$ctrls.Count
        for ($i = 1; $i -le $n; $i++) {
            $ctrl = $null
            try { $ctrl = $ctrls.Item($i); if ([string]$ctrl.Tag -eq 'AttachLinkCtx') { $found = $true } } finally { Release-Com $ctrl }
            if ($found) { break }
        }
    } finally { Release-Com $ctrls; Release-Com $bar; Release-Com $bars }
    return $found
}
function Remove-MenuTagControls($app) {
    # 删除右键菜单上所有带 AttachLinkCtx 标签的控件；集合与条目逐个释放
    $bars = $null; $bar = $null; $ctrls = $null
    try {
        $bars = $app.CommandBars
        $bar = $bars.Item('Cell')
        $ctrls = $bar.Controls
        for ($i = [int]$ctrls.Count; $i -ge 1; $i--) {
            $ctrl = $null
            try { $ctrl = $ctrls.Item($i); if ([string]$ctrl.Tag -eq 'AttachLinkCtx') { $ctrl.Delete() } } finally { Release-Com $ctrl }
        }
    } finally { Release-Com $ctrls; Release-Com $bar; Release-Com $bars }
}
function Read-HostUiState($app) {
    # 宿主的 DisplayAlerts / ScreenUpdating；读不到返回 $null（调用方按「读不到」处理，不能当成 True 或 False）
    for ($k = 0; $k -lt 8; $k++) {
        try {
            $alertsValue = $app.DisplayAlerts
            $screenValue = $app.ScreenUpdating
            if ($null -ne $alertsValue -and $null -ne $screenValue) { return @{ alerts = [bool]$alertsValue; screen = [bool]$screenValue } }
        } catch {}
        Start-Sleep -Milliseconds 150
    }
    return $null
}
function Wait-Exit([uint32]$ProcId, [int]$Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        if (-not (Get-Process -Id $ProcId -ErrorAction SilentlyContinue)) { return $true }
        Start-Sleep -Seconds 1
    }
    return $false
}
function Get-ProcBody([string]$Text, [string]$ProcName) {
    # 取过程体（从「Sub|Function <名>(」那一行之后，到最近的 End Sub|End Function 之前）。找不到返回 $null
    $lines = @($Text -split "`r?`n")
    $start = -1
    $head = '^\s*(Public |Private )?(Sub|Function) ' + [regex]::Escape($ProcName) + '\s*\('
    for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match $head) { $start = $i; break } }
    if ($start -lt 0) { return $null }
    $body = New-Object System.Collections.Generic.List[string]
    for ($j = $start + 1; $j -lt $lines.Count; $j++) {
        if ($lines[$j] -match '^\s*End (Sub|Function)\s*$') { break }
        $body.Add($lines[$j])
    }
    return ($body -join "`n")
}
function Test-HostUiOff($ui, [string]$label) {
    if ($null -eq $ui) { Check $false ($label + '：宿主设置读不到'); return }
    Check ((-not $ui.alerts) -and (-not $ui.screen)) ($label + ('（实际 DisplayAlerts={0} ScreenUpdating={1}）' -f $ui.alerts, $ui.screen))
}
function Close-AllNonAddInBooks($wbs) {
    $n = 0
    try { $n = [int]$wbs.Count } catch { $n = 0 }
    for ($i = $n; $i -ge 1; $i--) {
        $w = $null
        try { $w = $wbs.Item($i); if (-not [bool]$w.IsAddin) { $w.Close($false) } } catch {} finally { Release-Com $w }
    }
}
function Save-AllWorkbooks($wbs) {
    $n = 0
    try { $n = [int]$wbs.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $w = $null
        try { $w = $wbs.Item($i); $w.Saved = $true } catch {} finally { Release-Com $w }
    }
}
function Start-Guard([uint32]$ProcId, [int]$Seconds) {
    # 防卡死：实例若因模态框（例如编译错误的提示框）等原因卡住，到点由一个独立的隐藏进程按本测试记下的
    # 进程号结束它——卡住的 COM 调用随之报错，测试继续往下走并记 FAIL，而不是无限期挂着
    if ($ProcId -eq 0) { return $null }
    $cmd = "Start-Sleep -Seconds $Seconds; if (Get-Process -Id $ProcId -ErrorAction SilentlyContinue) { Stop-Process -Id $ProcId -Force -ErrorAction SilentlyContinue }"
    return Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile', '-WindowStyle', 'Hidden', '-Command', $cmd -PassThru -WindowStyle Hidden
}

function Invoke-Contract([bool]$EventsAtLoad, [string]$Tag) {
    # 在一个隐藏的 COM 实例里走一遍契约。$EventsAtLoad：打开加载项那一刻宿主的事件开关
    $xl = $null; $wbs = $null; $wb = $null; $ws = $null; $guard = $null
    $procId = [uint32]0
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.ScreenUpdating = $false
        $xl.EnableEvents = $EventsAtLoad
        [void][AttachLinkInitPid]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$procId)
        $guard = Start-Guard $procId 150
        $menuAtStart = Test-Menu $xl
        if ($menuAtStart) { Write-Host "    注意：$Tag 实例一创建菜单就在位（命令栏缓存残留），「打开不隐式初始化」这一条没有判别力" }

        $null = $xl.Workbooks.Open($xlam)
        Test-HostUiOff (Read-HostUiState $xl) "$Tag 打开加载项后宿主两项设置仍为 False（装载那一刻不再复位）"
        if (-not $menuAtStart) { Check (-not (Test-Menu $xl)) "$Tag 打开加载项不隐式初始化：菜单不在位" }

        $null = $xl.Run('AttachLink_Initialize')
        $xl.EnableEvents = $true
        Check (Test-Menu $xl) "$Tag 显式初始化后菜单在位"
        Check ($xl.Run('AttachLink_SelfTest') -eq $true) "$Tag SelfTest 为真"
        $ver = [string]$xl.Run('AttachLink_GetVersion')
        Check ($ver -eq $script:expectVer) ('{0} 版本 = {1}（实际 {2}）' -f $Tag, $script:expectVer, $ver)
        Test-HostUiOff (Read-HostUiState $xl) "$Tag 初始化后宿主两项设置仍为 False"

        # 新建并保存一本工作簿：链接指向它的附件夹里的文件（绝对路径），保存前应被归一化成相对路径
        $proj = Join-Path $base ('proj_' + $Tag)
        $attDir = Join-Path $proj $attach
        New-Item -ItemType Directory -Force -Path $attDir | Out-Null
        Set-Content -LiteralPath (Join-Path $attDir 'a.txt') -Value 'init-contract'
        $wbs = $xl.Workbooks
        $wb = $wbs.Add()
        $ws = $wb.Worksheets.Item(1)
        $wb.SaveAs((Join-Path $proj 'main.xlsx'), 51)
        Test-HostUiOff (Read-HostUiState $xl) "$Tag 新建并保存工作簿后宿主两项设置仍为 False"
        # 应用事件已绑定的证据：保存前归一化只在「保存前」事件里跑。给这本工作簿放一个指向其附件夹的绝对路径
        # 链接，再用测试钩子让第一次改指失败一次——保存触发归一化时会记下一笔失败，并经 OnTime 在状态栏提醒。
        # 事件没绑定的话，保存不会触发归一化，失败记录不存在，状态栏不变
        $cell = $null; $hl = $null
        try {
            $cell = $ws.Cells.Item(1, 1)
            $null = $ws.Hyperlinks.Add($cell, (Join-Path $attDir 'a.txt'), '', 'A')
            $hl = $cell.Hyperlinks.Item(1)
            $hl.Address = (Join-Path $attDir 'a.txt')
        } finally { Release-Com $hl; Release-Com $cell }
        $ws.Protect()
        $null = $xl.Run('AttachLink_TestRepointFail', 1)
        $xl.StatusBar = $false
        $wb.Save()
        $notified = $false
        for ($k = 0; $k -lt 24; $k++) {
            Start-Sleep -Milliseconds 250
            if ([string]$xl.StatusBar -like '*未能归一化为相对路径*') { $notified = $true; break }
        }
        if (-not $notified) {
            # OnTime 没赶上就直接问一次：有失败记录才会改状态栏
            $null = $xl.Run('AttachLink_NotifyNormalizeFailure', $wb)
            $notified = ([string]$xl.StatusBar -like '*未能归一化为相对路径*')
        }
        Check $notified ('{0} 保存前归一化事件已绑定：保存触发了归一化，失败记录在状态栏提醒（状态栏 [{1}]）' -f $Tag, [string]$xl.StatusBar)
        $null = $xl.Run('AttachLink_TestRepointFail', 0)
        $ws.Unprotect()
        $xl.StatusBar = $false

        # 关掉再打开这本工作簿：宿主两项设置不再被改回 True
        $wb.Close($true); Release-Com $wb; $wb = $null
        Release-Com $ws; $ws = $null
        $wb = $wbs.Open((Join-Path $proj 'main.xlsx'))
        Test-HostUiOff (Read-HostUiState $xl) "$Tag 打开另一本工作簿后宿主两项设置仍为 False（「打开工作簿」的事件处理已去掉）"

        # 关闭全部工作簿再开一本
        Close-AllNonAddInBooks $wbs
        Release-Com $wb; $wb = $null
        $wb = $wbs.Add()
        $ws = $wb.Worksheets.Item(1)
        $wb.SaveAs((Join-Path $proj 'again.xlsx'), 51)
        Check (Test-Menu $xl) "$Tag 关闭全部工作簿再开一本：菜单仍在位"
        Remove-MenuTagControls $xl
        Check (-not (Test-Menu $xl)) "$Tag 外部删除菜单控件后菜单不在位"
        $null = $xl.Run('AttachLink_EnsureMenu')
        Check (Test-Menu $xl) "$Tag EnsureMenu（右键自愈用的过程）把菜单补回"
        Test-HostUiOff (Read-HostUiState $xl) "$Tag 全部步骤之后宿主两项设置仍为 False"
    } catch {
        Check $false ("$Tag 执行异常: " + $_.Exception.Message.Split("`n")[0])
    } finally {
        if ($null -ne $guard) { try { if (-not $guard.HasExited) { $guard.Kill() } } catch {} }
        if ($null -ne $xl) { try { Save-AllWorkbooks $wbs } catch {}; try { $xl.Quit() } catch {} }
        Release-Com $ws; Release-Com $wb; Release-Com $wbs; Release-Com $xl
        if ($procId -gt 0) {
            $exited = Wait-Exit $procId 30
            Check $exited "$Tag 实例在 Quit 后 30 秒内退出（关后即退）"
            if (-not $exited) { try { Stop-Process -Id $procId -Force } catch {}; Write-Host "    已清理残留测试实例 pid=$procId" }
        }
    }
}

'=== A. 静态断言（源码） ==='
$tw = ''; $bas = ''
if ((Test-Path -LiteralPath $srcTw) -and (Test-Path -LiteralPath $srcBas) -and (Test-Path -LiteralPath $xlam)) {
    $tw = [IO.File]::ReadAllText($srcTw, [Text.Encoding]::UTF8)
    $bas = [IO.File]::ReadAllText($srcBas, [Text.Encoding]::UTF8)
    Check $true 'A0 源码与 dist\AttachLink.xlam 都在'
} else {
    Check $false 'A0 找不到 src\AttachLink.ThisWorkbook.vba、src\AttachLink.bas 或 dist\AttachLink.xlam（先运行 build\build.ps1）'
}
$handlers = @([regex]::Matches($tw, '(?m)^\s*Private Sub (app_\w+)\(') | ForEach-Object { $_.Groups[1].Value } | Sort-Object)
Check (@([regex]::Matches($tw, '(?m)^\s*(Public |Private )?Sub Workbook_Open\s*\(')).Count -eq 0) 'A1 ThisWorkbook 没有 Workbook_Open'
Check (@($handlers | Where-Object { $_ -eq 'app_WorkbookOpen' }).Count -eq 0) 'A2 不处理「打开工作簿」事件（没有 app_WorkbookOpen）'
Check (($handlers -join ',') -eq 'app_SheetBeforeRightClick,app_WorkbookBeforeSave') ('A3 应用事件只有右键与保存前两条（实际：' + ($handlers -join ',') + '）')
$bindBody = Get-ProcBody $tw 'AttachLink_BindEvents'
Check ($null -ne $bindBody -and $bindBody.Contains('Set app = Application') -and -not $bindBody.Contains('On Error')) 'A4 AttachLink_BindEvents：过程体是 Set app = Application，不吞错误'
Check (@([regex]::Matches($tw, '(?m)^\s*Public Sub AttachLink_BindEvents\s*\(')).Count -eq 1) 'A4 AttachLink_BindEvents 是 Public（标准模块要调用它）'
$initBody = Get-ProcBody $bas 'AttachLink_Initialize'
Check ($null -ne $initBody -and $initBody.Contains('ThisWorkbook.AttachLink_BindEvents') -and $initBody.IndexOf('ThisWorkbook.AttachLink_BindEvents') -lt $initBody.IndexOf('RegisterMenu True')) 'A5 AttachLink_Initialize 先挂应用事件（带 ThisWorkbook 对象名）、再注册菜单'
$autoBody = Get-ProcBody $bas 'Auto_Open'
Check ($null -ne $autoBody -and $autoBody.Contains('AttachLink_Initialize')) 'A6 Auto_Open 仍调用 AttachLink_Initialize（交互式启动的入口）'
$rcBody = Get-ProcBody $tw 'app_SheetBeforeRightClick'
Check ($null -ne $rcBody -and $rcBody.Contains('AttachLink_EnsureMenu') -and $rcBody.Contains('AttachLink_NotifyNormalizeFailure')) 'A7 右键处理：补回菜单 + 归一化失败提醒（原「打开时提醒」改到这里）'
$verMatch = [regex]::Match($bas, 'ATTACHLINK_VERSION As String = "([0-9.]+)"')
$script:expectVer = $verMatch.Groups[1].Value
Check ($script:expectVer -ne '') 'A8 源码里的版本常量可读'

if ($script:expectVer -ne '') {
    '=== B. 隐藏实例：装载时事件开启 ==='
    Invoke-Contract $true 'B'
    '=== C. 隐藏实例：装载与初始化时事件关闭，之后打开事件 ==='
    Invoke-Contract $false 'C'
}

Start-Sleep -Seconds 1
try { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue } catch {}
Write-Host ("RESULT: " + $(if ($script:ok) { 'PASS' } else { 'FAIL' }))
if ($script:ok) { exit 0 } else { exit 1 }
