# =====================================================================
# AttachLink 看门狗 / 安装脚本 —— 单元级回归（v1.4.4）
# ---------------------------------------------------------------------
# A. 看门狗装载判定（用假对象逐条覆盖，不依赖 Excel）：宿主里在位的副本是否就是看门狗配置的那一份，
#    决定「能否执行初始化宏」。只有 trusted 才允许；查询出错不得当成「不在位」去打开第二份。
#    v1.6.0（R1 只装载不登记）：装载只发生 Workbooks.Open，notfound 时断言没有任何 AddIns 写入
#    （不 Add、不置 Installed）。v1.6.0（R2 在位判定加固）：Get-PresenceDecision 五结论逐条回归
#    （notloaded / inplace / foreign / deadmenu / unknown）。
#    A2 用假状态 + 受控时钟回归「隐藏文档窗口状态机」：读不到 Window.Visible 只延迟重试；重新评估只在真实的
#    「不可见 → 可见」跳变时计数，用完后仍有低频重查——任何路径都不会让窗口被永久跳过。
# B. 同名异路径副本（真实 Excel，隐藏实例）：宿主里装载了一份来自别处、没有宏的 AttachLink.xlam 时，
#    看门狗不打开第二份、不执行它的宏，该实例随后仍能正常退出。
# C. 安装脚本宿主识别（build\install.ps1 的 Test-IsNotRealExcel）：无法确认是 Microsoft Excel 时按
#    「不是 Excel」处理（fail-closed）。
# D. 路径比较与 Release-Com 的容错。
# E. 隐藏的文档窗口不作为处理对象（真实 Excel，隐藏实例）：加载项自身的隐藏窗口零处理；被隐藏的工作簿窗口
#    显示出来之后才处理。
# F. 加载项列表登记清理（build\install.ps1，临时注册表键，不依赖 Excel）：Excel 的加载项列表按文件名识别
#    条目，安装前清掉指向其他路径的 AttachLink 登记、卸载时清掉全部；删 OPEN 条目后序列必须仍然连续。
# G. 附着退避与成功判定（v1.6.1）：Get-BackoffDelay 纯函数逐条回归；主循环接线做静态断言（退避检查排在
#    取 Application 之前、finally 里记账、窗口消失后清理；宏名限定到工作簿；初始化/菜单/版本三项齐备才标 seen）。
# 用法：powershell -ExecutionPolicy Bypass -File test-watchdog-units.ps1      期望末行 RESULT: PASS
# 需要：桌面版 Excel（B、E 段用隐藏实例，不弹窗口）。本机常驻看门狗若与被测脚本版本不一致，B、E 段整段跳过
#       （它会对同一实例做自己的动作，结果无法归因）。
# 清理约定：只按本测试记录下来的进程号结束自己拉起的 Excel，绝不按「测试期间新出现的 Excel 进程」批量清理
#           ——那会误杀用户恰好在这时打开的 Excel。本测试碰过的 COM 对象都显式释放到计数归零，不依赖 GC
#           （被枚举过的集合只靠 GC 回收不可靠，会把实例钉住，让「能退出」类的检查失真）。
# 编码：本文件含中文，必须保存为 UTF-8 with BOM（WORKLOG 根因 #11）。
# =====================================================================
# 可选参数：-AssumeSameVersion  跳过「常驻看门狗与被测脚本同版」的检查（仅在本机常驻看门狗已暂停、要单独验证
#           工作区版本时使用；常驻看门狗在运行时不要加，它会对同一实例做自己的动作，结果无法归因）。
param([switch]$AssumeSameVersion)
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$repo = Split-Path -Parent $PSScriptRoot
$wdPath = Join-Path $PSScriptRoot 'watchdog.ps1'
$installPath = Join-Path $repo 'build\install.ps1'
$base = Join-Path $env:TEMP 'opencode\wdunits'
Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path (Join-Path $base 'foreign') | Out-Null

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }
function Get-FunctionsText([string]$File, [string[]]$Names) {
    # 从脚本文件里按名字取出函数定义文本（不执行脚本主体）。调用方须在脚本作用域用「. 」载入——
    # 放在函数内部点源载入只会定义到那个函数的局部作用域，返回后就没了。
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($File, [ref]$null, [ref]$null)
    $found = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false) |
        Where-Object { $Names -contains $_.Name }
    return (@($found | ForEach-Object { $_.Extent.Text }) -join "`n")
}

Add-Type -TypeDefinition @'
using System; using System.Collections.Generic; using System.Runtime.InteropServices; using System.Text;
public static class AttachLinkUnitsPid {
  public delegate bool P(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool EnumWindows(P cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, P cb, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int m);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  // 某进程的 EXCEL7 文档窗口句柄（XLMAIN 的子窗口）；shownOnly=true 时只取 Win32 可见的
  public static List<long> Excel7(uint pid, bool shownOnly) {
    var r = new List<long>();
    EnumWindows((t, l) => {
      var c = new StringBuilder(64); GetClassName(t, c, 64);
      if (c.ToString() == "XLMAIN") {
        EnumChildWindows(t, (k, l2) => {
          uint p; GetWindowThreadProcessId(k, out p);
          var kc = new StringBuilder(64); GetClassName(k, kc, 64);
          if (p == pid && kc.ToString() == "EXCEL7" && (!shownOnly || IsWindowVisible(k))) r.Add(k.ToInt64());
          return true; }, IntPtr.Zero);
      }
      return true; }, IntPtr.Zero);
    return r;
  }
}
'@

# 机器级常驻看门狗与被测脚本是否同版（哈希一致）：B、E 段的闸
$pfWd = 'C:\Program Files\AttachLink\watchdog.ps1'
$deployedDiffers = $false
if ((-not $AssumeSameVersion) -and (Test-Path $pfWd)) { $deployedDiffers = ((Get-FileHash $pfWd).Hash -ne (Get-FileHash $wdPath).Hash) }

# ---------------------------------------------------------------------
'=== D. 路径比较与 Release-Com 容错（纯函数）==='
$fnText = Get-FunctionsText $wdPath @('Release-Com', 'Test-SamePath', 'Get-LoadedAddInPath', 'Get-PresenceDecision', 'Remove-DeadMenuControls', 'Menu-Exists', 'Get-HostUiState', 'Restore-HostUiState', 'Provision-App', 'New-WindowState', 'Get-HiddenDecision', 'Test-HiddenReevaluate', 'Get-BackoffDelay')
Check ($fnText.Contains('function Release-Com') -and $fnText.Contains('function Test-SamePath') -and $fnText.Contains('function Get-LoadedAddInPath') -and $fnText.Contains('function Get-PresenceDecision') -and $fnText.Contains('function Remove-DeadMenuControls') -and $fnText.Contains('function Menu-Exists') -and $fnText.Contains('function Get-HostUiState') -and $fnText.Contains('function Restore-HostUiState') -and $fnText.Contains('function Provision-App') -and $fnText.Contains('function New-WindowState') -and $fnText.Contains('function Get-HiddenDecision') -and $fnText.Contains('function Test-HiddenReevaluate') -and $fnText.Contains('function Get-BackoffDelay')) '取到 watchdog.ps1 的相关函数'
. ([scriptblock]::Create($fnText))
$XlamPath = 'C:\Program Files\AttachLink\AttachLink.xlam'      # A 段假对象场景用的「看门狗配置的那一份」
$script:wdLog = New-Object System.Collections.Generic.List[string]
function WdLog($m) { [void]$script:wdLog.Add([string]$m) }

Check (Test-SamePath 'C:\A\AttachLink.xlam' 'c:\a\ATTACHLINK.XLAM') '同一路径（大小写 / 分隔符不同）判为相同'
Check (-not (Test-SamePath 'C:\A\AttachLink.xlam' 'C:\B\AttachLink.xlam')) '不同目录的同名文件判为不同'
Check (-not (Test-SamePath '' 'C:\A\AttachLink.xlam')) '空路径不判为相同'
Check (-not (Test-SamePath 'C:\A\AttachLink.xlam' 'relative\AttachLink.xlam')) '相对路径不误判为相同'
$threw = $false
try { Release-Com ([pscustomobject]@{ X = 1 }); Release-Com $null; Release-Com 'string' } catch { $threw = $true }
Check (-not $threw) 'Release-Com 对非 COM 对象 / $null 不抛异常'

# ---------------------------------------------------------------------
'=== G. 附着退避与成功判定（v1.6.1：纯函数 + 主循环接线静态断言）==='
# 间隔：首次取一个周期（2 秒），之后逐次翻倍，封顶 30 秒；prev<=0 视为首次；base 大于 max 时以 max 为准。
Check ((Get-BackoffDelay 0 2 30) -eq 2) '首次失败：间隔取一个周期（2 秒）'
Check ((Get-BackoffDelay 2 2 30) -eq 4) '第二次失败：间隔翻倍（4 秒）'
Check ((Get-BackoffDelay 8 2 30) -eq 16) '第三次之后继续翻倍（8 → 16 秒）'
Check ((Get-BackoffDelay 16 2 30) -eq 30) '间隔到上限后不再增长（16 → 30，而非 32）'
Check ((Get-BackoffDelay 30 2 30) -eq 30) '已在上限：保持 30 秒'
Check ((Get-BackoffDelay -5 2 30) -eq 2) 'prev 为负数按首次处理'
Check ((Get-BackoffDelay 0 5 3) -eq 3) 'base 大于 max 时以 max 为准'
Check ((Get-BackoffDelay 3 5 30) -eq 6) 'prev 小于 base 时照常翻倍（3 → 6），不回退'
$wdText = [IO.File]::ReadAllText($wdPath)
Check ($wdText.Contains('$initMacro = "''AttachLink.xlam''!AttachLink_Initialize"') -and $wdText.Contains('$versionMacro = "''AttachLink.xlam''!AttachLink_GetVersion"')) '宏名限定到工作簿（初始化与版本两个宏）'
Check (-not $wdText.Contains("Run('AttachLink_Initialize')") -and -not $wdText.Contains("Run('AttachLink_GetVersion')")) '主循环不再使用不带工作簿名的宏名'
Check ($wdText.Contains('if ($initOk -and $menuOk -and $ver)') -and $wdText.Contains('try { $null = $app.Run($initMacro); $initOk = $true }')) '成功判定：初始化、菜单、版本三者齐备才标 seen'
$boGate = 'if ($null -ne $bo -and (Get-Date) -lt $bo.until) { continue }'
Check ($wdText.Contains($boGate) -and $wdText.IndexOf($boGate) -lt $wdText.IndexOf('GetApplicationFromWindow($h, [ref]$winVisible)')) '退避检查排在取 Application 之前：退避期内不向宿主发调用'
Check ($wdText.Contains('if ($attachFailed) {') -and $wdText.Contains('$backoff[$key] = @{ until = (Get-Date).AddSeconds($boNextDelay); delay = $boNextDelay }') -and $wdText.Contains('$backoff.Remove($key)')) 'finally 里记账：失败排下次可试的时刻，得出非失败的结论即清零'
Check ($wdText.Contains('foreach ($k in @($backoff.Keys)) { if (-not $liveKeys.ContainsKey($k)) { $backoff.Remove($k) } }')) '窗口消失后清理它的退避记录'
Check (-not $wdText.Contains('Start-Sleep -Milliseconds 800')) '不再用固定 800 毫秒等待（由按窗口退避取代）'
Check ($wdText.Contains('（退避重试）') -and $wdText.Contains('WdLog ("附着失败 hwnd="')) '失败日志仍带「退避重试」字样（健康检查按它归类自愈类重试）'

# ---------------------------------------------------------------------
'=== A. 看门狗装载判定（假对象，不依赖 Excel）==='
$badIndex = [int]0x8002000B      # DISP_E_BADINDEX：Excel 对按名字取不到的工作簿抛的错误（实测）
function New-FakeWorkbook([string]$Path, [switch]$PathThrows) {
    $o = [pscustomobject]@{ Path = $Path }
    if ($PathThrows) { Add-Member -InputObject $o ScriptProperty FullName { throw [System.Runtime.InteropServices.COMException]::new('路径不可读', [int]0x80004005) } -Force }
    else { Add-Member -InputObject $o ScriptProperty FullName { $this.Path } -Force }
    return $o
}
function New-FakeApp([string]$Lookup, $LoadedWorkbook, [string]$OpenedPath, [switch]$OpenedPathThrows, $AddIn, [bool]$Alerts = $true, [bool]$Screen = $true, [switch]$AlertsUnreadable, [switch]$OpenThrows) {
    # Lookup: 'found' | 'notfound' | 'error'  —— 模拟 Workbooks.Item('AttachLink.xlam') 的三种结果
    # Open 追踪：把每次 Open 的路径记到 $script:openedPaths
    # 宿主界面设置：Alerts / Screen 是装载前的 DisplayAlerts / ScreenUpdating；Open 照 Excel 的真实行为把两项复位为 True
    # （Workbook_Open 事件处理结束时发生，见 Get-HostUiState）；AlertsUnreadable 模拟读不到 DisplayAlerts（得到 $null，
    # 且对它的写入被记入 UiWriteCalls——用来断言看门狗没有尝试写回读不到的项）；OpenThrows 模拟 Open 在事件
    # 跑完、复位已发生之后抛错
    $wbs = [pscustomobject]@{}
    Add-Member -InputObject $wbs ScriptMethod Item {
        param($name)
        if ($name -eq 'AttachLink.xlam') {
            switch ($this.Lookup) {
                'notfound' { throw [System.Runtime.InteropServices.COMException]::new('无效索引。', [int]0x8002000B) }
                'error'    { throw [System.Runtime.InteropServices.COMException]::new('被呼叫方拒绝接收呼叫。', [int]0x80010001) }
                default    { return $this.Loaded }
            }
        }
        throw [System.Runtime.InteropServices.COMException]::new('无效索引。', [int]0x8002000B)
    } -Force
    Add-Member -InputObject $wbs ScriptMethod Open {
        param($path)
        [void]$script:openedPaths.Add([string]$path)
        # Excel 的真实行为：Workbook_Open 事件处理结束时把宿主的 DisplayAlerts / ScreenUpdating 复位为 True
        if ($null -ne $this.OwnerApp) {
            if (-not $this.OwnerApp.AlertsUnreadable) { try { $this.OwnerApp.DisplayAlerts = $true } catch {} }
            try { $this.OwnerApp.ScreenUpdating = $true } catch {}
        }
        if ($this.OpenThrows) { throw [System.Runtime.InteropServices.COMException]::new('被呼叫方拒绝接收呼叫。', [int]0x80010001) }
        if ($this.OpenedThrows) { return (New-FakeWorkbook '' -PathThrows) }
        return (New-FakeWorkbook $this.Opened)
    } -Force
    Add-Member -InputObject $wbs NoteProperty Lookup $Lookup -Force
    Add-Member -InputObject $wbs NoteProperty Loaded $LoadedWorkbook -Force
    Add-Member -InputObject $wbs NoteProperty Opened $OpenedPath -Force
    Add-Member -InputObject $wbs NoteProperty OpenedThrows ([bool]$OpenedPathThrows) -Force
    Add-Member -InputObject $wbs NoteProperty OpenThrows ([bool]$OpenThrows) -Force
    $app = [pscustomobject]@{}
    Add-Member -InputObject $app ScriptProperty Workbooks { $this.Wbs } -Force
    Add-Member -InputObject $app ScriptProperty AddIns { $this.Ains } -Force
    if ($AlertsUnreadable) {
        Add-Member -InputObject $app NoteProperty UiWriteCalls 0 -Force
        Add-Member -InputObject $app -MemberType ScriptProperty -Name DisplayAlerts -Value { return $null } -SecondValue { $this.UiWriteCalls = [int]$this.UiWriteCalls + 1 } -Force
    }
    else { Add-Member -InputObject $app NoteProperty DisplayAlerts ([bool]$Alerts) -Force }
    Add-Member -InputObject $app NoteProperty ScreenUpdating ([bool]$Screen) -Force
    Add-Member -InputObject $app NoteProperty AlertsUnreadable ([bool]$AlertsUnreadable) -Force
    Add-Member -InputObject $app NoteProperty Wbs $wbs -Force
    Add-Member -InputObject $app NoteProperty Ains $AddIn -Force
    Add-Member -InputObject $wbs NoteProperty OwnerApp $app -Force
    return $app
}
function New-FakeAddIns([string]$EntryPath, [bool]$Installed) {
    # AddIns 集合：至多一个名为 AttachLink.xlam 的登记项。AddCalls 记录 Add 被调用的次数
    # ——R1（只装载不登记）后看门狗不得再碰 AddIns 集合，任何 Add 调用都是回归。
    $entry = $null
    if ($EntryPath) { $entry = [pscustomobject]@{ Name = 'AttachLink.xlam'; FullName = $EntryPath; Installed = $Installed } }
    $ains = [pscustomobject]@{ AddCalls = 0 }
    Add-Member -InputObject $ains ScriptProperty Count { if ($this.Entry) { 1 } else { 0 } } -Force
    Add-Member -InputObject $ains ScriptMethod Item { param($i) if ($i -eq 1 -and $this.Entry) { return $this.Entry }; throw [System.Runtime.InteropServices.COMException]::new('无效索引。', [int]0x8002000B) } -Force
    Add-Member -InputObject $ains ScriptMethod Add { param($p, $c) $this.AddCalls++; $this.Entry = [pscustomobject]@{ Name = 'AttachLink.xlam'; FullName = [string]$p; Installed = $false }; return $this.Entry } -Force
    Add-Member -InputObject $ains NoteProperty Entry $entry -Force
    return $ains
}
function Invoke-Provision($app) {
    $script:openedPaths = New-Object System.Collections.Generic.List[string]
    $script:wdLog.Clear()
    $ret = Provision-App $app
    return @{ ret = ($ret -is [bool] -and $ret); state = $script:provState; opened = @($script:openedPaths); log = @($script:wdLog) }
}

$r = Invoke-Provision (New-FakeApp 'notfound' $null $XlamPath -AddIn (New-FakeAddIns $XlamPath $true))
Check ($r.ret -and $r.state -eq 'trusted' -and $r.opened.Count -eq 1 -and $r.opened[0] -eq $XlamPath) '不在位 → 从配置路径打开，打开后复核一致 → trusted'

$r = Invoke-Provision (New-FakeApp 'notfound' $null 'C:\Elsewhere\AttachLink.xlam' -AddIn (New-FakeAddIns $XlamPath $true))
Check ($r.ret -and $r.state -eq 'mismatch') ('打开后读到的在位副本不是配置的那份 → mismatch（不执行宏）；实际 state=' + $r.state)

$r = Invoke-Provision (New-FakeApp 'notfound' $null '' -OpenedPathThrows -AddIn (New-FakeAddIns $XlamPath $true))
Check ($r.ret -and $r.state -eq 'mismatch') ('打开后路径读不到 → mismatch（不执行宏）；实际 state=' + $r.state)

$r = Invoke-Provision (New-FakeApp 'error' $null $XlamPath -AddIn (New-FakeAddIns $XlamPath $true))
Check ($r.ret -and $r.state -eq 'unknown' -and $r.opened.Count -eq 0) ('查询出错（宿主繁忙）→ unknown，且没有打开任何文件；实际 state=' + $r.state + ' opened=' + $r.opened.Count)

$r = Invoke-Provision (New-FakeApp 'found' (New-FakeWorkbook 'C:\Elsewhere\AttachLink.xlam') $XlamPath -AddIn (New-FakeAddIns $XlamPath $true))
Check ($r.ret -and $r.state -eq 'foreign' -and $r.opened.Count -eq 0) ('在位的是异路径副本 → foreign，不打开第二份；实际 state=' + $r.state)

$r = Invoke-Provision (New-FakeApp 'found' (New-FakeWorkbook '' -PathThrows) $XlamPath -AddIn (New-FakeAddIns $XlamPath $true))
Check ($r.ret -and $r.state -eq 'unknown' -and $r.opened.Count -eq 0) ('在位但路径读不到 → unknown，不打开第二份；实际 state=' + $r.state)

$r = Invoke-Provision (New-FakeApp 'found' (New-FakeWorkbook $XlamPath) $XlamPath -AddIn (New-FakeAddIns $XlamPath $true))
Check ($r.ret -and $r.state -eq 'trusted' -and $r.opened.Count -eq 0) '在位的就是配置的那份 → trusted，不重复打开'

# R1 只装载不登记：notfound 时不得有任何 AddIns 写入（不 Add、不置 Installed），只发生 Workbooks.Open($XlamPath)
$ainsForeign = New-FakeAddIns 'C:\Elsewhere\AttachLink.xlam' $false
$r = Invoke-Provision (New-FakeApp 'notfound' $null $XlamPath -AddIn $ainsForeign)
Check ($r.state -eq 'trusted' -and $r.opened.Count -eq 1 -and $r.opened[0] -eq $XlamPath -and -not $ainsForeign.Entry.Installed -and $ainsForeign.AddCalls -eq 0) '登记在别处的同名条目：不被置为 Installed、不新增登记（看门狗不碰 AddIns 集合）'
$ainsMine = New-FakeAddIns $XlamPath $false
$r = Invoke-Provision (New-FakeApp 'notfound' $null $XlamPath -AddIn $ainsMine)
Check ($r.state -eq 'trusted' -and -not $ainsMine.Entry.Installed -and $ainsMine.AddCalls -eq 0) '登记路径就是配置路径的条目：也不被置为 Installed（登记位 ≠ 装载态，置位无装载效果且会写脏列表）'
$ainsNone = New-FakeAddIns '' $false
$r = Invoke-Provision (New-FakeApp 'notfound' $null $XlamPath -AddIn $ainsNone)
Check ($r.state -eq 'trusted' -and $r.opened.Count -eq 1 -and $null -eq $ainsNone.Entry -and $ainsNone.AddCalls -eq 0) '没有登记项：不新增任何登记，只从配置路径装载'

# 装载不留副作用（v1.6.0）：Excel 在 Workbook_Open 事件处理结束时会把宿主的 DisplayAlerts / ScreenUpdating 复位为 True
# （2026-11-26 实测，任何带 Workbook_Open 的加载项都会触发；假对象的 Open 照此模拟）。看门狗装载后须恢复装载前的值——
# 否则别的程序拉起的 Excel 里，被关掉的警告会重新打开，一个「是否替换」提示框就能把隐藏的自动化实例卡死（实测卡了 22 分钟）
$st = Get-HostUiState ([pscustomobject]@{ DisplayAlerts = $false; ScreenUpdating = $false })
Check ($st.alerts -eq $false -and $st.screen -eq $false) 'Get-HostUiState：读到宿主两项设置的原值'
$uo = [pscustomobject]@{ DisplayAlerts = $true; ScreenUpdating = $true }
Restore-HostUiState $uo $st
Check ($uo.DisplayAlerts -eq $false -and $uo.ScreenUpdating -eq $false) 'Restore-HostUiState：把两项写回原值'
$uo = [pscustomobject]@{ DisplayAlerts = $true; ScreenUpdating = $true }
Restore-HostUiState $uo @{ alerts = $null; screen = $false }
Check ($uo.DisplayAlerts -eq $true -and $uo.ScreenUpdating -eq $false) 'Restore-HostUiState：快照里没读到的项（$null）不写回'
$threw = $false; try { Restore-HostUiState $uo $null } catch { $threw = $true }
Check (-not $threw) 'Restore-HostUiState：快照为 $null 时不做任何事、不抛异常'
$appUi = New-FakeApp 'notfound' $null $XlamPath -Alerts $false -Screen $false -AddIn (New-FakeAddIns '' $false)
$r = Invoke-Provision $appUi
Check ($r.state -eq 'trusted' -and $appUi.DisplayAlerts -eq $false -and $appUi.ScreenUpdating -eq $false) ('装载后宿主设置恢复原值：装载前两项均为 False → 装载后仍为 False；实际 DisplayAlerts=' + $appUi.DisplayAlerts + ' ScreenUpdating=' + $appUi.ScreenUpdating)
$appUi = New-FakeApp 'notfound' $null $XlamPath -AddIn (New-FakeAddIns '' $false)
$r = Invoke-Provision $appUi
Check ($r.state -eq 'trusted' -and $appUi.DisplayAlerts -eq $true -and $appUi.ScreenUpdating -eq $true) '装载后宿主设置恢复原值：装载前两项均为 True（交互式会话的常态）→ 装载后仍为 True'
$appUi = New-FakeApp 'notfound' $null $XlamPath -Alerts $false -AddIn (New-FakeAddIns '' $false)
$r = Invoke-Provision $appUi
Check ($r.state -eq 'trusted' -and $appUi.DisplayAlerts -eq $false -and $appUi.ScreenUpdating -eq $true) ('装载后宿主设置恢复原值：只关了警告 → 逐项恢复；实际 DisplayAlerts=' + $appUi.DisplayAlerts + ' ScreenUpdating=' + $appUi.ScreenUpdating)
$appUi = New-FakeApp 'notfound' $null 'C:\Elsewhere\AttachLink.xlam' -Alerts $false -Screen $false -AddIn (New-FakeAddIns '' $false)
$r = Invoke-Provision $appUi
Check ($r.state -eq 'mismatch' -and $appUi.DisplayAlerts -eq $false -and $appUi.ScreenUpdating -eq $false) '打开后复核不符（mismatch）：复位已经发生，宿主设置照样恢复'
$appUi = New-FakeApp 'notfound' $null $XlamPath -OpenedPathThrows -Alerts $false -Screen $false -AddIn (New-FakeAddIns '' $false)
$r = Invoke-Provision $appUi
Check ($r.state -eq 'mismatch' -and $appUi.DisplayAlerts -eq $false -and $appUi.ScreenUpdating -eq $false) '打开后路径读不到（mismatch）：宿主设置照样恢复'
$appUi = New-FakeApp 'notfound' $null $XlamPath -OpenThrows -Alerts $false -Screen $false -AddIn (New-FakeAddIns '' $false)
$r = Invoke-Provision $appUi
Check ((-not $r.ret) -and $r.state -eq 'failed' -and $r.opened.Count -eq 1 -and $appUi.DisplayAlerts -eq $false -and $appUi.ScreenUpdating -eq $false) ('Workbooks.Open 中途抛错（事件可能已跑完、复位已发生）：装载判失败，宿主设置照样恢复；实际 state=' + $r.state + ' DisplayAlerts=' + $appUi.DisplayAlerts + ' ScreenUpdating=' + $appUi.ScreenUpdating)
$appUi = New-FakeApp 'found' (New-FakeWorkbook $XlamPath) $XlamPath -Alerts $false -Screen $false -AddIn (New-FakeAddIns $XlamPath $true)
$r = Invoke-Provision $appUi
Check ($r.state -eq 'trusted' -and $r.opened.Count -eq 0 -and $appUi.DisplayAlerts -eq $false -and $appUi.ScreenUpdating -eq $false) '在位的就是配置的那份（不打开）：不碰宿主设置'
$appUi = New-FakeApp 'notfound' $null $XlamPath -Alerts $false -Screen $false -AlertsUnreadable -AddIn (New-FakeAddIns '' $false)
$r = Invoke-Provision $appUi
Check ($r.state -eq 'trusted' -and $null -eq $appUi.DisplayAlerts -and $appUi.ScreenUpdating -eq $false -and $appUi.UiWriteCalls -eq 0) ('装载前读不到 DisplayAlerts（得到空值）：不臆测、不尝试写回该项（对它的写入次数须为 0，实际 ' + $appUi.UiWriteCalls + '），只恢复读得到的 ScreenUpdating')
$wdTextUi = [IO.File]::ReadAllText($wdPath, [Text.Encoding]::UTF8)
$iUiCapture = $wdTextUi.IndexOf('$uiState = Get-HostUiState $app')
$iUiOpen = $wdTextUi.IndexOf('$wb = $wbs.Open($XlamPath)')
$iUiRestore = $wdTextUi.IndexOf('Restore-HostUiState $app $uiState')
Check ($iUiCapture -gt 0 -and $iUiOpen -gt $iUiCapture -and $iUiRestore -gt $iUiOpen) '装载不留副作用：Provision-App 在 Workbooks.Open 之前记录宿主设置、之后立即恢复'

# Get-PresenceDecision 五结论（在位判定的纯函数，不碰 COM）
Check ((Get-PresenceDecision $true 'found' $true) -eq 'inplace') '在位判定：菜单在位 + 在位的就是配置路径 → inplace'
Check ((Get-PresenceDecision $true 'found' $false) -eq 'foreign') '在位判定：菜单在位 + 异路径同名副本 → foreign（维持不动）'
Check ((Get-PresenceDecision $true 'notfound' $false) -eq 'deadmenu') '在位判定：菜单在位 + 加载项工作簿明确不在位 → deadmenu（命令栏缓存残留，按未装载处理）'
Check ((Get-PresenceDecision $true 'error' $false) -eq 'unknown') '在位判定：菜单在位 + 查询出错 → unknown（不标 seen，下周期重试）'
foreach ($ls in @('found', 'notfound', 'error')) {
    Check ((Get-PresenceDecision $false $ls $false) -eq 'notloaded') ("在位判定：菜单不在位（lookupState=$ls）→ notloaded（走装载流程）")
}

# Remove-DeadMenuControls（死菜单根除：从外部删除 Cell 栏上带 AttachLinkCtx 标签的控件；假对象，不依赖 Excel）
function New-FakeCellBar([string[]]$Tags, [int]$TagThrowsAt, [int]$DeleteThrowsAt, [switch]$CountNull) {
    # $Tags：每个控件（下标从 1 起）的 Tag；$TagThrowsAt / $DeleteThrowsAt：读取 Tag / 删除时抛异常的下标（0 = 不抛）
    $script:fakeDeletes = New-Object System.Collections.Generic.List[int]
    $ctrls = [pscustomobject]@{}
    Add-Member -InputObject $ctrls ScriptProperty Count { if ($this.CountNull) { return $null }; $this.Tags.Length } -Force
    Add-Member -InputObject $ctrls ScriptMethod Item {
        param($i)
        if ($i -lt 1 -or $i -gt $this.Tags.Length) { throw [System.Runtime.InteropServices.COMException]::new('无效索引。', [int]0x8002000B) }
        $c = [pscustomobject]@{ Index = [int]$i }
        Add-Member -InputObject $c NoteProperty Owner $this -Force
        if ([int]$i -eq $this.TagThrowsAt) {
            # 「Tag 读不到」：按 PS 5.1 的真实表现建模——属性读取失败不上抛而是静默得到 $null（真实 Excel 里
            # 无标签控件的 Tag 是空串 ""，所以看门狗可以用 $null 判「读不到」）
            Add-Member -InputObject $c ScriptProperty Tag { return $null } -Force
        } else {
            Add-Member -InputObject $c NoteProperty Tag ([string]$this.Tags[[int]$i - 1]) -Force
        }
        Add-Member -InputObject $c ScriptMethod Delete {
            if ($this.Index -eq $this.Owner.DeleteThrowsAt) { throw [System.Runtime.InteropServices.COMException]::new('被呼叫方拒绝接收呼叫。', [int]0x80010001) }
            [void]$this.Owner.Deletes.Add($this.Index)
        } -Force
        return $c
    } -Force
    Add-Member -InputObject $ctrls NoteProperty Tags ([string[]]$Tags) -Force
    Add-Member -InputObject $ctrls NoteProperty TagThrowsAt $TagThrowsAt -Force
    Add-Member -InputObject $ctrls NoteProperty DeleteThrowsAt $DeleteThrowsAt -Force
    Add-Member -InputObject $ctrls NoteProperty CountNull ([bool]$CountNull) -Force
    Add-Member -InputObject $ctrls NoteProperty Deletes $script:fakeDeletes -Force
    $bar = [pscustomobject]@{}
    Add-Member -InputObject $bar ScriptProperty Controls { $this.Ctrls } -Force
    Add-Member -InputObject $bar NoteProperty Ctrls $ctrls -Force
    $bars = [pscustomobject]@{}
    Add-Member -InputObject $bars ScriptMethod Item { param($n) if ($n -eq 'Cell') { return $this.Bar }; throw [System.Runtime.InteropServices.COMException]::new('无效索引。', [int]0x8002000B) } -Force
    Add-Member -InputObject $bars NoteProperty Bar $bar -Force
    $app = [pscustomobject]@{}
    Add-Member -InputObject $app ScriptProperty CommandBars { $this.Bars } -Force
    Add-Member -InputObject $app NoteProperty Bars $bars -Force
    return $app
}
$appDm = New-FakeCellBar @('', 'AttachLinkCtx', '', 'AttachLinkCtx', '') 0 0
$rDm = Remove-DeadMenuControls $appDm
Check (($rDm -is [int]) -and $rDm -eq 2) ('死菜单根除：删除全部带标签的控件，返回删除个数；实际=' + $rDm)
Check (($script:fakeDeletes -join ',') -eq '4,2') ('死菜单根除：倒序下标遍历（删除不影响未遍历部分的下标）；实际删除顺序=' + ($script:fakeDeletes -join ','))
$appDm = New-FakeCellBar @('', 'Other', '') 0 0
$rDm = Remove-DeadMenuControls $appDm
Check (($rDm -is [int]) -and $rDm -eq 0 -and $script:fakeDeletes.Count -eq 0) '死菜单根除：没有带标签的控件 → 0，不删除任何控件'
$appDm = New-FakeCellBar @('', 'Other', '', 'AttachLinkCtx', '') 4 0
$rDm = Remove-DeadMenuControls $appDm
Check (($rDm -is [int]) -and $rDm -eq -1 -and $script:fakeDeletes.Count -eq 0) ('死菜单根除：带标签的控件 Tag 读不到（宿主繁忙）→ 一个都没删成时不谎报「未找到」，返回 -1；实际=' + $rDm)
$appDm = New-FakeCellBar @('', 'Other', 'x') 2 0
$rDm = Remove-DeadMenuControls $appDm
Check (($rDm -is [int]) -and $rDm -eq -1 -and $script:fakeDeletes.Count -eq 0) ('死菜单根除：无标签控件的 Tag 读不到、没有删任何控件 → 读不到的那个可能就是死菜单，返回 -1；实际=' + $rDm)
$appDm = New-FakeCellBar @('', 'AttachLinkCtx', '', 'AttachLinkCtx', '') 0 4
$rDm = Remove-DeadMenuControls $appDm
Check (($rDm -is [int]) -and $rDm -eq 1 -and ($script:fakeDeletes -join ',') -eq '2') ('死菜单根除：单个控件删不掉 → 只跳过它，其余照删，只计成功数；实际=' + $rDm + ' 删除顺序=' + ($script:fakeDeletes -join ','))
# 控件集合读不到 Count：PS 5.1 对名为 Count 的属性有特殊处理，getter 抛异常时不上抛而是静默得到空值（实测，
# 同一轮里名为 Tag 的属性抛异常能被正常捕获），所以按「Count 取到空值」建模——看门狗必须显式判空，不能当成 0 个控件
$script:fakeDeletes = New-Object System.Collections.Generic.List[int]
$rDm = Remove-DeadMenuControls (New-FakeCellBar @('', 'AttachLinkCtx') 0 0 -CountNull)
Check (($rDm -is [int]) -and $rDm -eq -1 -and $script:fakeDeletes.Count -eq 0) ('死菜单根除：控件集合读不到 Count（取到空值）→ -1，未删除任何控件；实际=' + $rDm)
$appBusy = [pscustomobject]@{}
Add-Member -InputObject $appBusy ScriptProperty CommandBars { throw [System.Runtime.InteropServices.COMException]::new('被呼叫方拒绝接收呼叫。', [int]0x80010001) } -Force
$threw = $false; $rDm = $null
try { $rDm = Remove-DeadMenuControls $appBusy } catch { $threw = $true }
Check ((-not $threw) -and ($rDm -is [int]) -and $rDm -eq -1) ('死菜单根除：命令栏整体不可访问（宿主正忙）→ -1，不抛异常；实际=' + $rDm)
$appNoBar = [pscustomobject]@{}
Add-Member -InputObject $appNoBar ScriptProperty CommandBars { $this.Bars } -Force
Add-Member -InputObject $appNoBar NoteProperty Bars ([pscustomobject]@{}) -Force
$threw = $false; $rDm = $null
try { $rDm = Remove-DeadMenuControls $appNoBar } catch { $threw = $true }
Check ((-not $threw) -and ($rDm -is [int]) -and $rDm -eq -1) ('死菜单根除：取不到 Cell 命令栏 → -1，不抛异常；实际=' + $rDm)
$wdAll = [IO.File]::ReadAllText($wdPath, [Text.Encoding]::UTF8)
$iDmInit = $wdAll.IndexOf('$deadMenu = $false')
$iDmFlag = $wdAll.IndexOf('$deadMenu = $true')
$iDmPurge = $wdAll.IndexOf('$deadRemoved = Remove-DeadMenuControls $app')
$iDmProv = $wdAll.IndexOf('if (Provision-App $app) {')
Check ($iDmInit -gt 0 -and $iDmFlag -gt $iDmInit -and $iDmPurge -gt $iDmFlag -and $iDmProv -gt $iDmPurge) '死菜单根除：主循环在每个窗口先复位 $deadMenu，deadmenu 结论之后、装载之前调用 Remove-DeadMenuControls'

# Menu-Exists：读不到 ≠ 不在位（v1.6.0，oracle 复核建议）。控件数读不到时直接抛异常；控件的 Tag 读不到时先记下继续扫，
# 没找到带标签控件才抛异常——都交给主循环的「附着失败（退避重试）」，不能当成「0 个控件 / 无标签」返回 $false，
# 否则已在运行的会话会被当成未装载再装载、初始化一次；而遇到第一个读不到的控件就抛，菜单明明在位也会永远装载不上
$threw = $false; $rMe = 'unset'
try { $rMe = Menu-Exists (New-FakeCellBar @('', 'AttachLinkCtx', '') 0 0) } catch { $threw = $true }
Check (-not $threw -and $rMe -eq $true) 'Menu-Exists：带标签控件在位 → $true'
$threw = $false; $rMe = 'unset'
try { $rMe = Menu-Exists (New-FakeCellBar @('', 'Other', '') 0 0) } catch { $threw = $true }
Check (-not $threw -and $rMe -eq $false) 'Menu-Exists：没有带标签的控件 → $false'
$threw = $false; $rMe = 'unset'
try { $rMe = Menu-Exists (New-FakeCellBar @('', 'Other', '') 0 0 -CountNull) } catch { $threw = $true }
Check ($threw -and $rMe -eq 'unset') 'Menu-Exists：控件数读不到（空值）→ 抛异常，不当成 0 个控件返回 $false'
$threw = $false; $rMe = 'unset'
try { $rMe = Menu-Exists (New-FakeCellBar @('', 'Other', 'x') 2 0) } catch { $threw = $true }
Check ($threw -and $rMe -eq 'unset') 'Menu-Exists：某个控件的 Tag 读不到（空值）且没有带标签控件 → 抛异常，不当成无标签放过'
$threw = $false; $rMe = 'unset'
try { $rMe = Menu-Exists (New-FakeCellBar @('Other', 'x', 'AttachLinkCtx') 2 0) } catch { $threw = $true }
Check (-not $threw -and $rMe -eq $true) 'Menu-Exists：读不到的控件排在带标签控件之前、菜单在位 → 仍判在位（扫完再判，不因个别控件读不到而永远装载不上；旧的「遇到即抛」实现在此会抛异常）'
$threw = $false; $rMe = 'unset'
try { $rMe = Menu-Exists ([pscustomobject]@{}) } catch { $threw = $true }
Check ($threw -and $rMe -eq 'unset') 'Menu-Exists：命令栏整体取不到 → 异常原样抛给主循环'

# Get-LoadedAddInPath 三态
$p = Get-LoadedAddInPath ([pscustomobject]@{}) 'AttachLink.xlam'
Check ($null -eq $p) 'Get-LoadedAddInPath：对象不支持 Item（查询出错）→ 返回 $null'
Check ($script:lookupState -eq 'error') 'Get-LoadedAddInPath：查询出错 → lookupState=error（不是 notfound）'

# ---------------------------------------------------------------------
'=== A2. 隐藏文档窗口状态机（假状态 + 受控时钟，不依赖 Excel）==='
# 涉及「窗口被永久跳过」的路径必须在这里有用例（三级复验指出的两条路径都在此回归）：
#   ① Window.Visible 读不到时不能被判成隐藏——隐藏的自动化实例里 Win32 永远不可见，判下去就永久跳过；
#   ② 重新评估的次数只在真实的「不可见 → 可见」跳变时消耗，用完之后仍有低频重查，不是终态。
$t0 = Get-Date                                      # 受控时钟的起点：各用例用 At <秒> 取时刻
function At([int]$s) { return $script:t0.AddSeconds($s) }

$st = New-WindowState
$d = @((Get-HiddenDecision $st -1 $false (At 0) 2), (Get-HiddenDecision $st -1 $false (At 2) 2), (Get-HiddenDecision $st -1 $false (At 4) 2))
Check (($d -join ',') -eq 'wait,wait,wait' -and -not $st.done) 'A2 读不到 Window.Visible（Win32 也不可见）：连续多个周期都是「等一下再看」，没有被判成隐藏窗口'
Check ((Get-HiddenDecision $st 1 $false (At 6) 2) -eq 'process') 'A2 读不到之后恢复为可见：按工作窗口处理'

$st = New-WindowState
Check ((Get-HiddenDecision $st 0 $false (At 0) 2) -eq 'wait') 'A2 隐藏窗口首见：宽限一个周期'
Check ((Get-HiddenDecision $st 0 $false (At 1) 2) -eq 'wait') 'A2 宽限期未满：继续等'
Check ((Get-HiddenDecision $st 0 $false (At 2) 2) -eq 'skip' -and $st.done) 'A2 宽限期满仍是隐藏：判定为「不作为处理对象」'

$st = New-WindowState
$null = Get-HiddenDecision $st 0 $false (At 0) 2; $null = Get-HiddenDecision $st 0 $false (At 2) 2
Check (-not (Test-HiddenReevaluate $st $false (At 4) 3 30)) 'A2 Win32 仍不可见：不重新评估（加载项自身的窗口、不可见实例里的隐藏窗口不会被反复打扰）'
Check (Test-HiddenReevaluate $st $true (At 6) 3 30) 'A2 Win32 由不可见变为可见：触发重新评估'
Check ((Get-HiddenDecision $st 0 $true (At 6) 2) -eq 'wait') 'A2 重新评估先重新进入宽限，不立刻下结论'
Check ((Get-HiddenDecision $st 0 $true (At 8) 2) -eq 'skip' -and $st.done -and $st.rev -eq 1) 'A2 重新评估后仍是隐藏：再次判为不作为处理对象（跳变计数 1）'
Check (-not (Test-HiddenReevaluate $st $true (At 10) 3 30)) 'A2 Win32 持续可见、没有新跳变：不重复触发，也不消耗计数'
Check (-not (Test-HiddenReevaluate $st $true (At 37) 3 30)) 'A2 距上次评估不足 30 秒：不做低频重查'
Check (Test-HiddenReevaluate $st $true (At 39) 3 30) 'A2「Win32 可见、Excel 却报告隐藏」持续：至少每 30 秒做一次低频重查（不设次数上限）'
Check ($st.rev -eq 1 -and -not $st.done) 'A2 低频重查不消耗跳变计数，并把状态复位为待评估'
Check ((Get-HiddenDecision $st 1 $true (At 39) 2) -eq 'process') 'A2 低频重查时读到可见：按工作窗口处理'

# 连续三次「跳变后仍隐藏」把计数用完，之后才真正显示：跳变不再立即触发，但低频重查会接住，窗口不会被永久跳过
$st = New-WindowState
$null = Get-HiddenDecision $st 0 $false (At 0) 2; $null = Get-HiddenDecision $st 0 $false (At 2) 2
$clock = 4; $allTriggered = $true
for ($i = 1; $i -le 3; $i++) {
    if (-not (Test-HiddenReevaluate $st $true (At $clock) 3 30)) { $allTriggered = $false }
    $null = Get-HiddenDecision $st 0 $true (At $clock) 2
    $null = Get-HiddenDecision $st 0 $true (At ($clock + 2)) 2      # 仍隐藏 → 上次评估时刻 = clock + 2
    $null = Test-HiddenReevaluate $st $false (At ($clock + 3)) 3 30  # 又变回不可见（不触发）
    $clock += 4
}
Check ($allTriggered -and $st.rev -eq 3 -and $st.done) 'A2 连续三次跳变后仍隐藏：跳变计数用完（3）'
Check (-not (Test-HiddenReevaluate $st $true (At 16) 3 30)) 'A2 计数用完后的新跳变不再立即触发（距上次评估也不足 30 秒）'
Check (-not (Test-HiddenReevaluate $st $false (At 17) 3 30)) 'A2 再次变回不可见：什么都不做'
Check (Test-HiddenReevaluate $st $true (At 44) 3 30) 'A2 计数用完之后：Win32 可见的窗口至少每 30 秒被低频重查一次——不是终态'
Check ($st.rev -eq 3) 'A2 低频重查恰逢新跳变时也不越过计数上限（仍为 3）'
Check ((Get-HiddenDecision $st 1 $true (At 44) 2) -eq 'process') 'A2 低频重查读到窗口已真正显示：按工作窗口处理（不会被永久跳过）'

# 重新评估的过程中读不到
$st = New-WindowState
$null = Get-HiddenDecision $st 0 $false (At 0) 2; $null = Get-HiddenDecision $st 0 $false (At 2) 2
$null = Test-HiddenReevaluate $st $true (At 4) 3 30
Check ((Get-HiddenDecision $st -1 $true (At 4) 2) -eq 'wait' -and -not $st.done) 'A2 重新评估时读不到：只等待重试，不会被判成隐藏'
Check ((Get-HiddenDecision $st 0 $true (At 6) 2) -eq 'skip') 'A2 重新评估时读不到之后读到隐藏：照常判定为不作为处理对象'

# ---------------------------------------------------------------------
'=== C. 安装脚本宿主识别（fail-closed）==='
$hostText = Get-FunctionsText $installPath @('Test-IsNotRealExcel')
Check ($hostText.Contains('function Test-IsNotRealExcel')) '取到 install.ps1 的 Test-IsNotRealExcel'
. ([scriptblock]::Create($hostText))
Check (Test-IsNotRealExcel ([pscustomobject]@{ Path = 'D:\Software\Kingsoft\WPS Office\12.1.0.29690\office6' })) 'WPS 安装路径（含 kingsoft）判为「不是 Excel」'
Check (Test-IsNotRealExcel ([pscustomobject]@{ Path = 'D:\WPS Office\WPS Office\12.1.0.29690\office6' })) 'WPS 安装路径（仅含 wps）判为「不是 Excel」'
Check (-not (Test-IsNotRealExcel ([pscustomobject]@{ Path = 'C:\Program Files\Microsoft Office\Root\Office16' }))) '标准 Excel 路径判为 Excel'
Check (-not (Test-IsNotRealExcel ([pscustomobject]@{ Path = 'D:\MSOffice\Office16' }))) '自定义目录的 Excel 不被误判'
Check (Test-IsNotRealExcel ([pscustomobject]@{ Path = '' })) 'Path 为空 → 判为「不是 Excel」（不默认当成 Excel）'
Check (Test-IsNotRealExcel ([pscustomobject]@{})) '读不到 Path → 判为「不是 Excel」（fail-closed）'

# ---------------------------------------------------------------------
'=== F. 加载项列表登记清理（build\install.ps1，临时注册表键，不依赖 Excel）==='
# Excel 的加载项列表按【文件名】识别条目：AddIns.Add(路径) 遇到同名条目直接返回它（2026-11-24 实测）。
# 安装前必须清掉指向其他路径的 AttachLink 登记（Add-in Manager / OPENn），卸载时清掉全部——否则机器级
# 卸载留下的悬空登记会让「卸载后重装」绑定到它、Installed 置位静默失败。删除 OPEN 条目必须保持序列连续。
$regText = Get-FunctionsText $installPath @('ConvertFrom-AddInEntry', 'Test-IsAttachLinkEntry', 'Test-IsLibraryEntry', 'Test-SamePath', 'Remove-OpenEntries', 'Remove-AddInManagerEntries', 'Clear-AddInListRegistrations', 'Get-ExcelProcessIds', 'Test-ProcessAlive', 'Wait-HostExit', 'Get-NewExcelProcessIds', 'Invoke-AddInUnregister', 'Test-DisabledItemsBlob', 'Get-DisabledItemsHint')
Check ($regText.Contains('function ConvertFrom-AddInEntry') -and $regText.Contains('function Test-IsAttachLinkEntry') -and $regText.Contains('function Test-IsLibraryEntry') -and $regText.Contains('function Remove-OpenEntries') -and $regText.Contains('function Remove-AddInManagerEntries') -and $regText.Contains('function Clear-AddInListRegistrations') -and $regText.Contains('function Wait-HostExit') -and $regText.Contains('function Invoke-AddInUnregister') -and $regText.Contains('function Test-ProcessAlive') -and $regText.Contains('function Test-DisabledItemsBlob') -and $regText.Contains('function Get-DisabledItemsHint')) '取到 install.ps1 的加载项登记函数'
. ([scriptblock]::Create($regText))
$destXlam = 'C:\Users\u\AppData\Roaming\Microsoft\AddIns\AttachLink.xlam'      # install.ps1 由 $env:APPDATA 推出的库目录副本；函数调用时读取
$officeVersions = @('16.0', '15.0', '14.0')
$pfX = 'C:\Program Files\AttachLink\AttachLink.xlam'

Check ((ConvertFrom-AddInEntry '"AttachLink.xlam"') -eq 'AttachLink.xlam') 'F OPEN 值归一化：去掉引号'
Check ((ConvertFrom-AddInEntry '  "C:\Program Files\ExcelEnhancer\ExcelEnhancer.xlam" ') -eq 'C:\Program Files\ExcelEnhancer\ExcelEnhancer.xlam') 'F OPEN 值归一化：去首尾空白与引号'
Check ((ConvertFrom-AddInEntry '/A "C:\Program Files\AttachLink\AttachLink.xlam"') -eq $pfX) 'F OPEN 值归一化：去掉 /A（禁用）前缀'
Check ((Test-IsAttachLinkEntry 'AttachLink.xlam') -and (Test-IsAttachLinkEntry $pfX) -and (Test-IsAttachLinkEntry $destXlam)) 'F 库目录副本与其他路径的同名副本都识别为 AttachLink 登记'
Check ((Test-IsAttachLinkEntry 'C:\Tools\AttachLinkCopy\attachlink.xlam') -and (Test-IsAttachLinkEntry 'D:/x/y/AttachLink.xlam')) 'F 任意目录下的同名副本都识别为 AttachLink 登记（目录名不必叫 AttachLink；大小写、分隔符不敏感）'
Check (-not (Test-IsAttachLinkEntry 'C:\A\AttachLink2.xlam') -and -not (Test-IsAttachLinkEntry 'C:\A\AttachLink.xlam.bak') -and -not (Test-IsAttachLinkEntry 'C:\A\AttachLinkHelper.xlam')) 'F 名字相近的其他文件不误判'
Check (-not (Test-IsAttachLinkEntry 'C:\Program Files\ExcelEnhancer\ExcelEnhancer.xlam')) 'F 第三方加载项不误判'
Check (-not (Test-IsAttachLinkEntry '/R "C:\A\AuditCatAttachment.xll"')) 'F 第三方 XLL（带 /R 前缀的原始值）不误判'
Check ((Test-IsLibraryEntry 'attachlink.xlam') -and (Test-IsLibraryEntry $destXlam) -and -not (Test-IsLibraryEntry $pfX)) 'F 库目录登记：文件名 / 库目录完整路径算，Program Files 副本不算'

$regBase = 'HKCU:\Software\AttachLinkUnitsTest'
Remove-Item $regBase -Recurse -Force -ErrorAction SilentlyContinue
function Set-OpenSeq([string]$Key, [object[]]$Pairs) {
    # 按给定顺序写 OPEN 序列（每项为 @(名, 值)）；先清空键，每个用例都从干净状态开始
    Remove-Item $Key -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -Path $Key -Force | Out-Null
    foreach ($p in $Pairs) { New-ItemProperty -LiteralPath $Key -Name $p[0] -Value $p[1] -PropertyType String -Force | Out-Null }
}
function Get-OpenSeq([string]$Key) {
    # 现存 OPEN 序列（按序号），输出 "名=[值]"，分号连接
    $items = @((Get-ItemProperty -LiteralPath $Key).PSObject.Properties |
        Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] } |
        Sort-Object { [int]($_.Name -replace '\D', '') } | ForEach-Object { '{0}=[{1}]' -f $_.Name, $_.Value })
    return ($items -join ';')
}
$keepInstall = { param($e) (Test-IsAttachLinkEntry $e) -and -not (Test-IsLibraryEntry $e) }   # 安装：只删其他路径的副本
$removeAll = { param($e) Test-IsAttachLinkEntry $e }                                         # 卸载：全删
$third1 = '"C:\Program Files\ExcelEnhancer\ExcelEnhancer.xlam"'
$third2 = '/R "C:\Users\u\AppData\Local\AuditCat\AttachmentManager\AuditCatAttachment.xll"'
$optKey = Join-Path $regBase 'Options'

# 「卸载后重装失败」现场：库目录副本已登记，列表里另有指向 Program Files（悬空）与早期子目录副本的登记
Set-OpenSeq $optKey @(@('OPEN', $third1), @('OPEN1', $third2), @('OPEN2', '"AttachLink.xlam"'), @('OPEN3', ('"{0}"' -f $pfX)), @('OPEN4', '/A "C:\Old\AttachLink\AttachLink.xlam"'))
$r = Remove-OpenEntries $optKey $keepInstall
Check ($r.Removed -eq 2 -and -not $r.Failed) ('F 安装前清理：删掉 2 条其他路径的登记；实际 Removed=' + $r.Removed + ' Failed=' + $r.Failed)
Check ((Get-OpenSeq $optKey) -eq ('OPEN=[{0}];OPEN1=[{1}];OPEN2=["AttachLink.xlam"]' -f $third1, $third2)) ('F 安装前清理：库目录登记保留，序列重排连续，第三方条目原值不动；实际 ' + (Get-OpenSeq $optKey))

$r = Remove-OpenEntries $optKey $removeAll
Check ($r.Removed -eq 1 -and -not $r.Failed) 'F 卸载：库目录登记一并删除'
Check ((Get-OpenSeq $optKey) -eq ('OPEN=[{0}];OPEN1=[{1}]' -f $third1, $third2)) ('F 卸载后只剩第三方条目（引号、/R 前缀原样保留）且序列连续；实际 ' + (Get-OpenSeq $optKey))

# 没有命中的条目：不改任何值名（包括原有的缺口——不删除就不重排）
Set-OpenSeq $optKey @(@('OPEN', $third1), @('OPEN2', $third2))
$before = Get-OpenSeq $optKey
$r = Remove-OpenEntries $optKey $removeAll
Check ($r.Removed -eq 0 -and -not $r.Failed -and ((Get-OpenSeq $optKey) -eq $before)) 'F 没有命中条目：不删除、不重排（原有缺口保持原样）'

Set-OpenSeq $optKey @(@('OPEN', '"AttachLink.xlam"'), @('OPEN1', $third1))
$r = Remove-OpenEntries $optKey $removeAll
Check ($r.Removed -eq 1 -and -not $r.Failed -and ((Get-OpenSeq $optKey) -eq ('OPEN=[' + $third1 + ']'))) ('F 卸载：删首项后第三方条目前移到 OPEN；实际 ' + (Get-OpenSeq $optKey))

Set-OpenSeq $optKey @(@('OPEN', ('"{0}"' -f $pfX)), @('OPEN1', '"AttachLink.xlam"'))
$r = Remove-OpenEntries $optKey $removeAll
Check ($r.Removed -eq 2 -and -not $r.Failed -and ((Get-OpenSeq $optKey) -eq '')) 'F 卸载：序列里只有 AttachLink 时全部删除，不留空壳'

Set-OpenSeq $optKey @(@('OPEN', ('"{0}"' -f $destXlam)), @('OPEN1', '"AttachLink.xlam"'))
$r = Remove-OpenEntries $optKey $keepInstall
Check ($r.Removed -eq 0 -and -not $r.Failed) 'F 安装前清理：库目录副本按完整路径登记时同样保留'

# Excel 实际写出的形态（2026-11-24 实测）：非库目录副本经 AddIns.Add + Installed 后，退出时写入带引号的完整路径
Set-OpenSeq $optKey @(@('OPEN', $third1), @('OPEN1', $third2), @('OPEN2', '"C:\Tools\AttachLinkCopy\AttachLink.xlam"'))
$r = Remove-OpenEntries $optKey $keepInstall
Check ($r.Removed -eq 1 -and -not $r.Failed -and ((Get-OpenSeq $optKey) -eq ('OPEN=[{0}];OPEN1=[{1}]' -f $third1, $third2))) ('F 安装前清理：Excel 写出的「其他路径完整路径」登记被删除，第三方条目前移；实际 ' + (Get-OpenSeq $optKey))

# Add-in Manager：值名即加载项完整路径
$amKey = Join-Path $regBase 'Add-in Manager'
$other = 'C:\Users\u\AppData\Local\AuditCat\AttachmentManager\active.xlam'
New-Item -Path $amKey -Force | Out-Null
foreach ($n in @($other, $pfX, $destXlam)) { New-ItemProperty -LiteralPath $amKey -Name $n -Value '' -PropertyType String -Force | Out-Null }
$am = Remove-AddInManagerEntries $amKey $keepInstall
Check ($am.Gone.Count -eq 1 -and $am.Gone[0] -eq $pfX -and -not $am.Failed) ('F Add-in Manager：安装只删其他路径的副本；实际删除 ' + ($am.Gone -join ' | '))
Check ((((Get-Item -LiteralPath $amKey).GetValueNames() | Sort-Object) -join '|') -eq ((@($other, $destXlam) | Sort-Object) -join '|')) 'F Add-in Manager：第三方登记与库目录副本的登记保留'
$am = Remove-AddInManagerEntries $amKey $removeAll
Check ($am.Gone.Count -eq 1 -and $am.Gone[0] -eq $destXlam -and -not $am.Failed) 'F Add-in Manager：卸载清掉全部 AttachLink 登记'
Check (@((Get-Item -LiteralPath $amKey).GetValueNames()).Count -eq 1 -and @((Get-Item -LiteralPath $amKey).GetValueNames())[0] -eq $other) 'F Add-in Manager：卸载后第三方登记仍在'
$am = Remove-AddInManagerEntries (Join-Path $regBase 'NoSuchKey') $removeAll
Check ($am.Gone.Count -eq 0 -and -not $am.Failed) 'F Add-in Manager：键不存在时什么都不做、不报错'

# Clear-AddInListRegistrations：安装与卸载共用的整套清理（临时配置单元 Office\16.0\Excel 下两处登记一起放）
$hive = Join-Path $regBase 'Office'
function Set-Hive([object[]]$OpenPairs, [string[]]$AmNames) {
    $o = Join-Path $hive '16.0\Excel\Options'; $a = Join-Path $hive '16.0\Excel\Add-in Manager'
    Remove-Item $hive -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $o, $a | Out-Null
    foreach ($pair in $OpenPairs) { New-ItemProperty -LiteralPath $o -Name $pair[0] -Value $pair[1] -PropertyType String -Force | Out-Null }
    foreach ($n in $AmNames) { New-ItemProperty -LiteralPath $a -Name $n -Value '' -PropertyType String -Force | Out-Null }
}
function Get-HiveState {
    $o = Join-Path $hive '16.0\Excel\Options'; $a = Join-Path $hive '16.0\Excel\Add-in Manager'
    $seq = @((Get-ItemProperty -LiteralPath $o).PSObject.Properties | Where-Object { $_.Name -match '^OPEN(\d+)?$' } | Sort-Object { [int]($_.Name -replace '\D', '') } | ForEach-Object { '{0}=[{1}]' -f $_.Name, $_.Value }) -join ';'
    return ('OPEN ' + $seq + ' | AM ' + (@((Get-Item -LiteralPath $a).GetValueNames() | Sort-Object) -join ';'))
}
$otherCopy = '"C:\Tools\AttachLinkCopy\AttachLink.xlam"'
# 安装前清理（OtherPaths）：两处的其他路径登记都清掉，库目录登记与第三方条目原样保留
Set-Hive @(@('OPEN', $third1), @('OPEN1', $otherCopy), @('OPEN2', '"AttachLink.xlam"'), @('OPEN3', $third2)) @($other, $pfX, $destXlam)
$clean = Clear-AddInListRegistrations -Mode OtherPaths -HiveRoot $hive 3>$null 6>$null
Check ($clean -eq $true) 'F Clear-AddInListRegistrations（安装前）：返回「清理干净」'
$expAm = (@($destXlam, $other) | Sort-Object) -join ';'
$expOpen = 'OPEN OPEN=[{0}];OPEN1=["AttachLink.xlam"];OPEN2=[{1}]' -f $third1, $third2
Check ((Get-HiveState) -eq ($expOpen + ' | AM ' + $expAm)) ('F Clear-AddInListRegistrations（安装前）：两处的其他路径登记都清掉、序列重排，库目录登记与第三方条目原样保留；实际 ' + (Get-HiveState))
# 卸载（All）：两处的 AttachLink 登记全清，第三方条目原值保留、序列连续
Set-Hive @(@('OPEN', '"AttachLink.xlam"'), @('OPEN1', $third1), @('OPEN2', $otherCopy), @('OPEN3', $third2)) @($other, $pfX, $destXlam)
$clean = Clear-AddInListRegistrations -Mode All -HiveRoot $hive 3>$null 6>$null
$expOpen = 'OPEN OPEN=[{0}];OPEN1=[{1}]' -f $third1, $third2
Check ($clean -eq $true -and ((Get-HiveState) -eq ($expOpen + ' | AM ' + $other))) ('F Clear-AddInListRegistrations（卸载）：两处的 AttachLink 登记全清，第三方条目原值保留、序列连续；实际 ' + (Get-HiveState))
# 没有目标登记：什么都不动（包括原有的缺口），仍返回清理干净
Set-Hive @(@('OPEN', $third1), @('OPEN2', $third2)) @($other)
$clean = Clear-AddInListRegistrations -Mode All -HiveRoot $hive 3>$null 6>$null
$expOpen = 'OPEN OPEN=[{0}];OPEN2=[{1}]' -f $third1, $third2
Check ($clean -eq $true -and ((Get-HiveState) -eq ($expOpen + ' | AM ' + $other))) 'F Clear-AddInListRegistrations：没有目标登记时不删除、不重排（原有缺口保持原样）'

# 删除失败不得报成功：用一个同名函数顶替 Remove-ItemProperty，让指定值名删不掉（Write-Error 遵循调用方的 -ErrorAction）
$script:blockName = ''
function Remove-ItemProperty {
    [CmdletBinding()] param([string]$LiteralPath, [string[]]$Name)
    if (@($Name) -contains $script:blockName) { Write-Error ("模拟删除失败：{0}" -f $Name); return }
    Microsoft.PowerShell.Management\Remove-ItemProperty @PSBoundParameters
}
Set-Hive @(@('OPEN', $third1), @('OPEN1', '"AttachLink.xlam"')) @($other, $pfX)
$script:blockName = $pfX
$clean = Clear-AddInListRegistrations -Mode All -HiveRoot $hive 3>$null 6>$null
Check ($clean -eq $false -and ((Get-Item -LiteralPath (Join-Path $hive '16.0\Excel\Add-in Manager')).GetValueNames() -contains $pfX)) 'F Clear-AddInListRegistrations：Add-in Manager 有登记删不掉 → 返回失败（安装会中止、卸载报 exit 9），终态核验看得到它'
$script:blockName = ''
Set-Hive @(@('OPEN', '"AttachLink.xlam"'), @('OPEN1', $third1), @('OPEN2', $otherCopy)) @()
$script:blockName = 'OPEN'
$r = Remove-OpenEntries (Join-Path $hive '16.0\Excel\Options') $removeAll 3>$null
$expOpen = 'OPEN OPEN=["AttachLink.xlam"];OPEN1=[{0}]' -f $third1
Check ($r.Failed -and ((Get-HiveState) -eq ($expOpen + ' | AM '))) ('F Remove-OpenEntries：一条删不掉、另一条已删 → 返回失败；已删的保持删除，其余条目按快照写回原名原值（序列没有被重排弄乱）；实际 ' + (Get-HiveState))
$script:blockName = ''
Remove-Item Function:\Remove-ItemProperty -ErrorAction SilentlyContinue

# 进程存活探针：只把「进程号已不存在」判为已退出，其余查询错误向外抛（吞掉它们就会把「查不到」当成「已退出」）
Check (Test-ProcessAlive $PID) 'F Test-ProcessAlive：自己所在的进程 → 还在'
Check (-not (Test-ProcessAlive 999999)) 'F Test-ProcessAlive：不存在的进程号 → 已退出（不抛异常）'
Check (@(Get-ExcelProcessIds).Count -eq @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count) 'F Get-ExcelProcessIds：没有 Excel 时是空数组，有则逐个列出（与 Get-Process 一致）'

# Wait-HostExit：确认不了退出就不能返回成功（卸载据此决定要不要清注册表）
$script:fakeAlive = @{}
$probe = { param($id) [bool]$script:fakeAlive[$id] }
Check (Wait-HostExit @(101, 102) -TimeoutSeconds 5 -PollMilliseconds 20 -IsAlive $probe 3>$null) 'F Wait-HostExit：要等的进程都已不在 → 确认已退出'
Check (-not (Wait-HostExit @() -TimeoutSeconds 5 -PollMilliseconds 20 -IsAlive $probe 3>$null)) 'F Wait-HostExit：没有可等的对象（没认出宿主进程）→ 不确认'
Check (-not (Wait-HostExit $null -TimeoutSeconds 5 -PollMilliseconds 20 -IsAlive $probe 3>$null)) 'F Wait-HostExit：拿不到进程列表 → 不确认'
Check (Wait-HostExit @(999999) -TimeoutSeconds 5 -PollMilliseconds 20 3>$null) 'F Wait-HostExit：默认探针——进程号已不存在 → 确认已退出'
$script:fakeAlive = @{ 101 = $true }
$sw = [Diagnostics.Stopwatch]::StartNew()
Check (-not (Wait-HostExit @(102, 101) -TimeoutSeconds 1 -PollMilliseconds 20 -IsAlive $probe 3>$null)) 'F Wait-HostExit：有进程到时没退出 → 不确认'
Check ($sw.Elapsed.TotalSeconds -ge 0.9) 'F Wait-HostExit：不确认是因为真的等到了超时'
Check (-not (Wait-HostExit @(101) -TimeoutSeconds 5 -PollMilliseconds 20 -IsAlive { param($id) throw '查询出错' } 3>$null)) 'F Wait-HostExit：查询是否存活时出错 → 不确认（fail-closed）'
$script:probeCalls = 0
$probeCounting = { param($id) $script:probeCalls++; ($script:probeCalls -le 2) }
Check (Wait-HostExit @(101) -TimeoutSeconds 5 -PollMilliseconds 20 -IsAlive $probeCounting 3>$null) 'F Wait-HostExit：前两次查询还在、第三次已退出 → 确认'
Check ($script:probeCalls -eq 3) ('F Wait-HostExit：「随后退出」确实经历了多次轮询；实际查询 ' + $script:probeCalls + ' 次')

# Get-NewExcelProcessIds：相对快照做差集；没有快照 / 枚举失败时返回 $null（做不了差集，不能当成「没有新进程」）
$script:fakeIds = [pscustomobject]@{ Calls = 1; Before = @(1); After = @(1, 2); Throw = $false; ThrowAfterFirst = $false }
Add-Member -InputObject $script:fakeIds ScriptMethod Next {
    # 第 1 次调用（创建前的快照）返回 Before，之后返回 After；Throw / ThrowAfterFirst 模拟枚举失败
    $this.Calls++
    if ($this.Throw -or ($this.ThrowAfterFirst -and $this.Calls -gt 1)) { throw '枚举失败' }
    if ($this.Calls -eq 1) { return $this.Before }
    return $this.After
} -Force
$getIds = { $script:fakeIds.Next() }
$script:fakeIds.Calls = 1          # 下一次调用相当于「创建实例之后」的那一次
$n = Get-NewExcelProcessIds @(1) $getIds 3>$null
Check ($null -ne $n -and $n.Count -eq 1 -and $n[0] -eq 2) 'F Get-NewExcelProcessIds：返回相对快照新出现的进程号'
$script:fakeIds.Calls = 1
$n = Get-NewExcelProcessIds @(1, 2) $getIds 3>$null
Check ($null -ne $n -and $n.Count -eq 0) 'F Get-NewExcelProcessIds：没有新进程时是空数组（不是 $null）'
$script:fakeIds.Calls = 1
Check ($null -eq (Get-NewExcelProcessIds $null $getIds 3>$null)) 'F Get-NewExcelProcessIds：没有创建前的快照 → $null'
$script:fakeIds.Calls = 1
$script:fakeIds.ThrowAfterFirst = $true
Check ($null -eq (Get-NewExcelProcessIds @(1) $getIds 3>$null)) 'F Get-NewExcelProcessIds：枚举失败 → $null（不能当成「没有新进程」）'

# Invoke-AddInUnregister（卸载第 1 步）：何时算「确认实例已退出」——只有两种情况为真，见函数注释
function New-FakeAddInList {
    $entry = [pscustomobject]@{ Installed = $true }
    $c = [pscustomobject]@{}
    Add-Member -InputObject $c NoteProperty Entry $entry -Force
    Add-Member -InputObject $c ScriptMethod Item { param($n) return $this.Entry } -Force
    return $c
}
function New-FakeExcelApp($Ains) {
    $a = [pscustomobject]@{}
    Add-Member -InputObject $a NoteProperty Ains $Ains -Force
    Add-Member -InputObject $a ScriptProperty AddIns { $this.Ains } -Force
    Add-Member -InputObject $a ScriptMethod Quit { $script:fakeQuit++ } -Force
    return $a
}
function Reset-UnregisterFixture([int[]]$Before, [int[]]$After) {
    $script:fakeIds.Calls = 0; $script:fakeIds.Before = $Before; $script:fakeIds.After = $After
    $script:fakeIds.Throw = $false; $script:fakeIds.ThrowAfterFirst = $false
    $script:fakeAlive = @{}; $script:fakeQuit = 0; $script:fakeLive = @{}
}
function Invoke-FakeUnregister([string]$Mode) {
    # 注意：局部变量不要叫 $create——PowerShell 变量名不分大小写，会和 [string]$Create 参数当成同一个而把脚本块转成字符串
    $make = { New-FakeExcelApp $ains }
    if ($Mode -eq 'throw') { $make = { throw '按 CLSID 激活得到的仍不是 Microsoft Excel' } }
    # $probe 供 Wait-HostExit（进程是否已退出）；$liveProbe 供 council C2 的「卸载前已存在进程是否仍存活」判定
    return (Invoke-AddInUnregister -CreateExcel $make -GetExcelIds $getIds -IsAlive $probe -IsLive $liveProbe -TimeoutSeconds 1 -PollMilliseconds 20 3>$null)
}
$liveProbe = { param($id) [bool]$script:fakeLive[$id] }
Reset-UnregisterFixture @(1) @(1, 2)
$ains = New-FakeAddInList
$r = Invoke-FakeUnregister 'ok'
Check ($r.Unregistered -and $r.ExitConfirmed -and -not $r.Error) 'F 卸载第 1 步：取得实例 → 注销 → 实例已退出 → 已注销、确认退出'
Check ($ains.Entry.Installed -eq $false -and $script:fakeQuit -eq 1 -and $script:fakeIds.Calls -eq 2) 'F 卸载第 1 步：确实置了 Installed=False、调了 Quit，并用创建前后的两次进程列表做了差集'
Reset-UnregisterFixture @(1) @(1, 2)
$script:fakeAlive = @{ 2 = $true }
$r = Invoke-FakeUnregister 'ok'
Check ($r.Unregistered -and -not $r.ExitConfirmed) 'F 卸载第 1 步：取得实例并注销了，但创建期间出现的进程没退出 → 不确认'
Reset-UnregisterFixture @(1) @(1, 3)
$script:fakeAlive = @{ 3 = $true }
$r = Invoke-FakeUnregister 'throw'
Check (-not $r.Unregistered -and -not $r.ExitConfirmed -and $r.Error -like '*不是 Microsoft Excel*') 'F 卸载第 1 步：创建实例之后才失败、那个进程没退出 → 不确认（不能当成「没拉起 Excel」）'
Reset-UnregisterFixture @(1) @(1, 4)
$r = Invoke-FakeUnregister 'throw'
Check (-not $r.Unregistered -and $r.ExitConfirmed) 'F 卸载第 1 步：创建之后才失败，但那个进程随后退出 → 确认'
Reset-UnregisterFixture @(1) @(1)
$r = Invoke-FakeUnregister 'throw'
Check (-not $r.Unregistered -and $r.ExitConfirmed) 'F 卸载第 1 步：失败且创建期间没有新的 EXCEL 进程（本机没有 Excel / 落到了 WPS）→ 确认'
Reset-UnregisterFixture @() @()
$script:fakeIds.Throw = $true
$ains = New-FakeAddInList
$r = Invoke-FakeUnregister 'ok'
Check ($r.Unregistered -and -not $r.ExitConfirmed) 'F 卸载第 1 步：创建前枚举不到进程（做不了差集）→ 注销照做，但不确认'
Reset-UnregisterFixture @(1) @(1, 2)
$script:fakeIds.ThrowAfterFirst = $true
$ains = New-FakeAddInList
$r = Invoke-FakeUnregister 'ok'
Check ($r.Unregistered -and -not $r.ExitConfirmed) 'F 卸载第 1 步：创建后枚举不到进程（认不出该等谁）→ 注销照做，但不确认'
Reset-UnregisterFixture @() @()
$script:fakeIds.Throw = $true
$r = Invoke-FakeUnregister 'throw'
Check (-not $r.Unregistered -and -not $r.ExitConfirmed) 'F 卸载第 1 步：既没取得实例、又枚举不到进程 → 不确认'

# 「禁用的项目」二进制值：按完整文件名匹配（值内是 UTF-16 文本：12 字节头 + 路径 + NUL [+ ProgID + NUL]）
function New-DisabledBlob([string[]]$Parts) {
    $ms = New-Object System.IO.MemoryStream
    $ms.Write([byte[]](1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), 0, 12)
    foreach ($s in $Parts) { $b = [Text.Encoding]::Unicode.GetBytes($s + [char]0); $ms.Write($b, 0, $b.Length) }
    return $ms.ToArray()
}
Check (Test-DisabledItemsBlob (New-DisabledBlob @('c:\users\u\appdata\roaming\microsoft\addins\attachlink.xlam'))) 'F 禁用项目：列有 AttachLink.xlam → 识别'
Check (Test-DisabledItemsBlob (New-DisabledBlob @('c:\program files\attachlink\attachlink.xlam', 'attachlink.connect'))) 'F 禁用项目：路径后面还跟着 ProgID 也识别'
Check (-not (Test-DisabledItemsBlob (New-DisabledBlob @('c:\users\u\appdata\roaming\microsoft\addins\attachlinkshim.dll', 'attachlinkshim.connect')))) 'F 禁用项目：引导器 AttachLinkShim.dll 的旧记录不算'
Check (-not (Test-DisabledItemsBlob (New-DisabledBlob @('c:\x\myattachlink.xlam')))) 'F 禁用项目：前缀相近的文件名不算'
Check (-not (Test-DisabledItemsBlob (New-DisabledBlob @('c:\x\attachlink.xlam.bak')))) 'F 禁用项目：后缀相近的文件名不算'
Check (-not (Test-DisabledItemsBlob ([byte[]]@())) -and -not (Test-DisabledItemsBlob $null)) 'F 禁用项目：空值 / 空引用不算'

$hint = Get-DisabledItemsHint
Check ($hint -is [string]) 'F Get-DisabledItemsHint：读取「禁用的项目」不抛异常，返回字符串'
Remove-Item $regBase -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------
'=== F2. 卸载预检 / 收尾核验 / 退出码合成 / 提权账户（假对象 + 临时位置，不依赖 Excel）==='
# council C1/C2/C4 回归：「卸载不干净却显示卸载完成」三连 + 提权账户假报「已部署」。
# install.ps1 / install-elevated.ps1 的新逻辑全部是可注入假对象 / 临时注册表键的函数。

# F2-1 卸载预检的文件占用判定
$f2dir = Join-Path $base 'f2'
New-Item -ItemType Directory -Force -Path $f2dir | Out-Null
$preText = Get-FunctionsText $installPath @('Test-AddInFileLocked')
Check ($preText.Contains('function Test-AddInFileLocked')) 'F2 取到 install.ps1 的预检函数'
. ([scriptblock]::Create($preText))
$lockFile = Join-Path $f2dir 'AttachLink.xlam'
[IO.File]::WriteAllText($lockFile, 'x')
Check (-not (Test-AddInFileLocked $lockFile)) 'F2 预检：文件存在且未被占用 → 不阻碍卸载'
Check (-not (Test-AddInFileLocked (Join-Path $f2dir 'no-such.xlam'))) 'F2 预检：文件不存在 → 不算占用'
$fsLock = [IO.File]::Open($lockFile, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
try {
    Check (Test-AddInFileLocked $lockFile) 'F2 预检：有「拒绝共享」的打开方在场 → 判定占用（卸载不得继续）'
} finally { $fsLock.Close() }

# F2-2 council C2：Invoke-AddInUnregister 追加「卸载前已存在的进程仍存活 → 不能确认退出」
#   探针①（复核者已复现）：旧 PID 111 一直活着、新 PID 222 已退出——原实现误报 ExitConfirmed=true
Reset-UnregisterFixture @(111) @(111, 222)
$ains = New-FakeAddInList
$script:fakeLive = @{ 111 = $true }
$r = Invoke-FakeUnregister 'ok'
Check ($r.Unregistered -and -not $r.ExitConfirmed) 'F2 注销确认：旧 PID 仍存活、新 PID 已退出 → 不确认（不得假确认）'
Check ($r.Error -like '*仍在运行*') 'F2 注销确认：失败原因写明「仍在运行」'
#   探针②（复核者已复现）：创建失败且没有新 PID、旧 PID 111 活着——原实现也返回 true
Reset-UnregisterFixture @(111) @(111)
$script:fakeLive = @{ 111 = $true }
$r = Invoke-FakeUnregister 'throw'
Check (-not $r.Unregistered -and -not $r.ExitConfirmed) 'F2 注销确认：创建失败、无新 PID，但旧 PID 仍存活 → 不确认'
#   反向：旧 PID 是墓碑残影（IsLive=false）→ 不拦确认
Reset-UnregisterFixture @(111) @(111, 222)
$ains = New-FakeAddInList
$script:fakeLive = @{ 111 = $false }
$r = Invoke-FakeUnregister 'ok'
Check ($r.Unregistered -and $r.ExitConfirmed) 'F2 注销确认：旧 PID 是墓碑残影（已退出）→ 照常确认'
#   IsLive 查询出错 → fail-closed，不确认
Reset-UnregisterFixture @(111) @(111, 222)
$ains = New-FakeAddInList
$liveProbeBoom = { param($id) throw '查询出错' }
$r = Invoke-AddInUnregister -CreateExcel { New-FakeExcelApp $ains } -GetExcelIds $getIds -IsAlive $probe -IsLive $liveProbeBoom -TimeoutSeconds 1 -PollMilliseconds 20 3>$null
Check ($r.Unregistered -and -not $r.ExitConfirmed) 'F2 注销确认：存活状态查询出错 → 不确认（fail-closed）'

# F2-3 用户级卸载收尾核验（Get-UninstallResidue：xlam / 旧子目录 / 全版本受信任位置 / 旧引导器 / 存活 Excel）
$regText2 = Get-FunctionsText $installPath @('Get-UninstallResidue')
Check ($regText2.Contains('function Get-UninstallResidue') -and $regText2.Contains('AddinsData\AttachLinkShim.Connect')) 'F2 取到残留核验函数，且覆盖旧 COM 缓存键 AddinsData\AttachLinkShim.Connect'
. ([scriptblock]::Create($regText2))
$resBase = Join-Path $f2dir 'res'
New-Item -ItemType Directory -Force -Path $resBase, (Join-Path $resBase 'AttachLink') | Out-Null
[IO.File]::WriteAllText((Join-Path $resBase 'AttachLink.xlam'), 'x')
[IO.File]::WriteAllText((Join-Path $resBase 'AttachLinkShim.dll'), 'x')
$rhive = Join-Path $regBase 'OfficeRes'
New-Item -Path "$rhive\15.0\Excel\Security\Trusted Locations\AttachLink" -Force | Out-Null
$noneAlive = { @() }
# ComShimKeys 显式传空：本机真实存在 AddinsData\AttachLinkShim.Connect（spec 实证），
# 单测只看注入的临时键，不读真实 Excel 键
$res = @(Get-UninstallResidue -ListCleaned $true -XlamPath (Join-Path $resBase 'AttachLink.xlam') `
    -LegacyDir (Join-Path $resBase 'AttachLink') -HiveRoot $rhive -ComShimKeys @() -GetLiveExcel $noneAlive)
Check ($res.Count -eq 4) ('F2 残留核验：副本 + 旧子目录 + 15.0 信任位置 + 旧引导器 DLL 共 4 项；实际 ' + $res.Count)
Check (($res | Where-Object { $_ -like '加载项副本*' }).Count -eq 1) 'F2 残留核验：加载项副本残留被列出'
Check (($res | Where-Object { $_ -like '旧子目录*' }).Count -eq 1) 'F2 残留核验：旧子目录残留被列出'
Check (($res | Where-Object { $_ -like '受信任位置登记*15.0*Security*' }).Count -eq 1) 'F2 残留核验：另一个 Office 版本的受信任位置残留被列出'
Check (($res | Where-Object { $_ -like '旧 COM 引导器文件*AttachLinkShim.dll*' }).Count -eq 1) 'F2 残留核验：AddIns 目录下的旧引导器 DLL 被列出'
# 旧 COM 缓存键（AddinsData）走注入的临时键验证
$shimKey = Join-Path $regBase 'ShimCache\AttachLinkShim.Connect'
New-Item -Path $shimKey -Force | Out-Null
$res = @(Get-UninstallResidue -ListCleaned $true -XlamPath (Join-Path $f2dir 'ghost.xlam') `
    -LegacyDir (Join-Path $f2dir 'ghostdir') -HiveRoot (Join-Path $regBase 'NoSuchHive') `
    -ComShimKeys @($shimKey) -GetLiveExcel $noneAlive)
Check ($res.Count -eq 1 -and $res[0] -like '旧 COM 引导器登记*') 'F2 残留核验：旧 COM 缓存键残留被列出'
Remove-Item (Split-Path -Parent $shimKey) -Recurse -Force -ErrorAction SilentlyContinue
$res = @(Get-UninstallResidue -ListCleaned $false -XlamPath (Join-Path $f2dir 'ghost.xlam') `
    -LegacyDir (Join-Path $f2dir 'ghostdir') -HiveRoot (Join-Path $regBase 'NoSuchHive') -ComShimKeys @() -GetLiveExcel $noneAlive)
Check ($res.Count -eq 1 -and $res[0] -like '加载项列表登记*') 'F2 残留核验：仅登记未清净 → 只有登记一项（退出码 9 契约）'
$res = @(Get-UninstallResidue -ListCleaned $true -XlamPath (Join-Path $f2dir 'ghost.xlam') `
    -LegacyDir (Join-Path $f2dir 'ghostdir') -HiveRoot (Join-Path $regBase 'NoSuchHive') -ComShimKeys @() -GetLiveExcel { @(5, 6) })
Check ($res.Count -eq 1 -and $res[0] -like '*pid 5、6*') 'F2 残留核验：收尾时仍有存活 Excel → 列出 PID，处置为重新运行卸载'
$res = @(Get-UninstallResidue -ListCleaned $true -XlamPath (Join-Path $f2dir 'ghost.xlam') `
    -LegacyDir (Join-Path $f2dir 'ghostdir') -HiveRoot (Join-Path $regBase 'NoSuchHive') -ComShimKeys @() -GetLiveExcel $noneAlive)
Check ($res.Count -eq 0) 'F2 残留核验：全部干净 → 空清单（才允许说「卸载完成」）'
Remove-Item $resBase -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $rhive -Recurse -Force -ErrorAction SilentlyContinue

# F2-4 setup.ps1 卸载退出码合成（优先级 8 > 10 > 9 > 0）
$suText = Get-FunctionsText (Join-Path $PSScriptRoot 'setup.ps1') @('Get-UninstallExitCode')
Check ($suText.Contains('function Get-UninstallExitCode')) 'F2 取到 setup.ps1 的退出码合成函数'
. ([scriptblock]::Create($suText))
Check ((Get-UninstallExitCode $true $true $true) -eq 8) 'F2 退出码合成：机器级残留 → 8（最高优先级）'
Check ((Get-UninstallExitCode $true $false $false) -eq 8) 'F2 退出码合成：仅机器级残留 → 8'
Check ((Get-UninstallExitCode $false $true $true) -eq 10) 'F2 退出码合成：用户级残留（含登记未清净）→ 10'
Check ((Get-UninstallExitCode $false $false $true) -eq 9) 'F2 退出码合成：仅登记未清净 → 9（既有契约）'
Check ((Get-UninstallExitCode $false $false $false) -eq 0) 'F2 退出码合成：全部干净 → 0'

# F2-5 install-elevated.ps1：SID 一致判定 + 机器级卸载残留核验
$elevPath = Join-Path $repo 'build\install-elevated.ps1'
$elText = Get-FunctionsText $elevPath @('Test-SameUserSid', 'Get-MachineUninstallResidue')
Check ($elText.Contains('function Test-SameUserSid') -and $elText.Contains('function Get-MachineUninstallResidue')) 'F2 取到 install-elevated.ps1 的 SID 判定与机器级残留核验函数'
. ([scriptblock]::Create($elText))
Check (Test-SameUserSid 'S-1-5-21-1' 's-1-5-21-1') 'F2 SID 判定：同一 SID（大小写不敏感）→ 一致'
Check (-not (Test-SameUserSid 'S-1-5-21-1' 'S-1-5-21-2')) 'F2 SID 判定：不同 SID（用了别的管理员账户）→ 不一致'
Check (Test-SameUserSid '' 'S-1-5-21-1') 'F2 SID 判定：交互用户未知 → 不拦（保持现有行为，只多提示）'
$mDir = Join-Path $f2dir 'machine'
New-Item -ItemType Directory -Force -Path $mDir | Out-Null
$mHive = Join-Path $regBase 'MachineOffice'
New-Item -Path "$mHive\14.0\Excel\Security\Trusted Locations\AttachLink" -Force | Out-Null
$res = @(Get-MachineUninstallResidue -Task 'AttachLink Watchdog' -Dir $mDir -TrustedRoot $mHive `
    -GetTask { param($n) @('x') } -GetWdProcs { @([pscustomobject]@{ ProcessId = 77 }) })
Check ($res.Count -eq 4) ('F2 机器级残留核验：任务 + 目录 + 14.0 信任位置 + 看门狗进程共 4 项；实际 ' + $res.Count)
Check (($res | Where-Object { $_ -like '计划任务*' }).Count -eq 1) 'F2 机器级残留核验：计划任务仍在 → 列出'
Check (($res | Where-Object { $_ -like '部署目录*' }).Count -eq 1) 'F2 机器级残留核验：部署目录仍在 → 列出'
Check (($res | Where-Object { $_ -like '*14.0*' }).Count -eq 1) 'F2 机器级残留核验：另一版本 HKLM 信任位置仍在 → 列出'
Check (($res | Where-Object { $_ -like '看门狗进程*77*' }).Count -eq 1) 'F2 机器级残留核验：看门狗进程仍在 → 列出 PID'
$res = @(Get-MachineUninstallResidue -Task 'X' -Dir (Join-Path $f2dir 'ghost') -TrustedRoot (Join-Path $regBase 'NoSuchHive') `
    -GetTask { param($n) @() } -TestDir { param($p) $false } -GetWdProcs { @() })
Check ($res.Count -eq 0) 'F2 机器级残留核验：全部干净 → 空清单（才允许报「卸载完成」）'
Remove-Item $mDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $mHive -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $f2dir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------
# 公共：被测看门狗（独立日志）与真实 Excel 实例
$trusted = $XlamPath
if (-not (Test-Path $trusted)) { $trusted = Join-Path $env:APPDATA 'Microsoft\AddIns\AttachLink.xlam' }
function Start-TestWatchdog([string]$Log) {
    return (Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $wdPath), '-LogFile', ('"{0}"' -f $Log), '-XlamPath', ('"{0}"' -f $trusted)) -PassThru -WindowStyle Hidden)
}
function Get-WdLogLines([string]$Log) {
    # 用「允许读写共享」的方式读：被测看门狗正往日志追加的那一瞬间，按只读共享打开会撞上
    #「文件正在被另一进程使用」（2026-11-24 实测偶发，ReadAllLines 即如此）；仍读不到就按空处理，下个周期再读
    if (-not (Test-Path $Log)) { return @() }
    $fs = $null; $sr = $null
    try {
        $fs = [IO.File]::Open($Log, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8)
        $text = $sr.ReadToEnd()
        # 末尾没有换行的最后一个分段是看门狗还没写完的半行，丢弃
        if ($text.Length -gt 0 -and -not $text.EndsWith("`n")) { $text = $text.Substring(0, $text.LastIndexOf("`n") + 1) }
        return @($text -split "`r?`n" | Where-Object { $_.Length -gt 0 })
    } catch { return @() } finally {
        if ($sr) { $sr.Close() } elseif ($fs) { $fs.Close() }
    }
}
function Wait-LogLine([string]$Log, [string]$Pattern, [int]$Seconds) {
    # 等到日志里出现匹配的行数（0 = 超时）
    $deadline = (Get-Date).AddSeconds($Seconds); $n = 0
    while ((Get-Date) -lt $deadline -and $n -lt 1) {
        Start-Sleep -Milliseconds 500
        $n = @(Get-WdLogLines $Log | Where-Object { $_ -match $Pattern }).Count
    }
    return $n
}
function Get-MenuPresent($app) {
    $found = $false; $bars = $null; $bar = $null; $ctrls = $null
    try {
        $bars = $app.CommandBars; $bar = $bars.Item('Cell'); $ctrls = $bar.Controls
        $n = [int]$ctrls.Count
        for ($i = 1; $i -le $n; $i++) {
            $c = $null
            try { $c = $ctrls.Item($i); if ($c.Tag -eq 'AttachLinkCtx') { $found = $true } } finally { Release-Com $c }
            if ($found) { break }
        }
    } finally { Release-Com $ctrls; Release-Com $bar; Release-Com $bars }
    return $found
}
function Test-TrustedLoaded($wbs) {
    # 宿主里是否已有名为 AttachLink.xlam 的工作簿在位（按名字取，读完即释放）
    $w = $null
    try { $w = $wbs.Item('AttachLink.xlam'); return $true } catch { return $false } finally { Release-Com $w }
}
function Close-Instance($xl, $wbs) {
    # 标记工作簿为已保存（免得 Quit 时弹保存提示）→ Quit → 显式释放（循环到计数归零）；不依赖 GC，避免把实例钉住
    try {
        $n = [int]$wbs.Count
        for ($i = 1; $i -le $n; $i++) { $wi = $null; try { $wi = $wbs.Item($i); $wi.Saved = $true } catch {} finally { Release-Com $wi } }
    } catch {}
    try { $xl.Quit() } catch {}
    Release-Com $wbs; Release-Com $xl
}
function Wait-OwnExit([uint32]$ProcId, [int]$Seconds) {
    # 等自己拉起的实例自行退出（引用已显式释放，通常 1~3 秒）；超时留给调用方按号清理
    if ($ProcId -eq 0) { return }
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline -and (Get-Process -Id $ProcId -ErrorAction SilentlyContinue)) { Start-Sleep -Seconds 1 }
}
function Stop-OwnProcess([uint32]$ProcId, [string]$What) {
    if ($ProcId -gt 0 -and (Get-Process -Id $ProcId -ErrorAction SilentlyContinue)) {
        try { Stop-Process -Id $ProcId -Force } catch {}
        "    已清理$What pid=$ProcId"
    }
}

# ---------------------------------------------------------------------
'=== B. 同名异路径副本（真实 Excel）：不打开第二份、不执行其宏 ==='
$foreign = Join-Path $base 'foreign\AttachLink.xlam'
$helperPid = [uint32]0; $testPidB = [uint32]0; $wd = $null

# 造一份「没有宏」的同名加载项：另起一个短命的隐藏实例另存为 xlam（55 = xlOpenXMLAddIn），里面没有任何 VBA
$x = New-Object -ComObject Excel.Application
$xwbs = $null; $xwb = $null
try {
    [void][AttachLinkUnitsPid]::GetWindowThreadProcessId([IntPtr][long]$x.Hwnd, [ref]$helperPid)
    $xwbs = $x.Workbooks
    $xwb = $xwbs.Add()
    $xwb.SaveAs($foreign, 55)
} catch { Check $false "构造无宏同名副本失败: $($_.Exception.Message.Split("`n")[0])" }
finally {
    try { $x.Quit() } catch {}
    Release-Com $xwb; Release-Com $xwbs; Release-Com $x
    foreach ($n in 'x', 'xwbs', 'xwb') { Remove-Variable $n -ErrorAction SilentlyContinue }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}
Wait-OwnExit $helperPid 10
Stop-OwnProcess $helperPid '构造副本用的 Excel'
Start-Sleep -Seconds 2

if (-not (Test-Path $foreign)) {
    Check $false '无宏同名副本未生成，B 段未执行'
} elseif ($deployedDiffers) {
    '    SKIP B 段：机器级常驻看门狗与被测脚本版本不一致，它会对同一实例做自己的动作，结果无法归因（请先重新部署）'
} else {
    $xl = $null; $wbs = $null; $log = Join-Path $base 'wd-b.log'
    try {
        $xl = New-Object -ComObject Excel.Application           # 隐藏实例
        [void][AttachLinkUnitsPid]::GetWindowThreadProcessId([IntPtr][long]$xl.Hwnd, [ref]$testPidB)
        $wbs = $xl.Workbooks
        $null = $wbs.Add()
        $null = $wbs.Open($foreign)                             # 宿主自己装载了「别处的」AttachLink.xlam（无宏 → 没有右键菜单）
        $loadedPath = Get-LoadedAddInPath $wbs 'AttachLink.xlam'
        $menu0 = Get-MenuPresent $xl
        Check ($loadedPath -ieq $foreign) '前置：宿主里在位的是那份异路径副本'
        Check (-not $menu0) '前置：异路径副本没有宏，右键菜单不在位'
        # 直接验证按名字查找与三态（Excel 的 Workbooks 不按序号枚举加载项工作簿，按序号会找不到）
        Check ($script:lookupState -eq 'found') 'Get-LoadedAddInPath：按名字取到在位的加载项工作簿路径，lookupState=found'
        $null = Get-LoadedAddInPath $wbs 'NoSuchAddIn_xyz.xlam'
        Check ($script:lookupState -eq 'notfound') 'Get-LoadedAddInPath：确实不在位时 lookupState=notfound（真实 Excel 的无效索引错误）'

        $wd = Start-TestWatchdog $log
        $skipPat = '同名异路径副本已装载且未初始化.*pid=' + $testPidB + '\b'
        $skipped = Wait-LogLine $log $skipPat 40
        $lines = Get-WdLogLines $log
        $mine = @($lines | Where-Object { $_ -match ('pid=' + $testPidB + '\b') })
        if ($skipped -ge 1) { 'pass: 看门狗记录「同名异路径副本…跳过」且带本实例 pid' } else { Check $false '看门狗未记录「同名异路径副本…跳过」'; $mine | ForEach-Object { '    ' + $_ } }
        Check (@($mine | Where-Object { $_ -match '已装载并初始化' }).Count -eq 0) '未对本实例记录「已装载并初始化」（没有执行初始化宏）'
        Check (@($lines | Where-Object { $_ -match 'Run AttachLink_Initialize|装载失败|复核不符' }).Count -eq 0) '日志中没有执行宏 / 装载失败 / 复核不符的痕迹'

        # 宿主侧事实核对：仍然只有那一份异路径副本在位，受信任副本没有被打开，菜单仍不在位
        $afterPath = Get-LoadedAddInPath $wbs 'AttachLink.xlam'
        Check ($afterPath -ieq $foreign) '受信任副本未被打开 / 未替换掉异路径副本（在位的仍是那份异路径副本）'
        Check (-not (Get-MenuPresent $xl)) '右键菜单仍不在位（看门狗没有对异路径副本执行初始化）'

        # 关后即退：看门狗仍在运行时，实例 Quit 后必须能退出
        $w = $null
        try { for ($i = 1; $i -le [int]$wbs.Count; $i++) { $wi = $null; try { $wi = $wbs.Item($i); $wi.Saved = $true } finally { Release-Com $wi } } } catch {}
        Close-Instance $xl $wbs
        $xl = $null; $wbs = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
        if ($wd.HasExited) {
            Check $false '被测看门狗在退出检查前已退出，本次退出检查无效'
        } else {
            $deadline = (Get-Date).AddSeconds(20); $exited = $false
            while ((Get-Date) -lt $deadline -and -not $exited) { Start-Sleep -Seconds 2; if (-not (Get-Process -Id $testPidB -ErrorAction SilentlyContinue)) { $exited = $true } }
            if ($exited -and -not $wd.HasExited) { 'pass: 该实例 Quit 后 20 秒内退出，且被测看门狗仍在运行' }
            elseif ($exited) { Check $false 'Excel 已退出，但被测看门狗在等待期间退出，检查无效' }
            else { Check $false 'Excel 在 Quit 后 20 秒仍未退出（看门狗钉住了进程）' }
        }
    } catch {
        Check $false "B 段执行异常: $($_.Exception.Message.Split("`n")[0])"
    } finally {
        if ($wd -and -not $wd.HasExited) { try { Stop-Process -Id $wd.Id -Force } catch {} }
        if ($null -ne $xl) { Close-Instance $xl $wbs }
        # 只结束本测试自己拉起的那个 Excel（按记录下来的进程号），不碰测试期间其他人 / 其他程序打开的 Excel
        Stop-OwnProcess $testPidB 'B 段拉起的 Excel'
    }
}

# ---------------------------------------------------------------------
'=== E. 隐藏的文档窗口不作为处理对象（真实 Excel）==='
# E1 用隐藏实例（不弹窗口）；E2 需要可见实例（窗口显示 / 隐藏才会改变 Win32 可见性，「重新评估」路径才走得到），
# 期间桌面上会出现一个空的 Excel 窗口约 15 秒。
$wd = $null; $testPidE1 = [uint32]0; $testPidE2 = [uint32]0
if ($deployedDiffers) {
    '    SKIP E 段：机器级常驻看门狗与被测脚本版本不一致，它会对同一实例做自己的动作，结果无法归因（请先重新部署）'
} else {
    $log = Join-Path $base 'wd-e.log'
    $hiddenFile = Join-Path $base 'hidden-window.xlsx'
    $pidTagE1 = ' pid=' + $testPidE1
    try {
        # ---------- E1：加载项自身的隐藏窗口被识别为隐藏窗口并跳过，不对它做任何装载动作 ----------
        $xl = $null; $wbs = $null
        try {
            $xl = New-Object -ComObject Excel.Application           # 隐藏实例：不弹窗口
            [void][AttachLinkUnitsPid]::GetWindowThreadProcessId([IntPtr][long]$xl.Hwnd, [ref]$testPidE1)
            $wbs = $xl.Workbooks
            $null = $wbs.Add()
            $before7 = @([AttachLinkUnitsPid]::Excel7($testPidE1, $false))
            Check ($before7.Count -eq 1) ('E1 前提：装载前实例里只有一个文档窗口（工作簿窗口）；实际 ' + $before7.Count + ' 个')
            $winH7 = [long]0; if ($before7.Count -ge 1) { $winH7 = [long]$before7[0] }
            $wd = Start-TestWatchdog $log
            # 装载可能由被测看门狗完成，也可能被先到一步的同版常驻看门狗抢了先（那时被测看门狗记的是「已在位，无需装载」）——
            # E1 要验证的是「加载项自身的隐藏窗口被识别并跳过」，由谁完成装载不影响结论，两种记录都接受。
            $prov = Wait-LogLine $log ('(已装载并初始化|已在位，无需装载): hwnd=' + $winH7 + ' pid=' + $testPidE1 + '\b') 30
            Start-Sleep -Seconds 8                                  # 再留约 4 个周期：加载项自身的隐藏窗口会被看到两次（宽限、判定）
            $linesE = @(Get-WdLogLines $log | Where-Object { $_ -match (' pid=' + $testPidE1 + '\b') })
            if ($prov -ge 1) { 'pass: E1 工作簿窗口已装载（菜单在位）' } else { Check $false 'E1 工作簿窗口既没有「已装载并初始化」也没有「已在位，无需装载」记录'; $linesE | ForEach-Object { '    ' + $_ } }
            $all7 = @([AttachLinkUnitsPid]::Excel7($testPidE1, $false))
            Check ($all7.Count -ge 2 -and $all7 -contains $winH7) ('E1 前提：装载后实例里除工作簿窗口外还有加载项自身的隐藏文档窗口；实际 ' + $all7.Count + ' 个')
            $skipPat = '隐藏窗口（加载项自身或被隐藏的工作簿），不作为处理对象: hwnd=(\d+) pid=' + $testPidE1 + '\b'
            $skipLines = @($linesE | Where-Object { $_ -match $skipPat })
            Check ($skipLines.Count -eq 1) ('E1 加载项自身的隐藏窗口被识别为隐藏窗口并跳过（恰好一条记录）；实际 ' + $skipLines.Count + ' 条')
            $other = @($linesE | Where-Object { $_ -match 'hwnd=' -and $_ -notmatch ('hwnd=' + $winH7 + '\b') -and $_ -notmatch $skipPat })
            Check ($other.Count -eq 0) ('E1 除「隐藏窗口…不作为处理对象」外，没有以其他窗口为对象的记录；实际 ' + $other.Count + ' 条')
            Check (@($linesE | Where-Object { $_ -match '复核不符|装载失败|附着失败|未取得 Application|未检测到菜单' }).Count -eq 0) 'E1 日志里没有复核不符 / 装载失败 / 附着失败记录'
            Check (Get-MenuPresent $xl) 'E1 右键菜单在位'
        } catch {
            Check $false "E1 执行异常: $($_.Exception.Message.Split("`n")[0])"
        } finally {
            if ($null -ne $xl) { Close-Instance $xl $wbs }
            $xl = $null; $wbs = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            Wait-OwnExit $testPidE1 10
            Stop-OwnProcess $testPidE1 'E1 拉起的 Excel'
        }

        # ---------- E2：被隐藏的工作簿窗口不触发装载；显示出来后重新评估、照常处理 ----------
        # 用「保存时就已隐藏窗口」的工作簿：打开即是隐藏窗口，不存在「先可见再隐藏」的构造竞态
        $helperPid = [uint32]0
        $x = New-Object -ComObject Excel.Application
        $xwbs = $null; $xwb = $null; $xwinC = $null; $xwin = $null
        try {
            [void][AttachLinkUnitsPid]::GetWindowThreadProcessId([IntPtr][long]$x.Hwnd, [ref]$helperPid)
            $xwbs = $x.Workbooks; $xwb = $xwbs.Add(); $xwinC = $xwb.Windows; $xwin = $xwinC.Item(1)
            $xwin.Visible = $false
            $xwb.SaveAs($hiddenFile, 51)
        } catch { Check $false "构造隐藏窗口工作簿失败: $($_.Exception.Message.Split("`n")[0])" }
        finally {
            try { $x.Quit() } catch {}
            Release-Com $xwin; Release-Com $xwinC; Release-Com $xwb; Release-Com $xwbs; Release-Com $x
            foreach ($n in 'x', 'xwbs', 'xwb', 'xwinC', 'xwin') { Remove-Variable $n -ErrorAction SilentlyContinue }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
        }
        Wait-OwnExit $helperPid 10
        Stop-OwnProcess $helperPid '构造隐藏窗口工作簿用的 Excel'
        Start-Sleep -Seconds 2

        if (-not (Test-Path $hiddenFile)) {
            Check $false '隐藏窗口工作簿未生成，E2 未执行'
        } elseif ($wd -and $wd.HasExited) {
            Check $false '被测看门狗在 E2 之前已退出'
        } else {
            $xl = $null; $wbs = $null; $wb = $null; $winC = $null; $win = $null
            try {
                $xl = New-Object -ComObject Excel.Application       # 可见实例（原因见本段开头）
                $xl.Visible = $true
                [void][AttachLinkUnitsPid]::GetWindowThreadProcessId([IntPtr][long]$xl.Hwnd, [ref]$testPidE2)
                $wbs = $xl.Workbooks
                $wb = $wbs.Open($hiddenFile)
                $winC = $wb.Windows; $win = $winC.Item(1)
                Check (-not [bool]$win.Visible) 'E2 前提：打开后工作簿窗口是隐藏的'
                $hidden7 = @([AttachLinkUnitsPid]::Excel7($testPidE2, $false))
                Check ($hidden7.Count -eq 1) ('E2 前提：实例里只有一个（隐藏的）文档窗口；实际 ' + $hidden7.Count + ' 个')
                Check (@([AttachLinkUnitsPid]::Excel7($testPidE2, $true)).Count -eq 0) 'E2 前提：没有 Win32 可见的文档窗口'
                $hiddenH7 = [long]0; if ($hidden7.Count -ge 1) { $hiddenH7 = [long]$hidden7[0] }
                $judged = Wait-LogLine $log ('隐藏窗口（加载项自身或被隐藏的工作簿），不作为处理对象: hwnd=' + $hiddenH7 + ' pid=' + $testPidE2 + '\b') 30
                Check ($judged -ge 1) 'E2 隐藏窗口宽限一个周期后被判为「不作为处理对象」'
                Check (-not (Test-TrustedLoaded $wbs)) 'E2 受信任副本未被打开'
                Check (-not (Get-MenuPresent $xl)) 'E2 右键菜单不在位（没有发生装载）'

                $win.Visible = $true                                   # 取消隐藏：窗口出现，看门狗重新评估并照常处理
                $reeval = Wait-LogLine $log ('先前隐藏的窗口已显示，重新评估: hwnd=' + $hiddenH7 + ' pid=' + $testPidE2 + '\b') 30
                Check ($reeval -ge 1) 'E2 窗口显示后看门狗记录「重新评估」'
                $prov2 = Wait-LogLine $log ('(已装载并初始化|已在位，无需装载): hwnd=' + $hiddenH7 + ' pid=' + $testPidE2 + '\b') 30   # 同 E1：两种记录都接受
                if ($prov2 -ge 1) { 'pass: E2 窗口显示出来后照常完成装载' } else { Check $false 'E2 窗口显示后没有装载记录'; Get-WdLogLines $log | Select-Object -Last 8 | ForEach-Object { '    ' + $_ } }
                Check (Get-MenuPresent $xl) 'E2 右键菜单在位'
                Release-Com $win; Release-Com $winC; Release-Com $wb; $win = $null; $winC = $null; $wb = $null
            } catch {
                Check $false "E2 执行异常: $($_.Exception.Message.Split("`n")[0])"
            } finally {
                if ($null -ne $xl) { Close-Instance $xl $wbs }
                $xl = $null; $wbs = $null
                [GC]::Collect(); [GC]::WaitForPendingFinalizers()
                Wait-OwnExit $testPidE2 10
                Stop-OwnProcess $testPidE2 'E2 拉起的 Excel'
            }
        }
    } finally {
        if ($wd -and -not $wd.HasExited) { try { Stop-Process -Id $wd.Id -Force } catch {} }
    }
}

'--- 清理 ---'
try { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue } catch {}
try { Remove-Item $regBase -Recurse -Force -ErrorAction SilentlyContinue } catch {}
"RESULT: " + $(if ($ok) { 'PASS' } else { 'FAIL' })
