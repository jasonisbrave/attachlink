# =====================================================================
# AttachLink 看门狗 / 安装脚本 —— 单元级回归（v1.4.3）
# ---------------------------------------------------------------------
# A. 看门狗装载判定（用假对象逐条覆盖，不依赖 Excel）：宿主里在位的副本是否就是看门狗配置的那一份，
#    决定「能否执行初始化宏」。只有 trusted 才允许；查询出错不得当成「不在位」去打开第二份。
#    A2 用假状态 + 受控时钟回归「隐藏文档窗口状态机」：读不到 Window.Visible 只延迟重试；重新评估只在真实的
#    「不可见 → 可见」跳变时计数，用完后仍有低频重查——任何路径都不会让窗口被永久跳过。
# B. 同名异路径副本（真实 Excel，隐藏实例）：宿主里装载了一份来自别处、没有宏的 AttachLink.xlam 时，
#    看门狗不打开第二份、不执行它的宏，该实例随后仍能正常退出。
# C. 安装脚本宿主识别（build\install.ps1 的 Test-IsNotRealExcel）：无法确认是 Microsoft Excel 时按
#    「不是 Excel」处理（fail-closed）。
# D. 路径比较与 Release-Com 的容错。
# E. 隐藏的文档窗口不作为处理对象（真实 Excel，隐藏实例）：加载项自身的隐藏窗口零处理；被隐藏的工作簿窗口
#    显示出来之后才处理。
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
$fnText = Get-FunctionsText $wdPath @('Release-Com', 'Test-SamePath', 'Get-LoadedAddInPath', 'Find-AddInIndex', 'Provision-App', 'New-WindowState', 'Get-HiddenDecision', 'Test-HiddenReevaluate')
Check ($fnText.Contains('function Release-Com') -and $fnText.Contains('function Test-SamePath') -and $fnText.Contains('function Get-LoadedAddInPath') -and $fnText.Contains('function Provision-App') -and $fnText.Contains('function New-WindowState') -and $fnText.Contains('function Get-HiddenDecision') -and $fnText.Contains('function Test-HiddenReevaluate')) '取到 watchdog.ps1 的相关函数'
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
'=== A. 看门狗装载判定（假对象，不依赖 Excel）==='
$badIndex = [int]0x8002000B      # DISP_E_BADINDEX：Excel 对按名字取不到的工作簿抛的错误（实测）
function New-FakeWorkbook([string]$Path, [switch]$PathThrows) {
    $o = [pscustomobject]@{ Path = $Path }
    if ($PathThrows) { Add-Member -InputObject $o ScriptProperty FullName { throw [System.Runtime.InteropServices.COMException]::new('路径不可读', [int]0x80004005) } -Force }
    else { Add-Member -InputObject $o ScriptProperty FullName { $this.Path } -Force }
    return $o
}
function New-FakeApp([string]$Lookup, $LoadedWorkbook, [string]$OpenedPath, [switch]$OpenedPathThrows, $AddIn) {
    # Lookup: 'found' | 'notfound' | 'error'  —— 模拟 Workbooks.Item('AttachLink.xlam') 的三种结果
    # Open 追踪：把每次 Open 的路径记到 $script:openedPaths
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
        if ($this.OpenedThrows) { return (New-FakeWorkbook '' -PathThrows) }
        return (New-FakeWorkbook $this.Opened)
    } -Force
    Add-Member -InputObject $wbs NoteProperty Lookup $Lookup -Force
    Add-Member -InputObject $wbs NoteProperty Loaded $LoadedWorkbook -Force
    Add-Member -InputObject $wbs NoteProperty Opened $OpenedPath -Force
    Add-Member -InputObject $wbs NoteProperty OpenedThrows ([bool]$OpenedPathThrows) -Force
    $app = [pscustomobject]@{}
    Add-Member -InputObject $app ScriptProperty Workbooks { $this.Wbs } -Force
    Add-Member -InputObject $app ScriptProperty AddIns { $this.Ains } -Force
    Add-Member -InputObject $app NoteProperty Wbs $wbs -Force
    Add-Member -InputObject $app NoteProperty Ains $AddIn -Force
    return $app
}
function New-FakeAddIns([string]$EntryPath, [bool]$Installed) {
    # AddIns 集合：至多一个名为 AttachLink.xlam 的登记项
    $entry = $null
    if ($EntryPath) { $entry = [pscustomobject]@{ Name = 'AttachLink.xlam'; FullName = $EntryPath; Installed = $Installed } }
    $ains = [pscustomobject]@{}
    Add-Member -InputObject $ains ScriptProperty Count { if ($this.Entry) { 1 } else { 0 } } -Force
    Add-Member -InputObject $ains ScriptMethod Item { param($i) if ($i -eq 1 -and $this.Entry) { return $this.Entry }; throw [System.Runtime.InteropServices.COMException]::new('无效索引。', [int]0x8002000B) } -Force
    Add-Member -InputObject $ains ScriptMethod Add { param($p, $c) $this.Entry = [pscustomobject]@{ Name = 'AttachLink.xlam'; FullName = [string]$p; Installed = $false }; return $this.Entry } -Force
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

# AddIns 登记：只对「登记的就是配置路径」的条目置 Installed=True
$ainsForeign = New-FakeAddIns 'C:\Elsewhere\AttachLink.xlam' $false
$r = Invoke-Provision (New-FakeApp 'notfound' $null $XlamPath -AddIn $ainsForeign)
Check ($r.state -eq 'trusted' -and -not $ainsForeign.Entry.Installed) '登记在别处的同名条目：不被置为 Installed（不会让 Excel 从那个路径装载）'
$ainsMine = New-FakeAddIns $XlamPath $false
$r = Invoke-Provision (New-FakeApp 'notfound' $null $XlamPath -AddIn $ainsMine)
Check ($r.state -eq 'trusted' -and $ainsMine.Entry.Installed) '登记路径就是配置路径的条目：被置为 Installed'
$ainsNone = New-FakeAddIns '' $false
$r = Invoke-Provision (New-FakeApp 'notfound' $null $XlamPath -AddIn $ainsNone)
Check ($r.state -eq 'trusted' -and $ainsNone.Entry -and $ainsNone.Entry.FullName -eq $XlamPath -and $ainsNone.Entry.Installed) '没有登记项：从配置路径新增登记并置为 Installed'

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
# 公共：被测看门狗（独立日志）与真实 Excel 实例
$trusted = $XlamPath
if (-not (Test-Path $trusted)) { $trusted = Join-Path $env:APPDATA 'Microsoft\AddIns\AttachLink.xlam' }
function Start-TestWatchdog([string]$Log) {
    return (Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $wdPath), '-LogFile', ('"{0}"' -f $Log), '-XlamPath', ('"{0}"' -f $trusted)) -PassThru -WindowStyle Hidden)
}
function Get-WdLogLines([string]$Log) {
    if (Test-Path $Log) { return @([IO.File]::ReadAllLines($Log, [Text.Encoding]::UTF8)) }
    return @()
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
"RESULT: " + $(if ($ok) { 'PASS' } else { 'FAIL' })
