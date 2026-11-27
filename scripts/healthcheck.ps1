# =====================================================================
# AttachLink —— 安装器健康检查（v1.5.0，纯只读）
# ---------------------------------------------------------------------
# 用途：安装器「健康检查」按钮与 /healthcheck 静默参数的实现本体；也可独立运行：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\healthcheck.ps1
# 检查三个层面：
#   user.*     用户级安装（加载项文件、版本一致性、受信任位置、启用登记、他路径登记、禁用的项目）
#   machine.*  机器级看门狗（部署门禁、计划任务（含登记账户核对）、任务运行结果、日志心跳、文件、版本、HKLM 信任、父进程白名单）
#   diag.*     诊断信息（看门狗日志可读性、近期错误关键字、安装器日志最近结论、OPEN 序列快照）
# 只读承诺：只读注册表 / 文件 / 计划任务元数据；不创建 COM、不提权、不启停任务、不枚举进程、
#           不写任何文件——唯一输出是 stdout 的协议行。
# 活性判定唯一依据 = 看门狗日志心跳（LastWriteTime 距今 ≤ 90 秒）。计划任务直接运行无窗口启动器
#   （AttachLink.WatchdogLauncher.exe），看门狗运行期间任务 State=Running；medium-IL 也枚举不到高完整性的
#   看门狗进程——两者都不作为依据。
# 看门狗引擎（v1.7.0 起，信息性）：默认 .NET 版 AttachLink.Watchdog.exe（依赖 .NET Framework 4.8，Win10 1903+ /
#   Win11 自带）；部署目录里没有它时启动器回退 Windows PowerShell 5.1 + watchdog.ps1。两种引擎写同一日志、
#   同一套措辞，machine.files 只在成功文案里报告当前是哪种，不是必需文件、不影响任何判定。
# 输出协议（每行一条，制表符分隔；消息内的制表符 / 换行已替换为空格）：
#   HC<TAB>id<TAB>status<TAB>message      status ∈ ok / warn / fail / skip / info
#   HCHINT<TAB>id<TAB>hint                可操作的处置建议（可选，跟随对应 HC 行）
#   HC<TAB>SUMMARY<TAB>ok=n warn=n fail=n skip=n info=n
# 退出码：0 = 无 fail；10 = 有 fail；1 = 脚本自身出错。
# 编码：本文件含中文，必须保存为 UTF-8 with BOM（WORKLOG 根因 #11）。
# =====================================================================
param(
    [string]$LogPath = (Join-Path $env:TEMP 'AttachLinkWatchdog.log'),
    [int]$HeartbeatWaitSeconds = 0,                               # 心跳陈旧 / 缺失时最多等多少秒让它刷新（安装后自动检查传 15）
    [string]$PayloadDir = '',                                     # 安装包根（含 dist\、scripts\）；默认取本脚本上一级目录
    [string]$OfficeRootKey = 'HKCU:\Software\Microsoft\Office',   # 测试可注入临时注册表根
    [string]$TaskName = 'AttachLink Watchdog',
    [string]$MachineDir = 'C:\Program Files\AttachLink',
    [string]$TaskQueryInjectError = ''                            # 测试注入：非空时模拟计划任务查询失败（错误消息即该值）
)
$ErrorActionPreference = 'Continue'

# 协议计数（点源载入时随调用方脚本作用域；Write-Hc 逐项累加）
$script:hcCounts = @{ ok = 0; warn = 0; fail = 0; skip = 0; info = 0 }

# 用户级库目录副本完整路径——Test-IsLibraryEntry 判定「库目录登记」的基准（与 install.ps1 同款脚本作用域变量）
$destXlam = Join-Path $env:APPDATA 'Microsoft\AddIns\AttachLink.xlam'

# ---------------- 协议输出 ----------------

function Write-Hc([string]$Id, [string]$Status, [string]$Message, [string]$Hint = '') {
    # 发一条检查项：HC 行必发，Hint 非空时补一条 HCHINT 行；消息内制表符 / 换行替换为空格，防止协议错位
    $m = ([string]$Message) -replace "`r?`n", ' ' -replace "`t", ' '
    $h = ([string]$Hint) -replace "`r?`n", ' ' -replace "`t", ' '
    switch ($Status) {
        'ok'   { $script:hcCounts['ok']++ }
        'warn' { $script:hcCounts['warn']++ }
        'fail' { $script:hcCounts['fail']++ }
        'skip' { $script:hcCounts['skip']++ }
        'info' { $script:hcCounts['info']++ }
    }
    Write-Output ("HC`t{0}`t{1}`t{2}" -f $Id, $Status, $m)
    if ($h) { Write-Output ("HCHINT`t{0}`t{1}" -f $Id, $h) }
}

function Write-HcSummary {
    Write-Output ("HC`tSUMMARY`tok={0} warn={1} fail={2} skip={3} info={4}" -f `
        $script:hcCounts['ok'], $script:hcCounts['warn'], $script:hcCounts['fail'], $script:hcCounts['skip'], $script:hcCounts['info'])
}

function Get-HcExitCode {
    # 有 fail → 10，否则 0（脚本自身异常走主守卫的 catch → 1）
    if ($script:hcCounts['fail'] -gt 0) { return 10 }
    return 0
}

# ---------------- 与 install.ps1 保持一致的判定函数（语义镜像，改动需两边同步）----------------

function ConvertFrom-AddInEntry([string]$raw) {
    # OPENn 值 → 加载项条目：去首尾空白与引号，去掉 /A（禁用）前缀
    $entry = $raw.Trim().Trim('"')
    if ($entry -match '^(/[Aa]|\\[Aa])\s+') { $entry = $entry.Substring(2).Trim().Trim('"') }
    return $entry
}
function Test-IsAttachLinkEntry([string]$entry) {
    # 是否 AttachLink 的登记——按文件名判断（与 Excel 识别加载项列表条目的规则一致）：
    # 库目录副本（文件名或完整路径）与任何其他路径的同名副本都算
    return [bool]($entry -match '(?i)(^|[\\/])AttachLink\.xlam$')
}
function Test-IsLibraryEntry([string]$entry) {
    # 是否库目录副本自身的登记：Excel 按库目录解析的文件名，或登记成了库目录的完整路径
    return [bool](($entry -ieq $destXlam) -or ($entry -ieq 'AttachLink.xlam'))
}
function Test-IsDisabledOpenValue([string]$raw) {
    # OPENn 原始值是否「禁用形态」（/A 前缀）——归一化后的条目看不出禁用态，必须判原始值
    return [bool]($raw.Trim() -match '^(/[Aa]|\\[Aa])\s')
}
function Test-DisabledItemsBlob([byte[]]$bytes) {
    # Excel「禁用的项目」（Resiliency\DisabledItems）的一个二进制值里是否列有 AttachLink.xlam。
    # 值内含条目路径的 UTF-16 文本，按完整文件名匹配：前后只能是路径分隔符 / 字符串结尾（NUL）——
    # AttachLinkShim.dll、MyAttachLink.xlam、AttachLink.xlam.bak 都不算。
    if ($null -eq $bytes -or $bytes.Length -eq 0) { return $false }
    try { return [bool]([Text.Encoding]::Unicode.GetString($bytes) -match '(?i)(^|[\\/])AttachLink\.xlam(?=[\x00]|$)') } catch { return $false }
}

# ---------------- 通用小函数 ----------------

function Test-SameTrustedPath([string]$a, [string]$b) {
    # 受信任位置路径比较：展开环境变量、去引号、去尾部反斜杠、大小写不敏感
    try {
        $x = [Environment]::ExpandEnvironmentVariables([string]$a).Trim().Trim('"').TrimEnd('\')
        $y = [Environment]::ExpandEnvironmentVariables([string]$b).Trim().Trim('"').TrimEnd('\')
        return [bool]($x -ieq $y)
    } catch { return $false }
}

function Get-MachineDeployState([bool]$TaskExists, [bool]$DirExists) {
    # 机器级部署形态：both = 完整；none = 未部署（不用鼎信诺时属正常）；partial = 半部署（异常）
    if ($TaskExists -and $DirExists) { return 'both' }
    if (-not $TaskExists -and -not $DirExists) { return 'none' }
    return 'partial'
}

function Get-OfficeVersions([string]$RootKey) {
    # 存在 Excel 键的 Office 主版本（16.0/15.0/14.0 顺序，首个即最高版）。
    # HKCU 一个都没有时探 HKLM——只为识别版本号；OPEN / 受信任位置等 per-user 数据仍读 HKCU 路径。
    $found = @()
    foreach ($v in @('16.0', '15.0', '14.0')) {
        if (Test-Path (Join-Path $RootKey "$v\Excel")) { $found += $v }
    }
    if ($found.Count -eq 0) {
        $hklmRoot = $RootKey -replace '(?i)^HKCU:', 'HKLM:'
        foreach ($v in @('16.0', '15.0', '14.0')) {
            if (Test-Path (Join-Path $hklmRoot "$v\Excel")) { $found += $v }
        }
    }
    return $found
}

function Get-LogAgeSeconds([string]$Path) {
    # 日志最后写入距现在的秒数；文件不存在 / 读不到返回 $null
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return $null }
        return [int](((Get-Date) - (Get-Item -LiteralPath $Path).LastWriteTime).TotalSeconds)
    } catch { return $null }
}

function Wait-Heartbeat([string]$Path, [int]$FreshSeconds, [int]$WaitSeconds) {
    # 等日志心跳变新鲜：每 1 秒看一次，新鲜即返回；超时返回最后一次观测。
    # 返回 @{ Fresh = bool; AgeSeconds = int|$null; Waited = int }
    if ($WaitSeconds -lt 0) { $WaitSeconds = 0 }
    $waited = 0
    while ($true) {
        $age = Get-LogAgeSeconds $Path
        if ($null -ne $age -and $age -le $FreshSeconds) { return @{ Fresh = $true; AgeSeconds = $age; Waited = $waited } }
        if ($waited -ge $WaitSeconds) { return @{ Fresh = $false; AgeSeconds = $age; Waited = $waited } }
        Start-Sleep -Seconds 1
        $waited++
    }
}

function Get-LogLines([string]$Path) {
    # 与 test-watchdog-units.ps1 的 Get-WdLogLines 同款容错：共享读写打开（看门狗可能正在追加），
    # 末尾没有换行的半行丢弃；读不到按空处理
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $fs = $null; $sr = $null
    try {
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8)
        $text = $sr.ReadToEnd()
        if ($text.Length -gt 0 -and -not $text.EndsWith("`n")) { $text = $text.Substring(0, $text.LastIndexOf("`n") + 1) }
        return @($text -split "`r?`n" | Where-Object { $_.Length -gt 0 })
    } catch { return @() } finally {
        if ($sr) { $sr.Close() } elseif ($fs) { $fs.Close() }
    }
}

function Get-Hash16([string]$Path) {
    # SHA-256 前 16 位。.NET 直调而非 Get-FileHash：安装器 exe 的子 PowerShell 环境可能加载不了 Utility 模块
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = [IO.File]::OpenRead($Path)
    try { return ([BitConverter]::ToString($sha.ComputeHash($fs)) -replace '-', '').Substring(0, 16) }
    finally { $fs.Close(); $sha.Dispose() }
}

function Test-TaskConfig($Task, [string]$ExpectedDir) {
    # 计划任务配置核对：未被禁用 + 最高权限 + 无窗口启动器位于部署目录（动作不带参数）+ 登录触发器 + 带重复兜底的时间触发器。
    # 返回 @{ Ok = bool; Detail = string }
    # 动作是 AttachLink.WatchdogLauncher.exe（GUI 子系统，无控制台窗口），计划任务跟踪的就是它——
    # State=Running 属正常（任务在运行 = 启动器与看门狗都在运行）；State=Disabled 才是问题（任务不会随登录拉起看门狗）
    $problems = @()
    try {
        if ("$($Task.State)" -eq 'Disabled') { $problems += '任务已被禁用（State=Disabled）' }
    } catch {}
    try {
        if ("$($Task.Principal.RunLevel)" -ne 'Highest') { $problems += "RunLevel=$($Task.Principal.RunLevel)（期望 Highest）" }
    } catch { $problems += 'Principal.RunLevel 读不到' }
    $exe = ''; $arg = ''
    try {
        $exe = [string]$Task.Actions[0].Execute
        $arg = ([string]$Task.Actions[0].Arguments).Trim()
    } catch {}
    $dirN = ''
    try { $dirN = ([string]$ExpectedDir).Trim().Trim('"').TrimEnd('\') } catch {}
    $exeOk = $false
    try {
        $exeOk = ($exe -match '(?i)AttachLink\.WatchdogLauncher\.exe$') -and
                 ($exe.Length -gt $dirN.Length + 1) -and
                 ($exe.Substring(0, $dirN.Length) -ieq $dirN) -and
                 ($exe[$dirN.Length] -eq '\')
    } catch {}
    if (-not $exeOk) { $problems += '启动动作不是预期的 AttachLink.WatchdogLauncher.exe（或部署目录不符）' }
    if ($arg) { $problems += '启动动作带了参数，与预期不符' }
    $logon = $false
    $logonDisabled = $false
    $logonRepeatOnly = $false
    $timeRepeat = $false
    try {
        foreach ($tr in @($Task.Triggers)) {
            $isLogon = (($tr.PSObject.TypeNames -join ' ') -match 'LogonTrigger')
            $trEnabled = $true
            try { if ($tr.Enabled -eq $false) { $trEnabled = $false } } catch {}
            if ($isLogon) {
                $logon = $true
                # 触发器整体 Enabled=False 时任务存在但永远不会随登录触发（任务自身的 State 可能仍是 Ready）
                if (-not $trEnabled) { $logonDisabled = $true }
            }
            # v1.6.3 起重复兜底挂在单独的时间触发器上（启动器被人结束或放弃后由下一次重复拉回）。挂在登录触发器上的
            # 重复不算：它只在那次登录触发之后才开始计时，而安装 / 升级都发生在登录之后，要等下次登录才有第一次重复
            # （实测：登记后 NextRunTime 为空）。Repetition / Interval 读不到按「没有」处理，不抛异常
            $hasRepeat = $false
            try {
                $rep = $tr.Repetition
                if ($rep -and ("$([string]$rep.Interval)").Trim()) { $hasRepeat = $true }
            } catch {}
            if ($hasRepeat) {
                if ($isLogon) { $logonRepeatOnly = $true } elseif ($trEnabled) { $timeRepeat = $true }
            }
        }
    } catch {}
    if (-not $logon) { $problems += '缺登录触发器' }
    if ($logonDisabled) { $problems += '登录触发器被禁用（Enabled=False，任务不会随登录触发）' }
    if (-not $timeRepeat) {
        if ($logonRepeatOnly) { $problems += '重复兜底只挂在登录触发器上（要等下次登录才生效），应另有一个每 5 分钟重复的时间触发器——请重新「一键安装」' }
        else { $problems += '没有带重复兜底的时间触发器（v1.6.3 起应每 5 分钟重复）' }
    }
    return @{ Ok = ($problems.Count -eq 0); Detail = ($problems -join '；') }
}

function Test-TaskOwnedByOtherUser($Task, [string]$CurrentUserSid) {
    # 计划任务登记的账户是否 ≠ 当前用户（council C4：标准用户用别的管理员账号过 UAC 部署时，
    # 任务登记在管理员名下，真正使用者登录时看门狗不启动）。任务登记的 UserId 可能是
    # "DOMAIN\user" 或 SID——统一换算成 SID 再比。任一侧读不到 → 返回 $false（不误报）。
    try {
        $uid = [string]$Task.Principal.UserId
        if (-not $uid) { return $false }
        $sid = ''
        if ($uid -match '^S-1-') {
            $sid = $uid
        } else {
            try { $sid = ([Security.Principal.NTAccount]$uid).Translate([Security.Principal.SecurityIdentifier]).Value }
            catch { return $false }
        }
        if (-not $CurrentUserSid) { return $false }
        return ($sid -ine $CurrentUserSid)
    } catch { return $false }
}

function Get-TaskResultCode($Value) {
    # 把计划任务返回码归一成 0..4294967295 的整数：CIM 给的是无符号 32 位（如 2147946720），别处可能给同值的负数或字符串。
    # 不能直接 [int] 强转——超过 2147483647 的值会抛溢出异常（2026-11-27 实测）。读不到返回 -1
    try {
        if ($null -eq $Value) { return [long]0 }
        if ($Value -is [uint32]) { return [long]$Value }
        if ($Value -is [int64]) { $v = [long]$Value }
        else { $v = [long]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture) }
    } catch { return [long]-1 }
    if ($v -lt 0) { $v += 4294967296 }
    return $v
}

function Get-TaskResultStatus($Code) {
    # 计划任务上次运行返回码：
    #   0 = 正常；267009 (0x41301) = 正在运行（任务跟踪启动器进程，看门狗常驻期被查询时即此值）；
    #   2147946720 (0x800710E0) = 每 5 分钟的重复触发发现任务已在运行、新实例被忽略——看门狗运行满一个重复间隔后查到的
    #   就是它，任务计划程序界面里显示为「操作员或管理员拒绝了请求」，属正常；其余 → warn
    $u = Get-TaskResultCode $Code
    if ($u -eq 0 -or $u -eq 267009 -or $u -eq 2147946720) { return 'ok' }
    return 'warn'
}

function Get-ReadyStateNote($LastTaskResult) {
    # machine.task 在 State=Ready（任务当前没有在运行，看门狗不在运行）时附在文案后的说明，按上次运行结果分两种口径：
    # 0 / 267009 / 2147946720（任务刚被结束，或重复触发被忽略时留下的值）→ 提示 5 分钟内重复触发会拉回；
    # 其他值 = 启动器已放弃（看门狗连续 3 次没能稳定运行），把启动器的退出码带出来并指向看门狗日志
    $u = Get-TaskResultCode $LastTaskResult
    if ($u -eq 0 -or $u -eq 267009 -or $u -eq 2147946720) {
        return '；任务当前没有在运行（看门狗不在运行）：若刚结束过任务，最多 5 分钟后重复触发会重新启动；否则请重新「一键安装」或手动启动任务'
    }
    return ("；任务当前没有在运行（看门狗不在运行）：上次运行结果 = {0}：启动器已放弃（看门狗连续 3 次没能稳定运行），请看 %TEMP%\AttachLinkWatchdog.log 里「启动器:」开头的行" -f $u)
}

# =====================================================================
# 主逻辑（点源载入时不执行——单元测试只取上面的函数）
# =====================================================================
if ($MyInvocation.InvocationName -ne '.') {
    try {
        try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
        if (-not $PayloadDir) { $PayloadDir = Split-Path -Parent $PSScriptRoot }
        try { $PayloadDir = [IO.Path]::GetFullPath($PayloadDir) } catch {}
        if ($HeartbeatWaitSeconds -lt 0) { $HeartbeatWaitSeconds = 0 }

        $addInsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
        $userXlam  = Join-Path $addInsDir 'AttachLink.xlam'
        $distXlam  = Join-Path $PayloadDir 'dist\AttachLink.xlam'
        $srcWd     = Join-Path $PayloadDir 'scripts\watchdog.ps1'

        # Office 版本：HKCU 优先，HKLM 仅用于识别版本号；$ver 为空 = 未检测到 Office
        $foundVers = @(Get-OfficeVersions $OfficeRootKey)
        $ver = ''
        if ($foundVers.Count -gt 0) { $ver = $foundVers[0] }
        $mRoot = $OfficeRootKey -replace '(?i)^HKCU:', 'HKLM:'

        # ==================== user.* ====================

        # user.file：用户级加载项文件
        if (Test-Path -LiteralPath $userXlam) {
            Write-Hc 'user.file' 'ok' "用户级加载项文件在位（$userXlam）"
        } else {
            Write-Hc 'user.file' 'fail' "用户级加载项文件缺失：$userXlam" '重新运行安装器点「一键安装」（或「更多选项 → 仅安装加载项」）'
        }

        # user.version：用户级副本 vs 安装包 dist 副本
        if (-not (Test-Path -LiteralPath $userXlam)) {
            Write-Hc 'user.version' 'skip' '用户级文件缺失，无法比对版本'
        } elseif (-not (Test-Path -LiteralPath $distXlam)) {
            Write-Hc 'user.version' 'skip' '安装包内未找到 dist\AttachLink.xlam，无法比对版本'
        } else {
            try {
                $hU = Get-Hash16 $userXlam
                $hD = Get-Hash16 $distXlam
                if ($hU -eq $hD) {
                    Write-Hc 'user.version' 'ok' "用户级副本与安装包版本一致（SHA-256 前 16 位 $hU）"
                } else {
                    Write-Hc 'user.version' 'warn' "用户级副本与安装包版本不一致（前 16 位 $hU vs $hD），请重新安装" '点「一键安装」覆盖安装即可同步版本'
                }
            } catch {
                Write-Hc 'user.version' 'warn' "版本比对失败：$($_.Exception.Message)"
            }
        }

        # user.trusted：HKCU 受信任位置指向库目录
        if (-not $ver) {
            Write-Hc 'user.trusted' 'skip' '未检测到 Office，无法检查受信任位置'
        } else {
            $tlKey = "$OfficeRootKey\$ver\Excel\Security\Trusted Locations\AttachLink"
            $tlPath = ''
            try { if (Test-Path -LiteralPath $tlKey) { $tlPath = [string](Get-ItemProperty -LiteralPath $tlKey -Name Path -ErrorAction Stop).Path } } catch {}
            if ($tlPath -and (Test-SameTrustedPath $tlPath "$addInsDir\")) {
                Write-Hc 'user.trusted' 'ok' "受信任位置已登记（Office $ver → $tlPath）"
            } else {
                Write-Hc 'user.trusted' 'fail' "受信任位置未登记或路径不符（Office $ver，期望 $addInsDir\）" '重新「一键安装」；未受信任时宏被禁用，加载项不会加载'
            }
        }

        # user.open：最高版本的 OPEN 序列里有「启用形态」的库目录登记
        if (-not $ver) {
            Write-Hc 'user.open' 'skip' '未检测到 Office，无法检查启用登记'
        } else {
            $optKey = "$OfficeRootKey\$ver\Excel\Options"
            $libEnabled = $false; $libDisabled = $false; $regNames = @()
            if (Test-Path -LiteralPath $optKey) {
                $openProps = @((Get-ItemProperty -LiteralPath $optKey).PSObject.Properties |
                    Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] })
                foreach ($p in $openProps) {
                    $entry = ConvertFrom-AddInEntry ([string]$p.Value)
                    if (-not (Test-IsAttachLinkEntry $entry)) { continue }
                    if (Test-IsLibraryEntry $entry) {
                        if (Test-IsDisabledOpenValue ([string]$p.Value)) { $libDisabled = $true }
                        else { $libEnabled = $true; $regNames += $p.Name }
                    }
                }
            }
            if ($libEnabled) {
                Write-Hc 'user.open' 'ok' "加载项已登记启用（Office $ver，$($regNames -join '、')）"
            } elseif ($libDisabled) {
                Write-Hc 'user.open' 'fail' "加载项登记被标记为禁用（/A 前缀，Office $ver）" '重新「一键安装」（安装会清掉禁用形态登记并重新启用）'
            } else {
                Write-Hc 'user.open' 'fail' "加载项未在 Excel 启用列表登记（Office $ver）" '重新「一键安装」'
            }
        }

        # user.openforeign：指向其他路径的 AttachLink 登记（悬空登记会让重装绑定失败）——扫所有在场版本的 OPEN + Add-in Manager
        $foreign = New-Object System.Collections.Generic.List[string]
        foreach ($v in $foundVers) {
            $optKeyV = "$OfficeRootKey\$v\Excel\Options"
            if (Test-Path -LiteralPath $optKeyV) {
                $openPropsV = @((Get-ItemProperty -LiteralPath $optKeyV).PSObject.Properties |
                    Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] })
                foreach ($p in $openPropsV) {
                    $entry = ConvertFrom-AddInEntry ([string]$p.Value)
                    if ((Test-IsAttachLinkEntry $entry) -and -not (Test-IsLibraryEntry $entry)) {
                        $foreign.Add("Office $v Options\$($p.Name) → $entry")
                    }
                }
            }
            $amKey = "$OfficeRootKey\$v\Excel\Add-in Manager"
            if (Test-Path -LiteralPath $amKey) {
                try {
                    foreach ($n in @((Get-Item -LiteralPath $amKey).GetValueNames())) {
                        if ((Test-IsAttachLinkEntry $n) -and -not (Test-IsLibraryEntry $n)) {
                            $foreign.Add("Office $v Add-in Manager → $n")
                        }
                    }
                } catch {}
            }
        }
        if ($foundVers.Count -eq 0) {
            Write-Hc 'user.openforeign' 'skip' '未检测到 Office，无法检查他路径登记'
        } elseif ($foreign.Count -eq 0) {
            Write-Hc 'user.openforeign' 'ok' '无指向其他路径的 AttachLink 登记'
        } else {
            Write-Hc 'user.openforeign' 'fail' ("发现指向其他路径的 AttachLink 登记：" + ($foreign -join '；')) 'v1.4.4 起安装前会自动清理，请重新一键安装'
        }

        # user.disabled：Excel「禁用的项目」里列有 AttachLink（列在其中时启用置位静默失败）
        if ($foundVers.Count -eq 0) {
            Write-Hc 'user.disabled' 'skip' '未检测到 Office，无法检查「禁用的项目」'
        } else {
            $hitVer = ''
            foreach ($v in $foundVers) {
                $diKey = "$OfficeRootKey\$v\Excel\Resiliency\DisabledItems"
                if (-not (Test-Path -LiteralPath $diKey)) { continue }
                try {
                    $diItem = Get-Item -LiteralPath $diKey
                    foreach ($n in @($diItem.GetValueNames())) {
                        if (Test-DisabledItemsBlob ($diItem.GetValue($n))) { $hitVer = $v; break }
                    }
                } catch {}
                if ($hitVer) { break }
            }
            if ($hitVer) {
                Write-Hc 'user.disabled' 'fail' "Excel 已把 AttachLink 列入「禁用的项目」（Office $hitVer）" '在 Excel「文件 → 选项 → 加载项 → 管理：禁用的项目 → 转到」里选中 AttachLink 点「启用」'
            } else {
                Write-Hc 'user.disabled' 'ok' 'Excel「禁用的项目」里没有 AttachLink'
            }
        }

        # ==================== machine.* ====================

        $task = $null; $taskQueryErr = ''
        if ($TaskQueryInjectError) {
            $taskQueryErr = $TaskQueryInjectError   # 测试注入：模拟查询抛异常（权限不足 / 服务不可用等）
        } else {
            try { $task = @(Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop)[0] } catch { $taskQueryErr = $_.Exception.Message }
        }
        $taskExists = ($null -ne $task)
        # 区分「任务不存在」与「查询失败」：后者不得当成未部署（漏报机器级整体消失 / 误报半部署）
        $taskNotFound = (-not $taskExists) -and (-not $taskQueryErr -or $taskQueryErr -match 'cannot find|not find|does not exist|找不到|不存在')
        $taskUncertain = (-not $taskExists) -and (-not $taskNotFound)
        $dirExists = Test-Path -LiteralPath $MachineDir
        if ($taskUncertain) {
            $deploy = 'uncertain'   # 无法确认：不得据此判「未部署」或「半部署」
        } else {
            $deploy = Get-MachineDeployState $taskExists $dirExists
        }
        $skipMachine = ($deploy -eq 'none') -or $taskUncertain
        $machineSkipMsg = '未部署机器级看门狗——不使用鼎信诺（管理员运行）时属正常'
        if ($taskUncertain) {
            $machineSkipMsg = "无法确认机器级部署状态（计划任务查询失败：$taskQueryErr）——相关各项按跳过处理"
        }

        # machine.deploy：部署门禁（任务 + 目录同时在位才完整；都不在 = 未部署，不算异常）
        if ($taskUncertain) {
            Write-Hc 'machine.deploy' 'warn' "无法确认机器级部署状态（计划任务查询失败：$taskQueryErr）" '稍后重试；持续失败请把健康检查结果发给维护者'
        } elseif ($deploy -eq 'both') {
            Write-Hc 'machine.deploy' 'ok' '机器级看门狗已部署（计划任务 + Program Files 目录均在位）'
        } elseif ($deploy -eq 'none') {
            Write-Hc 'machine.deploy' 'skip' $machineSkipMsg
        } else {
            $lack = @()
            if (-not $taskExists) { $lack += '缺计划任务' }
            if (-not $dirExists) { $lack += "缺目录 $MachineDir" }
            Write-Hc 'machine.deploy' 'fail' ("机器级部署不完整：" + ($lack -join '、')) '重新「一键安装」补全机器级部署'
        }

        # machine.task：任务存在性与配置（任务直接运行无窗口启动器：看门狗运行期间 State=Running 属正常；
        # State=Ready = 任务当前没有在运行、看门狗不在运行 → warn，说明按上次运行结果分两种口径）
        if ($taskUncertain) {
            Write-Hc 'machine.task' 'warn' "无法确认计划任务（查询失败：$taskQueryErr）" '稍后重试；持续失败请把健康检查结果发给维护者'
        } elseif ($skipMachine) {
            Write-Hc 'machine.task' 'skip' $machineSkipMsg
        } elseif (-not $taskExists) {
            Write-Hc 'machine.task' 'fail' "计划任务「$TaskName」不存在" '重新「一键安装」部署机器级看门狗'
        } else {
            $cfg = Test-TaskConfig $task $MachineDir
            $stateNote = ''
            if ("$($task.State)" -eq 'Running') { $stateNote = '；任务在运行（启动器是任务的直接动作）' }
            elseif ("$($task.State)" -eq 'Ready') {
                $ltRes = 0
                try { $ltRes = Get-TaskResultCode (Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction Stop).LastTaskResult } catch { $ltRes = 0 }
                $stateNote = Get-ReadyStateNote $ltRes
            }
            $taskOwnerSid = ''
            try { $taskOwnerSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value } catch {}
            if ($cfg.Ok -and (Test-TaskOwnedByOtherUser $task $taskOwnerSid)) {
                # 任务登记账户 ≠ 当前用户：任务只在该账户登录时启动，当前用户登录时看门狗不会运行
                Write-Hc 'machine.task' 'warn' ("计划任务登记的账户（$($task.Principal.UserId)）不是当前用户——任务只在该账户登录时启动，当前用户登录时看门狗不会运行" + $stateNote) '用当前登录账户的管理员身份重新「一键安装」部署'
            } elseif ($cfg.Ok -and ("$($task.State)" -eq 'Ready')) {
                # 配置正常但任务没有在运行 → warn（看门狗必然不在运行）
                Write-Hc 'machine.task' 'warn' ("计划任务在位（登录触发 + 每 5 分钟重复兜底、最高权限、启动器指向 $MachineDir）" + $stateNote) '重新「一键安装」；反复出现就把看门狗日志发给维护者'
            } elseif ($cfg.Ok) {
                Write-Hc 'machine.task' 'ok' ("计划任务在位（登录触发 + 每 5 分钟重复兜底、最高权限、启动器指向 $MachineDir）" + $stateNote)
            } else {
                Write-Hc 'machine.task' 'warn' ("计划任务存在但配置与预期不符：" + $cfg.Detail + $stateNote) '重新「一键安装」修正任务配置'
            }
        }

        # machine.taskresult：上次运行返回码
        if ($skipMachine) {
            Write-Hc 'machine.taskresult' 'skip' $machineSkipMsg
        } elseif (-not $taskExists) {
            Write-Hc 'machine.taskresult' 'skip' '计划任务不存在，无运行结果可查'
        } else {
            try {
                $tinfo = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction Stop
                $rc = Get-TaskResultCode $tinfo.LastTaskResult
                if ((Get-TaskResultStatus $rc) -eq 'ok') {
                    $rcNote = ''
                    if ($rc -eq 2147946720) { $rcNote = '，即 0x800710E0：重复触发时任务已在运行，新实例被忽略——任务计划程序里显示的「操作员或管理员拒绝了请求」就是这个，不是错误' }
                    elseif ($rc -eq 267009) { $rcNote = '：正在运行' }
                    Write-Hc 'machine.taskresult' 'ok' "计划任务上次运行结果正常（返回码 $rc$rcNote）"
                } else {
                    Write-Hc 'machine.taskresult' 'warn' "计划任务上次运行返回码 $rc（0 为正常；上次运行时间 $($tinfo.LastRunTime)）" '看门狗未随任务启动时，重新「一键安装」或把日志发给维护者'
                }
            } catch {
                Write-Hc 'machine.taskresult' 'info' "任务运行结果读取失败：$($_.Exception.Message)"
            }
        }

        # machine.heartbeat：活性唯一依据 = 日志心跳（≤ 90 秒）；陈旧 / 缺失时按参数最多等 HeartbeatWaitSeconds 秒
        if ($skipMachine) {
            Write-Hc 'machine.heartbeat' 'skip' $machineSkipMsg
        } else {
            $hbHint = "确认没有安全软件禁用任务 → 管理员 PowerShell 执行 Start-ScheduledTask -TaskName '$TaskName' → 仍不行把日志发给维护者"
            $hb = Wait-Heartbeat $LogPath 90 $HeartbeatWaitSeconds
            if ($hb.Fresh) {
                Write-Hc 'machine.heartbeat' 'ok' "看门狗心跳正常（$($hb.AgeSeconds) 秒前）"
            } elseif ($null -eq $hb.AgeSeconds) {
                Write-Hc 'machine.heartbeat' 'fail' '看门狗尚未写过日志' $hbHint
            } else {
                Write-Hc 'machine.heartbeat' 'fail' "看门狗可能未运行：日志 $($hb.AgeSeconds) 秒未刷新" $hbHint
            }
        }

        # machine.files：部署目录三件齐全（旧版遗留的 watchdog_launcher.vbs 只提示、不算失败）
        if ($skipMachine) {
            Write-Hc 'machine.files' 'skip' $machineSkipMsg
        } else {
            $needFiles = @('AttachLink.xlam', 'watchdog.ps1', 'AttachLink.WatchdogLauncher.exe')
            $missingFiles = @($needFiles | Where-Object { -not (Test-Path -LiteralPath (Join-Path $MachineDir $_)) })
            if ($missingFiles.Count -eq 0) {
                $legacyNote = ''
                try {
                    if (Test-Path -LiteralPath (Join-Path $MachineDir 'watchdog_launcher.vbs')) {
                        $legacyNote = '；发现旧启动器 watchdog_launcher.vbs，重新「一键安装」会清掉'
                    }
                } catch {}
                # 看门狗引擎形态（信息性，不影响判定）：部署目录里有 AttachLink.Watchdog.exe = .NET 版（启动器优先用它）；
                # 没有它时启动器回退 PowerShell 5.1 + watchdog.ps1——它不是必需文件，缺失不算失败
                $engineNote = '；看门狗引擎=PowerShell（watchdog.ps1）'
                try {
                    if (Test-Path -LiteralPath (Join-Path $MachineDir 'AttachLink.Watchdog.exe')) {
                        $engineNote = '；看门狗引擎=.NET（AttachLink.Watchdog.exe）'
                    }
                } catch {}
                Write-Hc 'machine.files' 'ok' ("机器级文件齐全（$($needFiles -join '、')）" + $legacyNote + $engineNote)
            } else {
                Write-Hc 'machine.files' 'fail' ("机器级文件缺失：" + ($missingFiles -join '、')) '重新「一键安装」补全机器级部署'
            }
        }

        # machine.version：机器级副本 vs 安装包（xlam 与看门狗脚本都比对；旧版 = 鼎信诺场景运行旧代码）
        if ($skipMachine) {
            Write-Hc 'machine.version' 'skip' $machineSkipMsg
        } else {
            $diffs = @(); $compared = 0
            foreach ($pair in @(@((Join-Path $MachineDir 'AttachLink.xlam'), $distXlam, 'AttachLink.xlam'),
                                @((Join-Path $MachineDir 'watchdog.ps1'), $srcWd, 'watchdog.ps1'))) {
                $pfFile = $pair[0]; $refFile = $pair[1]; $label = $pair[2]
                if (-not (Test-Path -LiteralPath $pfFile)) { continue }   # 缺失由 machine.files 报告
                if (-not (Test-Path -LiteralPath $refFile)) { continue }  # 安装包缺基准 → 下面统一按 skip 处理
                try {
                    $compared++
                    $hPf = Get-Hash16 $pfFile
                    $hRef = Get-Hash16 $refFile
                    if ($hPf -ne $hRef) { $diffs += "$label（前 16 位 $hPf vs $hRef）" }
                } catch { $diffs += "$label 比对失败：$($_.Exception.Message)" }
            }
            if ($compared -eq 0) {
                Write-Hc 'machine.version' 'skip' '无可比对基准（机器级文件或安装包文件缺失）'
            } elseif ($diffs.Count -eq 0) {
                Write-Hc 'machine.version' 'ok' '机器级副本与安装包版本一致'
            } else {
                Write-Hc 'machine.version' 'warn' ("机器级副本为旧版（鼎信诺场景会运行旧代码），请重新一键安装：" + ($diffs -join '；')) '重新「一键安装」同步机器级副本'
            }
        }

        # machine.trusted：HKLM 受信任位置指向部署目录（高完整性 Excel 只认 HKLM）
        if ($skipMachine) {
            Write-Hc 'machine.trusted' 'skip' $machineSkipMsg
        } elseif (-not $ver) {
            Write-Hc 'machine.trusted' 'skip' '未检测到 Office，无法检查 HKLM 受信任位置'
        } else {
            $tlmKey = "$mRoot\$ver\Excel\Security\Trusted Locations\AttachLink"
            $tlmPath = ''
            try { if (Test-Path -LiteralPath $tlmKey) { $tlmPath = [string](Get-ItemProperty -LiteralPath $tlmKey -Name Path -ErrorAction Stop).Path } } catch {}
            if ($tlmPath -and (Test-SameTrustedPath $tlmPath "$MachineDir\")) {
                Write-Hc 'machine.trusted' 'ok' "机器级受信任位置已登记（HKLM，Office $ver → $tlmPath）"
            } else {
                Write-Hc 'machine.trusted' 'fail' "机器级受信任位置未登记或路径不符（HKLM，Office $ver，期望 $MachineDir\）" '重新「一键安装」；高完整性 Excel 只认 HKLM 信任位置'
            }
        }

        # machine.conf：父进程白名单（信息性快照；没有 conf = 默认不过滤，属正常。异常形态才 warn）
        if ($skipMachine) {
            Write-Hc 'machine.conf' 'skip' $machineSkipMsg
        } else {
            $confPath = Join-Path $MachineDir 'watchdog.conf'
            if (-not (Test-Path -LiteralPath $confPath)) {
                Write-Hc 'machine.conf' 'info' '未配置父进程白名单（无 watchdog.conf，默认附着所有 Excel 实例）'
            } else {
                try {
                    $confLine = @(Get-Content -LiteralPath $confPath -ErrorAction Stop |
                        Where-Object { $_ -match '^\s*ParentAllowlist\s*=' } | Select-Object -First 1)
                    if ($confLine.Count -gt 0) {
                        $confVal = ($confLine[0] -replace '^\s*ParentAllowlist\s*=\s*', '').Trim()
                        if ($confVal) {
                            Write-Hc 'machine.conf' 'info' "父进程白名单生效中：$confVal（仅附着这些进程拉起的 Excel）"
                        } else {
                            Write-Hc 'machine.conf' 'info' 'watchdog.conf 里 ParentAllowlist 为空（不过滤）'
                        }
                    } else {
                        Write-Hc 'machine.conf' 'warn' 'watchdog.conf 存在但没有 ParentAllowlist 行' '确认配置是否有意为之；重新「一键安装」可重置'
                    }
                } catch {
                    Write-Hc 'machine.conf' 'warn' "watchdog.conf 读取失败：$($_.Exception.Message)"
                }
            }
        }

        # ==================== diag.* ====================

        # diag.log：看门狗日志可读性（info：路径、大小、最后写入、总行数）
        if (-not (Test-Path -LiteralPath $LogPath)) {
            if ($skipMachine) { Write-Hc 'diag.log' 'skip' $machineSkipMsg }
            else { Write-Hc 'diag.log' 'warn' "看门狗日志不存在：$LogPath" }
        } else {
            try {
                $logItem = Get-Item -LiteralPath $LogPath
                $sizeKb = [math]::Round($logItem.Length / 1KB, 1)
                $lineCount = @(Get-LogLines $LogPath).Count
                Write-Hc 'diag.log' 'info' ("看门狗日志 {0}（{1} KB，共 {2} 行，最后写入 {3}）" -f $LogPath, $sizeKb, $lineCount, $logItem.LastWriteTime.ToString('MM-dd HH:mm:ss'))
            } catch {
                Write-Hc 'diag.log' 'warn' "看门狗日志读取失败：$($_.Exception.Message)"
            }
        }

        # diag.logerrors：最近 200 行的错误关键字（失败 / 致命 / 异常 / 仍不存在 / 读不到）
        if (-not (Test-Path -LiteralPath $LogPath)) {
            Write-Hc 'diag.logerrors' 'skip' '日志不存在，无内容可查'
        } else {
            $tailLines = @(Get-LogLines $LogPath | Select-Object -Last 200)
            $errLines = @($tailLines | Where-Object { $_ -match '失败|致命|异常|仍不存在|读不到' })
            # 自愈类瞬态不算问题：附着失败退避重试（目标实例正忙或正在退出，下轮自愈）、无法确认在位副本（宿主正忙）、
            # Window.Visible 暂时读不到。「已持续 60 秒读不到」是升级告警，不算瞬态。
            $benign = @($errLines | Where-Object { $_ -match '退避重试|无法确认在位的 AttachLink\.xlam|Window\.Visible 读不到，' })
            $serious = @($errLines | Where-Object { $benign -notcontains $_ })
            if ($errLines.Count -eq 0) {
                Write-Hc 'diag.logerrors' 'info' "最近 $($tailLines.Count) 行日志无错误关键字"
            } elseif ($serious.Count -eq 0) {
                # 少量自愈类重试 = 正常瞬态；但持续发生说明看门狗长期附着不上（目标实例无响应或反复被打断）——升级 warn
                $now = Get-Date
                $retryTimes = @()
                foreach ($bl in $benign) {
                    # 时间戳取日志行首 MM-dd HH:mm:ss；跨天按日期差处理（1 月接 12 月的极端跨度不影响「分钟级」判定）
                    if ($bl -match '^(\d{2})-(\d{2})[ ](\d{2}):(\d{2}):(\d{2})') {
                        try { $retryTimes += [DateTime]::new($now.Year, [int]$Matches[1], [int]$Matches[2], [int]$Matches[3], [int]$Matches[4], [int]$Matches[5]) } catch {}
                    }
                }
                $spanMin = 0
                if ($retryTimes.Count -ge 2) {
                    $retryTimes = @($retryTimes | Sort-Object)
                    $firstT = $retryTimes[0]; $lastT = $retryTimes[$retryTimes.Count - 1]
                    if ($lastT -lt $firstT) { $lastT = $lastT.AddDays(1) }
                    $spanMin = [int][math]::Floor(($lastT - $firstT).TotalMinutes)
                }
                $sustained = ($benign.Count -ge 20) -or ($benign.Count -ge 10 -and $spanMin -ge 10)
                if ($sustained) {
                    Write-Hc 'diag.logerrors' 'warn' ("看门狗持续附着失败（最近 {0} 行有 {1} 条退避重试，跨越约 {2} 分钟）——目标实例可能长期无响应，或看门狗反复被打断" -f $tailLines.Count, $benign.Count, $spanMin) '若持续出现请排查目标 Excel 实例是否卡死、看门狗是否被安全软件反复打断；仍无法定位时把看门狗日志发给维护者'
                } else {
                    Write-Hc 'diag.logerrors' 'info' "最近 $($tailLines.Count) 行日志有 $($benign.Count) 条自愈类重试（附着失败退避重试等，属正常瞬态），无严重错误"
                }
            } else {
                $lastTwo = @($serious | Select-Object -Last 2 | ForEach-Object { if ($_.Length -gt 200) { $_.Substring(0, 200) } else { $_ } })
                Write-Hc 'diag.logerrors' 'warn' "最近 $($tailLines.Count) 行日志有 $($serious.Count) 条错误记录（另有 $($benign.Count) 条自愈类重试属正常）" ("最近两条：" + ($lastTwo -join ' ｜ '))
            }
        }

        # diag.setuplog：安装器日志里最近一次结论（[结果] 行）
        $setupLog = ''
        foreach ($cand in @((Join-Path (Split-Path -Parent $PayloadDir) 'AttachLinkSetup.log'), (Join-Path $env:TEMP 'AttachLinkSetup.log'))) {
            if ($cand -and (Test-Path -LiteralPath $cand)) { $setupLog = $cand; break }
        }
        if (-not $setupLog) {
            Write-Hc 'diag.setuplog' 'info' '未找到安装器日志（AttachLinkSetup.log）——尚未跑过安装动作时属正常'
        } else {
            $resultLines = @()
            try { $resultLines = @(Get-LogLines $setupLog | Where-Object { $_ -match '\[结果\]' } | Select-Object -Last 1) } catch {}
            if ($resultLines.Count -gt 0) {
                $rl = [string]$resultLines[0]
                if ($rl.Length -gt 200) { $rl = $rl.Substring(0, 200) }
                Write-Hc 'diag.setuplog' 'info' "安装器日志最近结论：$rl"
            } else {
                Write-Hc 'diag.setuplog' 'info' '安装器日志里没有 [结果] 行（可能尚未跑过安装动作）'
            }
        }

        # diag.openlist：OPEN 序列快照（排障时对照第三方加载项；逐条一行，Python 逐行推送）
        if (-not $ver) {
            Write-Hc 'diag.openlist' 'skip' '未检测到 Office，无 OPEN 序列'
        } else {
            $optKeyD = "$OfficeRootKey\$ver\Excel\Options"
            if (-not (Test-Path -LiteralPath $optKeyD)) {
                Write-Hc 'diag.openlist' 'skip' "Office $ver 无 Options 键"
            } else {
                $openItems = @((Get-ItemProperty -LiteralPath $optKeyD).PSObject.Properties |
                    Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] } |
                    Sort-Object { [int]($_.Name -replace '\D', '') })
                if ($openItems.Count -eq 0) {
                    Write-Hc 'diag.openlist' 'info' "OPEN 序列（Office $ver）为空"
                } else {
                    foreach ($oi in $openItems) {
                        $val = [string]$oi.Value
                        if ($val.Length -gt 300) { $val = $val.Substring(0, 300) + '…' }
                        Write-Hc 'diag.openlist' 'info' ("{0}={1}" -f $oi.Name, $val)
                    }
                }
            }
        }

        Write-HcSummary
        exit (Get-HcExitCode)
    } catch {
        try { Write-Hc 'script' 'fail' ("健康检查脚本自身出错：" + $_.Exception.Message) } catch {}
        try { Write-HcSummary } catch {}
        exit 1
    }
}
