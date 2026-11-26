# =====================================================================
# AttachLink 健康检查（scripts\healthcheck.ps1）—— 单元 + 端到端回归（v1.5.0）
# ---------------------------------------------------------------------
# A. 点源载入：主逻辑有守卫（不产生输出、不触碰环境），函数与协议计数可用；两个文件都是 UTF-8 with BOM
# B. OPEN 条目识别与 build\install.ps1 对拍：同一批样本过两边同名函数，结果必须一致（防语义漂移）；
#    OPEN 序列键名 ^OPEN(\d+)?$ 不误伤 OPENROAM 等诱饵键
# C. 受信任路径比较 / 禁用形态 OPEN / 机器级部署门禁 / 任务配置（含 State=Disabled 与 Ready） / 任务返回码（纯函数，伪造对象）
# D. 心跳：年龄计算 + 等待刷新（临时文件 + 后台 Job 触碰）
# E. DisabledItems 二进制值（含中文名边界）
# F. 协议输出：制表符 / 换行清洗、计数、汇总行、退出码
# G. 端到端（子进程 + 注入临时注册表根 / 临时目录 / 临时日志）：协议形态纯净、机器级门禁、退出码自洽
# 用法：powershell -ExecutionPolicy Bypass -File test-healthcheck.ps1      期望末行 RESULT: PASS
# 编码：本文件含中文，必须保存为 UTF-8 with BOM（WORKLOG 根因 #11）。
# =====================================================================
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$hcPath = Join-Path $PSScriptRoot 'healthcheck.ps1'
$installPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'build\install.ps1'

$ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; "FAIL: $label" } else { "pass: $label" } }
function Get-FunctionsText([string]$File, [string[]]$Names) {
    # 从脚本文件里按名字取出函数定义文本（不执行脚本主体）。调用方须在脚本作用域用「. 」载入
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($File, [ref]$null, [ref]$null)
    $found = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false) |
        Where-Object { $Names -contains $_.Name }
    return (@($found | ForEach-Object { $_.Extent.Text }) -join "`n")
}

# ---------------------------------------------------------------------
'=== A. 点源载入与编码 ==='
$hcBytes = [IO.File]::ReadAllBytes($hcPath)
Check ($hcBytes.Length -gt 3 -and $hcBytes[0] -eq 0xEF -and $hcBytes[1] -eq 0xBB -and $hcBytes[2] -eq 0xBF) 'A healthcheck.ps1 保存为 UTF-8 with BOM'
$selfBytes = [IO.File]::ReadAllBytes($PSCommandPath)
Check ($selfBytes.Length -gt 3 -and $selfBytes[0] -eq 0xEF -and $selfBytes[1] -eq 0xBB -and $selfBytes[2] -eq 0xBF) 'A 本测试文件保存为 UTF-8 with BOM'

$dotOut = @(. $hcPath)
Check ($dotOut.Count -eq 0) 'A 点源载入零输出（主逻辑有守卫，不触碰真实环境）'
foreach ($fn in @('Write-Hc', 'Write-HcSummary', 'Get-HcExitCode', 'ConvertFrom-AddInEntry', 'Test-IsAttachLinkEntry',
                  'Test-IsLibraryEntry', 'Test-IsDisabledOpenValue', 'Test-DisabledItemsBlob', 'Test-SameTrustedPath',
                  'Get-MachineDeployState', 'Get-OfficeVersions', 'Get-LogAgeSeconds', 'Wait-Heartbeat', 'Get-LogLines',
                  'Get-Hash16', 'Test-TaskConfig', 'Test-TaskOwnedByOtherUser', 'Get-TaskResultStatus')) {
    Check ([bool](Get-Command $fn -CommandType Function -ErrorAction SilentlyContinue)) "A 函数已载入：$fn"
}

# ---------------------------------------------------------------------
'=== B. OPEN 条目识别：与 install.ps1 对拍 ==='
$openNames = @('ConvertFrom-AddInEntry', 'Test-IsAttachLinkEntry', 'Test-IsLibraryEntry')
$ren = Get-FunctionsText $installPath $openNames
Check ($ren.Contains('function ConvertFrom-AddInEntry') -and $ren.Contains('function Test-IsAttachLinkEntry') -and $ren.Contains('function Test-IsLibraryEntry')) 'B 取到 install.ps1 的三个判定函数'
foreach ($n in $openNames) { $ren = $ren -replace ("\b" + [regex]::Escape($n) + "\b"), ("Install_" + $n) }
. ([scriptblock]::Create($ren))
$destXlam = 'C:\Users\u\AppData\Roaming\Microsoft\AddIns\AttachLink.xlam'   # 两边函数都读这个脚本作用域变量
$samples = @(
    '"AttachLink.xlam"',
    '  "C:\Program Files\ExcelEnhancer\ExcelEnhancer.xlam" ',
    '/A "C:\Program Files\AttachLink\AttachLink.xlam"',
    '/A "AttachLink.xlam"',
    '"C:\Tools\AttachLinkCopy\AttachLink.xlam"',
    '/R "C:\Users\u\AppData\Local\AuditCat\AttachmentManager\AuditCatAttachment.xll"',
    'C:\A\AttachLink2.xlam',
    'D:/x/y/attachlink.XLAM',
    'AttachLink.xlam.bak',
    'C:\Users\u\AppData\Roaming\Microsoft\AddIns\AttachLink.xlam'
)
foreach ($s in $samples) {
    $a1 = ConvertFrom-AddInEntry $s; $b1 = Install_ConvertFrom-AddInEntry $s
    Check ($a1 -eq $b1) "B ConvertFrom 对拍一致：[$s] → [$a1]"
    Check ((Test-IsAttachLinkEntry $a1) -eq (Install_Test-IsAttachLinkEntry $b1)) "B IsAttachLink 对拍一致：[$a1]"
    Check ((Test-IsLibraryEntry $a1) -eq (Install_Test-IsLibraryEntry $b1)) "B IsLibrary 对拍一致：[$a1]"
}
# 期望值抽查（不只两边一致，还要符合语义）
Check ((ConvertFrom-AddInEntry '/A "AttachLink.xlam"') -eq 'AttachLink.xlam') 'B /A 前缀归一化后期望为库文件名'
Check ((Test-IsAttachLinkEntry 'AttachLink.xlam') -and (Test-IsLibraryEntry 'AttachLink.xlam')) 'B 库文件名：是 AttachLink 且是库目录登记'
Check (-not (Test-IsAttachLinkEntry 'C:\A\AttachLink2.xlam')) 'B 相近文件名不误判'
Check ((Test-IsDisabledOpenValue '/A "AttachLink.xlam"') -and (Test-IsDisabledOpenValue ' /a "x"')) 'B 禁用形态 OPEN：/A 前缀识别'
Check (-not (Test-IsDisabledOpenValue '"AttachLink.xlam"') -and -not (Test-IsDisabledOpenValue '/R "C:\a\b.xll"')) 'B 禁用形态 OPEN：正常值与 /R 不误判'
# OPEN 序列键名识别：^OPEN(\d+)?$ 不得误伤诱饵键（OPENROAM 是 Excel 自己会写的相邻值名）
Check (('OPEN' -match '^OPEN(\d+)?$') -and ('OPEN12' -match '^OPEN(\d+)?$')) 'B OPEN 键名：OPEN / OPENn 命中 ^OPEN(\d+)?$'
Check (-not ('OPENROAM' -match '^OPEN(\d+)?$') -and -not ('OPEN1A' -match '^OPEN(\d+)?$') -and -not ('MyOPEN' -match '^OPEN(\d+)?$')) 'B OPEN 键名：OPENROAM 等诱饵键不命中'

# ---------------------------------------------------------------------
'=== C. 纯函数：信任路径 / 部署门禁 / 任务配置 / 返回码 ==='
Check (Test-SameTrustedPath 'C:\Users\u\AppData\Roaming\Microsoft\AddIns\' 'c:\users\u\appdata\roaming\microsoft\addins') 'C 信任路径：大小写与尾反斜杠不敏感'
Check (Test-SameTrustedPath '"C:\Program Files\AttachLink\"' 'C:\Program Files\AttachLink') 'C 信任路径：引号不敏感'
Check (Test-SameTrustedPath '%APPDATA%\Microsoft\AddIns\' ($env:APPDATA + '\Microsoft\AddIns\')) 'C 信任路径：环境变量展开后比较'
Check (-not (Test-SameTrustedPath 'C:\A\' 'C:\B\')) 'C 信任路径：不同路径判不同'
Check ((Get-MachineDeployState $true $true) -eq 'both') 'C 部署门禁：任务+目录 → both'
Check ((Get-MachineDeployState $false $false) -eq 'none') 'C 部署门禁：都无 → none'
Check ((Get-MachineDeployState $true $false) -eq 'partial' -and (Get-MachineDeployState $false $true) -eq 'partial') 'C 部署门禁：只其一 → partial'

$goodTask = [pscustomobject]@{
    Principal = [pscustomobject]@{ RunLevel = 'Highest' }
    Actions   = @([pscustomobject]@{ Execute = 'wscript.exe'; Arguments = '"C:\Program Files\AttachLink\watchdog_launcher.vbs"' })
    Triggers  = @([pscustomobject]@{ Enabled = $true })
}
$goodTask.Triggers[0].PSObject.TypeNames.Insert(0, 'Microsoft.Management.Infrastructure.CimInstance#MSFT_TaskLogonTrigger')
Check ((Test-TaskConfig $goodTask 'C:\Program Files\AttachLink').Ok) 'C 任务配置：标准部署判 Ok'
$badLevel = [pscustomobject]@{
    Principal = [pscustomobject]@{ RunLevel = 'Limited' }
    Actions   = $goodTask.Actions
    Triggers  = $goodTask.Triggers
}
Check (-not (Test-TaskConfig $badLevel 'C:\Program Files\AttachLink').Ok) 'C 任务配置：RunLevel 非 Highest → 不 Ok'
$badAction = [pscustomobject]@{
    Principal = [pscustomobject]@{ RunLevel = 'Highest' }
    Actions   = @([pscustomobject]@{ Execute = 'wscript.exe'; Arguments = '"D:\Elsewhere\watchdog_launcher.vbs"' })
    Triggers  = $goodTask.Triggers
}
Check (-not (Test-TaskConfig $badAction 'C:\Program Files\AttachLink').Ok) 'C 任务配置：启动器不在部署目录 → 不 Ok'
$noTrigger = [pscustomobject]@{
    Principal = [pscustomobject]@{ RunLevel = 'Highest' }
    Actions   = $goodTask.Actions
    Triggers  = @()
}
Check (-not (Test-TaskConfig $noTrigger 'C:\Program Files\AttachLink').Ok) 'C 任务配置：无触发器 → 不 Ok'
$disabledTask = [pscustomobject]@{
    Principal = [pscustomobject]@{ RunLevel = 'Highest' }
    Actions   = $goodTask.Actions
    Triggers  = $goodTask.Triggers
    State     = 'Disabled'
}
Check (-not (Test-TaskConfig $disabledTask 'C:\Program Files\AttachLink').Ok) 'C 任务配置：State=Disabled → 不 Ok（任务不会随登录拉起）'
$readyTask = [pscustomobject]@{
    Principal = [pscustomobject]@{ RunLevel = 'Highest' }
    Actions   = $goodTask.Actions
    Triggers  = $goodTask.Triggers
    State     = 'Ready'
}
Check ((Test-TaskConfig $readyTask 'C:\Program Files\AttachLink').Ok) 'C 任务配置：State=Ready → Ok（启动器不等待看门狗进程，属正常）'
# NB-1a：登录触发器存在但被禁用（Enabled=False，任务自身 State 可能仍是 Ready）→ 配置不正常
$disTriggerTask = [pscustomobject]@{
    Principal = [pscustomobject]@{ RunLevel = 'Highest' }
    Actions   = $goodTask.Actions
    Triggers  = @([pscustomobject]@{ Enabled = $false })
    State     = 'Ready'
}
$disTriggerTask.Triggers[0].PSObject.TypeNames.Insert(0, 'Microsoft.Management.Infrastructure.CimInstance#MSFT_TaskLogonTrigger')
$rDisTr = Test-TaskConfig $disTriggerTask 'C:\Program Files\AttachLink'
Check (-not $rDisTr.Ok -and $rDisTr.Detail.Contains('触发器被禁用')) 'C 任务配置：登录触发器 Enabled=False → 不 Ok，提示「触发器被禁用」'

Check ((Get-TaskResultStatus 0) -eq 'ok') 'C 任务返回码：0 → ok'
Check ((Get-TaskResultStatus 267009) -eq 'ok') 'C 任务返回码：0x41301（运行中）→ ok'
Check ((Get-TaskResultStatus 1) -eq 'warn' -and (Get-TaskResultStatus -2147024894) -eq 'warn') 'C 任务返回码：其他 → warn'

# 任务登记账户 ≠ 当前用户（council C4：用别的管理员账号过 UAC 部署时任务绑错账户）：
# UserId 可能是 SID 或 DOMAIN\user，统一换算成 SID 再比；任一侧读不到 → 不误报
$mySid = ''
try { $mySid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value } catch {}
Check (Test-TaskOwnedByOtherUser ([pscustomobject]@{ Principal = [pscustomobject]@{ UserId = 'S-1-5-21-9-999' } }) 'S-1-5-21-9-998') 'C 任务登记账户：SID 不同 → 判定「其他用户」（健康检查应 warn）'
Check (-not (Test-TaskOwnedByOtherUser ([pscustomobject]@{ Principal = [pscustomobject]@{ UserId = 'S-1-5-21-9-999' } }) 'S-1-5-21-9-999')) 'C 任务登记账户：SID 相同 → 非其他用户'
if ($mySid) {
    $meTask = [pscustomobject]@{ Principal = [pscustomobject]@{ UserId = ($env:USERDOMAIN + '\' + $env:USERNAME) } }
    Check (-not (Test-TaskOwnedByOtherUser $meTask $mySid)) 'C 任务登记账户：DOMAIN\user 换算 SID 后与当前用户一致 → 非其他用户'
}
Check (-not (Test-TaskOwnedByOtherUser ([pscustomobject]@{ Principal = [pscustomobject]@{ UserId = '' } }) $mySid)) 'C 任务登记账户：任务 UserId 为空 → 不误报'
Check (-not (Test-TaskOwnedByOtherUser ([pscustomobject]@{ Principal = [pscustomobject]@{ UserId = 'S-1-5-21-9-999' } }) '')) 'C 任务登记账户：当前用户 SID 读不到 → 不误报'

# ---------------------------------------------------------------------
'=== D. 心跳：年龄与等待刷新 ==='
$hbDir = Join-Path $env:TEMP ('opencode\hctest-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $hbDir | Out-Null
$hbLog = Join-Path $hbDir 'wd.log'
'x' | Set-Content $hbLog
[IO.File]::SetLastWriteTime($hbLog, (Get-Date).AddSeconds(-300))
$r = Wait-Heartbeat $hbLog 90 0
Check (-not $r.Fresh -and $r.AgeSeconds -ge 290) ('D 心跳：300 秒未刷新判不新鲜；Age=' + $r.AgeSeconds)
'x' | Set-Content $hbLog
$r = Wait-Heartbeat $hbLog 90 0
Check ($r.Fresh -and $r.AgeSeconds -le 5) 'D 心跳：刚写的日志判新鲜'
Remove-Item $hbLog -Force
$r = Wait-Heartbeat $hbLog 90 0
Check (-not $r.Fresh -and $null -eq $r.AgeSeconds) 'D 心跳：日志不存在 → AgeSeconds 为空'
'x' | Set-Content $hbLog
[IO.File]::SetLastWriteTime($hbLog, (Get-Date).AddMinutes(-5))
$r = Wait-Heartbeat $hbLog 90 2
Check (-not $r.Fresh -and $r.Waited -eq 2) 'D 心跳：等待 2 秒仍陈旧 → 超时返回（Waited=2）'
'x' | Set-Content $hbLog
[IO.File]::SetLastWriteTime($hbLog, (Get-Date).AddMinutes(-5))
$job = Start-Job -ScriptBlock { param($p) Start-Sleep -Seconds 2; [IO.File]::SetLastWriteTime($p, (Get-Date)) } -ArgumentList $hbLog
try {
    $r = Wait-Heartbeat $hbLog 90 12
    Check ($r.Fresh -and $r.Waited -ge 1 -and $r.Waited -lt 12) ('D 心跳：等待期间被后台刷新 → 判新鲜；Waited=' + $r.Waited)
} finally {
    Remove-Job $job -Force -ErrorAction SilentlyContinue
}
Remove-Item $hbDir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------
'=== E. DisabledItems 二进制值（含中文名边界）==='
function New-DisabledBlob([string[]]$Parts) {
    # 与 test-watchdog-units.ps1 同款构造：12 字节头 + 若干 UTF-16 段（各带 NUL 结尾）
    $ms = New-Object System.IO.MemoryStream
    $ms.Write([byte[]](1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), 0, 12)
    foreach ($s in $Parts) { $b = [Text.Encoding]::Unicode.GetBytes($s + [char]0); $ms.Write($b, 0, $b.Length) }
    return $ms.ToArray()
}
Check (Test-DisabledItemsBlob (New-DisabledBlob @('c:\users\u\appdata\roaming\microsoft\addins\attachlink.xlam'))) 'E 禁用项目：库目录副本识别'
Check (Test-DisabledItemsBlob (New-DisabledBlob @('c:\program files\attachlink\attachlink.xlam', 'attachlink.connect'))) 'E 禁用项目：路径后带 ProgID 也识别'
Check (-not (Test-DisabledItemsBlob (New-DisabledBlob @('c:\users\u\appdata\roaming\microsoft\addins\attachlinkshim.dll', 'attachlinkshim.connect')))) 'E 禁用项目：Shim 引导器旧记录不算'
Check (-not (Test-DisabledItemsBlob (New-DisabledBlob @('c:\x\myattachlink.xlam')))) 'E 禁用项目：前缀相近文件名不算'
Check (-not (Test-DisabledItemsBlob (New-DisabledBlob @('c:\x\attachlink.xlam.bak')))) 'E 禁用项目：后缀相近文件名不算'
Check (Test-DisabledItemsBlob (New-DisabledBlob @('c:\底稿\附件\AttachLink.xlam'))) 'E 禁用项目：中文目录名边界识别'
Check (-not (Test-DisabledItemsBlob (New-DisabledBlob @('c:\x\AttachLink.xlam中文')))) 'E 禁用项目：文件名后紧跟中文不算（非 NUL 边界）'
Check (-not (Test-DisabledItemsBlob (New-DisabledBlob @('c:\x\中AttachLink.xlam')))) 'E 禁用项目：文件名前是中文（非分隔符）不算'
Check (-not (Test-DisabledItemsBlob ([byte[]]@())) -and -not (Test-DisabledItemsBlob $null)) 'E 禁用项目：空值 / 空引用不算'

# ---------------------------------------------------------------------
'=== F. 协议输出：清洗 / 计数 / 汇总 / 退出码 ==='
$script:hcCounts = @{ ok = 0; warn = 0; fail = 0; skip = 0; info = 0 }
$fLines = @(Write-Hc 'user.file' 'ok' ("第一行`n第二行`t带制表符") ("提示`t也清洗`n换行"))
Check ($fLines.Count -eq 2) 'F Write-Hc：消息 + 提示共两行'
$f0 = $fLines[0] -split "`t"
Check ($f0.Count -eq 4 -and $f0[0] -eq 'HC' -and $f0[1] -eq 'user.file' -and $f0[2] -eq 'ok') 'F 协议：HC 行恰四段（消息内制表符已清洗）'
Check ($f0[3] -notmatch "`r" -and $f0[3] -notmatch "`n" -and $f0[3].Contains('第一行 第二行 带制表符')) 'F 协议：消息内换行已清洗为空格'
$f1 = $fLines[1] -split "`t"
Check ($f1.Count -eq 3 -and $f1[0] -eq 'HCHINT' -and $f1[1] -eq 'user.file' -and $f1[2].Contains('提示 也清洗 换行')) 'F 协议：HCHINT 行三段且同样清洗'
Check ($script:hcCounts['ok'] -eq 1) 'F 计数：ok +1'
Write-Hc 'a' 'warn' 'w' | Out-Null
Write-Hc 'b' 'fail' 'f' | Out-Null
Write-Hc 'c' 'skip' 's' | Out-Null
Write-Hc 'd' 'info' 'i' | Out-Null
$sumLines = @(Write-HcSummary)
Check ($sumLines.Count -eq 1 -and $sumLines[0] -eq "HC`tSUMMARY`tok=1 warn=1 fail=1 skip=1 info=1") ('F 汇总行格式与计数；实际 ' + $sumLines[0])
Check ((Get-HcExitCode) -eq 10) 'F 退出码：有 fail → 10'
$script:hcCounts = @{ ok = 5; warn = 1; fail = 0; skip = 2; info = 1 }
Check ((Get-HcExitCode) -eq 0) 'F 退出码：无 fail → 0（warn 不影响）'

# ---------------------------------------------------------------------
'=== G. 端到端（子进程 + 注入临时环境）==='
function Parse-HcOutput([string[]]$Lines) {
    # 严格解析：任何不符合协议的行都进 Bad（协议纯净度断言用）
    $items = @(); $hints = @{}; $summary = $null; $bad = @()
    foreach ($ln in $Lines) {
        $s = ([string]$ln) -replace '^\ufeff', ''
        if (-not $s.Trim()) { continue }
        $parts = $s -split "`t"
        if ($parts[0] -eq 'HC' -and $parts.Count -ge 3 -and $parts[1] -eq 'SUMMARY') { $summary = $parts[2]; continue }
        if ($parts[0] -eq 'HC' -and $parts.Count -eq 4 -and @('ok', 'warn', 'fail', 'skip', 'info') -contains $parts[2]) {
            $items += [pscustomobject]@{ Id = $parts[1]; Status = $parts[2]; Message = $parts[3] }
            continue
        }
        if ($parts[0] -eq 'HCHINT' -and $parts.Count -eq 3) { $hints[$parts[1]] = $parts[2]; continue }
        $bad += $s
    }
    return @{ Items = $items; Hints = $hints; Summary = $summary; Bad = $bad }
}
function Get-St($Parsed, [string]$Id) {
    $hit = @($Parsed.Items | Where-Object { $_.Id -eq $Id } | Select-Object -First 1)
    if ($hit.Count -eq 0) { return '(missing)' }
    return $hit[0].Status
}

$g = Join-Path $env:TEMP ('opencode\hce2e-' + [guid]::NewGuid().ToString('N'))
$regRoot = 'HKCU:\Software\AttachLinkHealthTest\' + [guid]::NewGuid().ToString('N')
$officeRoot = "$regRoot\Office"
$machineDir = Join-Path $g 'Machine'
$payloadDir = Join-Path $g 'Payload'
$gLog = Join-Path $g 'wd.log'
try {
    # 注入用户级环境：16.0 有 Excel 键 + OPEN 库目录登记（启用形态）+ 第三方条目 + OPENROAM 诱饵键
    New-Item -Path "$officeRoot\16.0\Excel\Options" -Force | Out-Null
    New-ItemProperty -LiteralPath "$officeRoot\16.0\Excel\Options" -Name 'OPEN' -Value '"C:\Program Files\ExcelEnhancer\ExcelEnhancer.xlam"' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath "$officeRoot\16.0\Excel\Options" -Name 'OPEN1' -Value '"AttachLink.xlam"' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath "$officeRoot\16.0\Excel\Options" -Name 'OPENROAM' -Value '"C:\Roam\Decoy.xlam"' -PropertyType String | Out-Null
    # 安装包与机器级目录（内容一致 → machine.version ok）
    New-Item -ItemType Directory -Force -Path "$payloadDir\dist", "$payloadDir\scripts", $machineDir | Out-Null
    [IO.File]::WriteAllText("$payloadDir\dist\AttachLink.xlam", 'dummy-xlam-v1', [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText("$payloadDir\scripts\watchdog.ps1", 'dummy-wd-v1', [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText("$machineDir\AttachLink.xlam", 'dummy-xlam-v1', [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText("$machineDir\watchdog.ps1", 'dummy-wd-v1', [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText("$machineDir\watchdog_launcher.vbs", 'dummy-vbs', [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText("$machineDir\watchdog.conf", "ParentAllowlist=sjp7.exe`r`n", [Text.UTF8Encoding]::new($false))
    # 安装器日志（payload 上一级 = $g）
    [IO.File]::WriteAllText("$g\AttachLinkSetup.log", "===== AttachLinkSetup /install-all =====`r`n[结果] 一键安装完成（退出码 0）`r`n", [Text.UTF8Encoding]::new($false))

    # --- 场景 1：半部署（有目录、无任务）+ 无受信任位置 → 有 fail，退出码 10 ---
    '11-24 10:00:00 看门狗启动', '11-24 10:00:01 附着失败 hwnd=1 : 测试错误（退避重试）', '11-24 10:00:02 装载失败（Workbooks.Open）: 测试用严重错误' | Set-Content $gLog -Encoding UTF8
    $out1 = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $hcPath `
        -LogPath $gLog -PayloadDir $payloadDir -OfficeRootKey $officeRoot `
        -TaskName 'AttachLinkHealthTest-NoSuchTask' -MachineDir $machineDir)
    $code1 = $LASTEXITCODE
    $p1 = Parse-HcOutput $out1
    Check ($p1.Bad.Count -eq 0) ('G1 协议纯净：所有输出行都是协议行；杂行=' + ($p1.Bad -join ' | '))
    Check ($null -ne $p1.Summary) 'G1 有 SUMMARY 行'
    Check ((Get-St $p1 'machine.deploy') -eq 'fail') 'G1 machine.deploy：半部署 → fail'
    Check ((Get-St $p1 'machine.task') -eq 'fail') 'G1 machine.task：任务不存在 → fail'
    Check ((Get-St $p1 'machine.taskresult') -eq 'skip') 'G1 machine.taskresult：无任务 → skip'
    Check ((Get-St $p1 'machine.heartbeat') -eq 'ok') 'G1 machine.heartbeat：日志新鲜 → ok'
    Check ((Get-St $p1 'machine.files') -eq 'ok') 'G1 machine.files：三件齐全 → ok'
    Check ((Get-St $p1 'machine.version') -eq 'ok') 'G1 machine.version：与安装包一致 → ok'
    Check ((Get-St $p1 'machine.trusted') -eq 'fail') 'G1 machine.trusted：HKLM 无登记 → fail'
    $confItem = @($p1.Items | Where-Object { $_.Id -eq 'machine.conf' } | Select-Object -First 1)
    Check ($confItem.Count -eq 1 -and $confItem[0].Status -eq 'info' -and $confItem[0].Message.Contains('sjp7.exe')) 'G1 machine.conf：白名单可读 → info 且消息含内容'
    Check ((Get-St $p1 'user.open') -eq 'ok') 'G1 user.open：注入的库目录登记 → ok（OfficeRootKey 注入生效；OPENROAM 诱饵键不影响）'
    Check ((Get-St $p1 'user.trusted') -eq 'fail') 'G1 user.trusted：注入根里无信任位置 → fail'
    $dlog = @($p1.Items | Where-Object { $_.Id -eq 'diag.log' } | Select-Object -First 1)
    Check ($dlog.Count -eq 1 -and $dlog[0].Status -eq 'info' -and $dlog[0].Message.Contains('wd.log') -and ($dlog[0].Message -match '共 \d+ 行')) 'G1 diag.log：info 且含路径与总行数'
    $le1 = @($p1.Items | Where-Object { $_.Id -eq 'diag.logerrors' } | Select-Object -First 1)
    Check ($le1.Count -eq 1 -and $le1[0].Status -eq 'warn' -and $le1[0].Message.Contains('1 条错误记录') -and $le1[0].Message.Contains('1 条自愈类重试')) 'G1 diag.logerrors：严重错误 1 条 → warn，且区分自愈类重试 1 条'
    Check ($p1.Hints.ContainsKey('diag.logerrors') -and $p1.Hints['diag.logerrors'].Contains('装载失败') -and -not $p1.Hints['diag.logerrors'].Contains('退避重试')) 'G1 diag.logerrors：HCHINT 只带严重错误的原始行'
    $olItems = @($p1.Items | Where-Object { $_.Id -eq 'diag.openlist' })
    Check ($olItems.Count -eq 2 -and @($olItems | Where-Object { $_.Status -ne 'info' }).Count -eq 0) 'G1 diag.openlist：逐条一行（OPEN/OPEN1 两条，均 info）'
    Check (-not ((($olItems | ForEach-Object { $_.Message }) -join ' ').Contains('OPENROAM'))) 'G1 diag.openlist：诱饵键 OPENROAM 不出现在快照'
    Check (@($olItems | Where-Object { $_.Message -match '^OPEN=' }).Count -eq 1 -and @($olItems | Where-Object { $_.Message -match '^OPEN1=' }).Count -eq 1) 'G1 diag.openlist：每条带值名前缀（OPEN= / OPEN1=）'
    $setupItem = @($p1.Items | Where-Object { $_.Id -eq 'diag.setuplog' } | Select-Object -First 1)
    Check ($setupItem.Count -eq 1 -and $setupItem[0].Status -eq 'info' -and $setupItem[0].Message.Contains('一键安装完成')) 'G1 diag.setuplog：取到 [结果] 行'
    Check ((Get-St $p1 'diag.openlist') -eq 'info') 'G1 diag.openlist：OPEN 快照 → info'
    Check ($code1 -eq 10) ('G1 退出码：有 fail → 10；实际 ' + $code1)
    # 汇总行计数与逐项重算一致
    $recount = @{ ok = 0; warn = 0; fail = 0; skip = 0; info = 0 }
    foreach ($it in $p1.Items) { $recount[$it.Status]++ }
    Check ($p1.Summary -eq ("ok={0} warn={1} fail={2} skip={3} info={4}" -f $recount.ok, $recount.warn, $recount.fail, $recount.skip, $recount.info)) ('G1 SUMMARY 计数与逐项一致：' + $p1.Summary)
    # 只读核验：注入的注册表值原样未动
    Check ((Get-ItemProperty -LiteralPath "$officeRoot\16.0\Excel\Options").OPEN1 -eq '"AttachLink.xlam"') 'G1 只读：OPEN 值运行后原样'

    # --- 场景 2：未部署机器级（无目录无任务）+ 用户级齐全 + 干净日志 → 机器级全 skip ---
    Remove-Item "$officeRoot\16.0\Excel\Options" -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -Path "$officeRoot\16.0\Excel\Options" -Force | Out-Null
    New-ItemProperty -LiteralPath "$officeRoot\16.0\Excel\Options" -Name 'OPEN' -Value '"AttachLink.xlam"' -PropertyType String | Out-Null
    New-Item -Path "$officeRoot\16.0\Excel\Security\Trusted Locations\AttachLink" -Force | Out-Null
    New-ItemProperty -LiteralPath "$officeRoot\16.0\Excel\Security\Trusted Locations\AttachLink" -Name 'Path' -Value "$env:APPDATA\Microsoft\AddIns\" -PropertyType String | Out-Null
    '11-24 10:05:00 看门狗启动' | Set-Content $gLog -Encoding UTF8
    $noDir = Join-Path $g 'NoSuchMachineDir'
    $out2 = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $hcPath `
        -LogPath $gLog -PayloadDir $payloadDir -OfficeRootKey $officeRoot `
        -TaskName 'AttachLinkHealthTest-NoSuchTask' -MachineDir $noDir)
    $code2 = $LASTEXITCODE
    $p2 = Parse-HcOutput $out2
    Check ($p2.Bad.Count -eq 0) 'G2 协议纯净'
    foreach ($mid in @('machine.deploy', 'machine.task', 'machine.taskresult', 'machine.heartbeat', 'machine.files', 'machine.version', 'machine.trusted', 'machine.conf')) {
        Check ((Get-St $p2 $mid) -eq 'skip') "G2 $mid：未部署 → skip"
    }
    $depItem = @($p2.Items | Where-Object { $_.Id -eq 'machine.deploy' } | Select-Object -First 1)
    Check ($depItem.Count -eq 1 -and $depItem[0].Message.Contains('未部署机器级看门狗')) 'G2 machine.deploy：消息说明「未部署属正常」'
    Check ((Get-St $p2 'user.trusted') -eq 'ok') 'G2 user.trusted：注入信任位置 → ok'
    Check ((Get-St $p2 'diag.logerrors') -eq 'info') 'G2 diag.logerrors：干净日志 → info'
    # G2b：只有自愈类瞬态（附着失败退避重试 / 无法确认在位 / Window.Visible 暂时读不到）→ info，不算需留意
    '11-24 10:06:00 附着失败 hwnd=5 : 被呼叫方拒绝接收呼叫（退避重试）', '11-24 10:06:02 无法确认在位的 AttachLink.xlam（查询出错或路径不可读），本周期不处理: hwnd=5 pid=7', '11-24 10:06:04 Window.Visible 读不到，本周期不处理（下周期重试）: hwnd=5 pid=7' | Set-Content $gLog -Encoding UTF8
    $out2b = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $hcPath `
        -LogPath $gLog -PayloadDir $payloadDir -OfficeRootKey $officeRoot `
        -TaskName 'AttachLinkHealthTest-NoSuchTask' -MachineDir $noDir)
    $p2b = Parse-HcOutput $out2b
    $le2b = @($p2b.Items | Where-Object { $_.Id -eq 'diag.logerrors' } | Select-Object -First 1)
    Check ($le2b.Count -eq 1 -and $le2b[0].Status -eq 'info' -and $le2b[0].Message.Contains('自愈类重试')) 'G2b diag.logerrors：只有自愈类重试 → info（不算需留意）'
    Check (-not $p2b.Hints.ContainsKey('diag.logerrors')) 'G2b diag.logerrors：info 不带处置提示（不进入需留意清单）'
    # G2c：升级告警「已持续 60 秒读不到」不是瞬态 → warn
    '11-24 10:07:00 Window.Visible 已持续 60 秒读不到，该窗口一直没有被处理: hwnd=5 pid=7' | Set-Content $gLog -Encoding UTF8
    $out2c = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $hcPath `
        -LogPath $gLog -PayloadDir $payloadDir -OfficeRootKey $officeRoot `
        -TaskName 'AttachLinkHealthTest-NoSuchTask' -MachineDir $noDir)
    Check ((Get-St (Parse-HcOutput $out2c) 'diag.logerrors') -eq 'warn') 'G2c diag.logerrors：升级告警「已持续 60 秒读不到」→ warn'
    # G2d（B-3）：200 条相同「退避重试」→ 持续附着失败，升级 warn（不得永久归为瞬态 info）
    $linesG2d = 1..200 | ForEach-Object { '11-24 10:08:00 附着失败 hwnd=5 : 测试错误（退避重试）' }
    $linesG2d | Set-Content $gLog -Encoding UTF8
    $out2d = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $hcPath `
        -LogPath $gLog -PayloadDir $payloadDir -OfficeRootKey $officeRoot `
        -TaskName 'AttachLinkHealthTest-NoSuchTask' -MachineDir $noDir)
    $p2d = Parse-HcOutput $out2d
    $le2d = @($p2d.Items | Where-Object { $_.Id -eq 'diag.logerrors' } | Select-Object -First 1)
    Check ($le2d.Count -eq 1 -and $le2d[0].Status -eq 'warn' -and $le2d[0].Message.Contains('看门狗持续附着失败') -and $le2d[0].Message.Contains('200 条退避重试')) ('G2d diag.logerrors：200 条相同重试 → warn（实际「' + $le2d[0].Message + '」）')
    # G2e（B-3）：12 条重试、最早与最晚跨度 15 分钟（≥10 分钟且 ≥10 条）→ warn
    $mins = @(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15)
    $linesG2e = $mins | ForEach-Object { '11-24 10:{0:D2}:00 附着失败 hwnd=5 : 测试错误（退避重试）' -f $_ }
    $linesG2e | Set-Content $gLog -Encoding UTF8
    $out2e = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $hcPath `
        -LogPath $gLog -PayloadDir $payloadDir -OfficeRootKey $officeRoot `
        -TaskName 'AttachLinkHealthTest-NoSuchTask' -MachineDir $noDir)
    $p2e = Parse-HcOutput $out2e
    $le2e = @($p2e.Items | Where-Object { $_.Id -eq 'diag.logerrors' } | Select-Object -First 1)
    Check ($le2e.Count -eq 1 -and $le2e[0].Status -eq 'warn' -and $le2e[0].Message.Contains('跨越约 15 分钟')) ('G2e diag.logerrors：12 条跨度 15 分钟 → warn（实际「' + $le2e[0].Message + '」）')
    # G2f（B-3 反向断言）：9 条、跨度 20 分钟——条数 < 10 → 仍是 info（阈值是「条数 ≥ 10 且跨度 ≥ 10 分钟」）
    $mins9 = @(0, 2, 4, 6, 8, 10, 12, 16, 20)
    $linesG2f = $mins9 | ForEach-Object { '11-24 10:{0:D2}:00 附着失败 hwnd=5 : 测试错误（退避重试）' -f $_ }
    $linesG2f | Set-Content $gLog -Encoding UTF8
    $out2f = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $hcPath `
        -LogPath $gLog -PayloadDir $payloadDir -OfficeRootKey $officeRoot `
        -TaskName 'AttachLinkHealthTest-NoSuchTask' -MachineDir $noDir)
    Check ((Get-St (Parse-HcOutput $out2f) 'diag.logerrors') -eq 'info') 'G2f diag.logerrors：9 条（<10）即使跨度 20 分钟 → 仍 info'
    # G3（NB-1b）：计划任务查询抛异常（注入）→ machine.deploy / machine.task = warn「无法确认」，其余机器项 skip 且消息可见「无法确认」
    '11-24 10:20:00 看门狗启动' | Set-Content $gLog -Encoding UTF8
    $out3 = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $hcPath `
        -LogPath $gLog -PayloadDir $payloadDir -OfficeRootKey $officeRoot `
        -TaskName 'AttachLinkHealthTest-NoSuchTask' -MachineDir $noDir `
        -TaskQueryInjectError '服务不可用（模拟查询异常）')
    $code3 = $LASTEXITCODE
    $p3 = Parse-HcOutput $out3
    Check ($p3.Bad.Count -eq 0) 'G3 协议纯净'
    $dep3 = @($p3.Items | Where-Object { $_.Id -eq 'machine.deploy' } | Select-Object -First 1)
    Check ($dep3.Count -eq 1 -and $dep3[0].Status -eq 'warn' -and $dep3[0].Message.Contains('无法确认')) ('G3 machine.deploy：任务查询失败 → warn「无法确认」，不得判「未部署/半部署」（实际「' + $dep3[0].Status + ' ' + $dep3[0].Message + '」）')
    $task3 = @($p3.Items | Where-Object { $_.Id -eq 'machine.task' } | Select-Object -First 1)
    Check ($task3.Count -eq 1 -and $task3[0].Status -eq 'warn' -and $task3[0].Message.Contains('无法确认计划任务（查询失败：')) ('G3 machine.task：查询失败 → warn「无法确认计划任务」（实际「' + $task3[0].Status + ' ' + $task3[0].Message + '」）')
    foreach ($mid in @('machine.taskresult', 'machine.heartbeat', 'machine.files', 'machine.version', 'machine.trusted', 'machine.conf')) {
        $it3 = @($p3.Items | Where-Object { $_.Id -eq $mid } | Select-Object -First 1)
        Check ($it3.Count -eq 1 -and $it3[0].Status -eq 'skip' -and $it3[0].Message.Contains('无法确认')) ("G3 $mid：依赖任务的项保持 skip，消息能看出「无法确认」（实际「$($it3[0].Status) $($it3[0].Message)」）")
    }
    Check ($code3 -eq 0) ('G3 退出码：无 fail → 0（实际 ' + $code3 + '）')
    $dlog2 = @($p2.Items | Where-Object { $_.Id -eq 'diag.setuplog' } | Select-Object -First 1)
    Check ($dlog2.Count -eq 1 -and $dlog2[0].Status -eq 'info' -and $dlog2[0].Message.Contains('一键安装完成')) 'G2 diag.setuplog：取到 [结果] 行'
    $fail2 = @($p2.Items | Where-Object { $_.Status -eq 'fail' }).Count
    $expect2 = $(if ($fail2 -gt 0) { 10 } else { 0 })
    Check ($code2 -eq $expect2) ('G2 退出码与 fail 数自洽：fail=' + $fail2 + ' exit=' + $code2 + '（user.* 依赖真机状态，不断言具体值）')
} finally {
    Remove-Item $regRoot -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $g -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
if ($ok) { 'RESULT: PASS'; exit 0 } else { 'RESULT: FAIL'; exit 1 }
