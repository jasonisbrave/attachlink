# =====================================================================
# AttachLink 机器级看门狗部署脚本（管理员运行一次）
# ---------------------------------------------------------------------
# 用途：鼎信诺等审计软件以 COM 自动化（CreateObject）拉起 Excel 时，Excel
# 跳过一切加载项初始化（.xlam 的 OPEN 注册、XLSTART、COM 加载项 OnConnection
# 均不触发，本机逐项实测证实）；且鼎信诺常以管理员运行，高完整性进程忽略
# HKCU per-user 注册。因此装载必须由一个外部常驻进程完成——看门狗
# （scripts\watchdog.ps1）周期枚举 Excel EXCEL7 文档窗口，通过
# AccessibleObjectFromWindow 取得 Application，强制装载 AttachLink.xlam。
#
# 本脚本做四件事：
#   1. 部署文件到 C:\Program Files\AttachLink\（xlam + 看门狗 + 隐藏启动器）
#   2. HKLM 受信任位置指向该目录（管理员控制目录，规避提权宏劫持面；
#      高完整性 Excel 不受 HKCU per-user 信任位置保障，必须走 HKLM）
#   3. 注册登录计划任务（当前用户，RunLevel=Highest）：看门狗以最高权限
#      常驻，可同时附着普通 Excel 与管理员权限的鼎信诺 Excel
#   4. 立即启动看门狗（无需注销重登），并清理旧 COM 引导器注册残留
#
# 运行方式（管理员 PowerShell）：
#   powershell -ExecutionPolicy Bypass -File install-elevated.ps1
# 卸载：
#   powershell -ExecutionPolicy Bypass -File install-elevated.ps1 -Uninstall
#
# 退出码：0 成功；8 机器级卸载后有残留；12 使用了其他管理员账户授权（安装未执行）。
#
# 前置：普通用户级安装（install.ps1）已完成——本脚本只负责鼎信诺场景的
# 看门狗；日常双击打开的 Excel 由用户级安装直接覆盖。
# =====================================================================
param(
    [switch]$Uninstall,
    # 看门狗父进程白名单（逗号分隔进程名，如 "sjp7.exe"）：仅附着由这些进程拉起的 Excel，
    # 收敛无差别装载面（Issue #1 / T13）。默认空 = 不写配置、保留既有 watchdog.conf（升级 /
    # 一键更新不动用户已有的加固配置）；要显式清除白名单用 -ClearWatchdogParentFilter。
    # 部署后也可直接编辑 C:\Program Files\AttachLink\watchdog.conf（ParentAllowlist= 行）。
    [string]$WatchdogParentFilter = '',
    [switch]$ClearWatchdogParentFilter,
    # 交互登录用户的 SID（调用方一般不传）：未传时自动探测与本进程同 Session 的 explorer.exe
    # 的所有者 SID。提权账户 ≠ 交互登录用户时安装分支如实失败（exit 12），不假装「已部署」。
    [string]$InteractiveUserSid = ''
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$machineDir = 'C:\Program Files\AttachLink'
$taskName   = 'AttachLink Watchdog'
$user       = "$env:USERDOMAIN\$env:USERNAME"
$officeVersions = @('16.0', '15.0', '14.0')

function Get-WatchdogProcessObjects {
    # 按命令行精确匹配看门狗进程（避免误伤其他 powershell/wscript）
    return @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='wscript.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like '*AttachLink*watchdog*' })
}

function Stop-WatchdogProcesses {
    Get-WatchdogProcessObjects | ForEach-Object {
        try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop } catch {}
    }
}

function Remove-LegacyShim {
    # 旧 COM 引导器注册残留（该路线已废弃，直接删键；DLL 不存在时 RegAsm 无法反注册）
    foreach ($k in @(
        'HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\AttachLinkShim.Connect',
        'HKLM:\SOFTWARE\Microsoft\Office\Excel\AddinsData\AttachLinkShim.Connect',
        'HKCU:\Software\Microsoft\Office\Excel\Addins\AttachLinkShim.Connect',
        'HKCU:\Software\Microsoft\Office\Excel\AddinsData\AttachLinkShim.Connect',
        'Registry::HKEY_CLASSES_ROOT\AttachLinkShim.Connect',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58}',
        'HKCU:\Software\Classes\AttachLinkShim.Connect',
        'HKCU:\Software\Classes\CLSID\{D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58}'
    )) { Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue }
}

function Get-InteractiveUserSid {
    # 探测交互登录用户的 SID：取与本进程同一 SessionId 的 explorer.exe，读其所有者 SID。
    # 探测不到返回 $null（「未知」——调用方保持现有行为，只多打印一行提示）。
    try {
        $mySession = (Get-Process -Id $PID -ErrorAction Stop).SessionId
        $exp = @(Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue |
            Where-Object { $_.SessionId -eq $mySession } | Select-Object -First 1)
        if ($exp.Count -eq 0) { return $null }
        $r = Invoke-CimMethod -InputObject $exp[0] -MethodName GetOwnerSid -ErrorAction Stop
        return [string]$r.Sid
    } catch { return $null }
}

function Test-SameUserSid([string]$InteractiveSid, [string]$ElevatedSid) {
    # 提权令牌用户 SID 与交互登录用户 SID 是否一致（用 SID 比较，不比较用户名——
    # 微软账户 / 域账户的名字格式不可靠）。任一侧未知 → 返回 $true（不拦，由调用方提示）。
    if (-not $InteractiveSid -or -not $ElevatedSid) { return $true }
    return ($InteractiveSid -ieq $ElevatedSid)
}

function Get-MachineUninstallResidue([string]$Task, [string]$Dir,
    [string]$TrustedRoot = 'HKLM:\SOFTWARE\Microsoft\Office', [string[]]$OfficeVers = $officeVersions,
    [scriptblock]$GetTask = { param($n) @(Get-ScheduledTask -TaskName $n -ErrorAction SilentlyContinue) },
    [scriptblock]$TestDir = { param($p) Test-Path -LiteralPath $p },
    [scriptblock]$GetWdProcs = { Get-WatchdogProcessObjects }) {
    # 机器级卸载收尾核验（council C1/C4）：删除动作全部做完之后按事实收集残留清单，
    # 每条形如「名称：路径（原因 / 处置）」。空数组 = 干净；否则卸载入口必须非零退出。
    $residue = @()
    $taskHit = @(& $GetTask $Task)
    if ($taskHit.Count -gt 0) {
        $residue += "计划任务：$Task（任务仍存在；处置：管理员 PowerShell 执行 Unregister-ScheduledTask -TaskName '$Task' -Confirm:`$false）"
    }
    if (& $TestDir $Dir) {
        $residue += "部署目录：$Dir（目录仍存在，可能被占用；处置：关闭全部 Excel 后重新运行卸载，仍失败请手动删除该目录）"
    }
    foreach ($v in @($OfficeVers)) {
        $k = Join-Path $TrustedRoot "$v\Excel\Security\Trusted Locations\AttachLink"
        if (Test-Path -LiteralPath $k) {
            $residue += "HKLM 受信任位置：$k（注册表键未能删除；处置：请手动删除该键）"
        }
    }
    $wd = @(& $GetWdProcs)
    if ($wd.Count -gt 0) {
        $residue += ('看门狗进程仍在运行（pid ' + (($wd | ForEach-Object { $_.ProcessId }) -join '、') + '）；处置：结束这些进程或重启后重新运行卸载')
    }
    return $residue
}

# =====================================================================
# 主逻辑（点源载入时不执行——单元测试只取上面的函数）
# =====================================================================
if ($MyInvocation.InvocationName -ne '.') {
    # 必须管理员
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { throw '本脚本需要管理员权限：请在管理员 PowerShell 中运行。' }

    # 探测 Office 主版本（HKLM 受信任位置用）
    $ver = $null
    foreach ($v in $officeVersions) {
        if (Test-Path "HKCU:\Software\Microsoft\Office\$v\Excel") { $ver = $v; break }
    }
    if (-not $ver) {
        foreach ($v in $officeVersions) {
            if (Test-Path "HKLM:\Software\Microsoft\Office\$v\Excel") { $ver = $v; break }
        }
    }
    if (-not $ver) { throw '未检测到已安装的桌面版 Excel。' }
    $tl = "HKLM:\SOFTWARE\Microsoft\Office\$ver\Excel\Security\Trusted Locations\AttachLink"

    if (-not $Uninstall) {
        # ---- 0) 提权账户 ≠ 交互登录用户：如实失败（council C4 双视角命中）----
        # 标准用户用别的管理员账号过 UAC 时，任务只能登记在该管理员名下，真正使用者登录时
        # 看门狗不启动。必须在改动任何文件 / 注册表 / 计划任务之前拦下（exit 12），不能假报「已部署」。
        if (-not $InteractiveUserSid) { $InteractiveUserSid = Get-InteractiveUserSid }
        $elevatedSid = ''
        try { $elevatedSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value } catch {}
        if (-not $InteractiveUserSid) {
            Write-Host '提示：未能探测到当前登录用户的 SID，将按现有行为继续（看门狗任务登记在提权账户名下）。'
        } elseif (-not (Test-SameUserSid $InteractiveUserSid $elevatedSid)) {
            Write-Host '你使用了另一个管理员账户授权。看门狗计划任务只能登记在该管理员名下，'
            Write-Host '当前登录的用户登录时不会启动，鼎信诺环境对当前用户无效。'
            Write-Host '请用当前登录账户的管理员身份重新授权；如果当前账户不是管理员，请联系维护者。'
            exit 12
        }

        $root     = Split-Path -Parent $PSScriptRoot
        $distXlam = Join-Path $root 'dist\AttachLink.xlam'
        $srcWd    = Join-Path $root 'scripts\watchdog.ps1'
        $srcVbs   = Join-Path $root 'scripts\watchdog_launcher.vbs'
        foreach ($f in @($distXlam, $srcWd, $srcVbs)) {
            if (-not (Test-Path $f)) { throw "未找到 $f 。" }
        }

        # ---- 1) 停旧看门狗、清理旧 COM 引导器、部署文件 ----
        Stop-WatchdogProcesses
        Remove-LegacyShim

        New-Item -ItemType Directory -Force -Path $machineDir | Out-Null
        Copy-Item $distXlam (Join-Path $machineDir 'AttachLink.xlam') -Force
        Copy-Item $srcWd    (Join-Path $machineDir 'watchdog.ps1') -Force
        Copy-Item $srcVbs   (Join-Path $machineDir 'watchdog_launcher.vbs') -Force
        Get-ChildItem $machineDir -File | ForEach-Object { Unblock-File $_.FullName -ErrorAction SilentlyContinue }

        # ---- 1b) 看门狗父进程过滤配置（可选；不写参数时保留既有 watchdog.conf，不清用户的加固配置）----
        $confPath = Join-Path $machineDir 'watchdog.conf'
        if ($WatchdogParentFilter) {
            "ParentAllowlist=$WatchdogParentFilter" | Set-Content -Path $confPath -Encoding ASCII
            Write-Host "看门狗父进程白名单已写入：$WatchdogParentFilter（仅附着这些进程拉起的 Excel；清空删除 watchdog.conf 即恢复不过滤）"
        } elseif ($ClearWatchdogParentFilter) {
            Remove-Item $confPath -Force -ErrorAction SilentlyContinue
            Write-Host '已删除看门狗父进程白名单配置（watchdog.conf），恢复默认不过滤。'
        }

        # ---- 2) HKLM 受信任位置（仅该目录本身；AllowSubFolders=0 最小信任范围）----
        New-Item -Path $tl -Force | Out-Null
        Set-ItemProperty -Path $tl -Name 'Path' -Value "$machineDir\"
        Set-ItemProperty -Path $tl -Name 'AllowSubFolders' -Value 0 -Type DWORD
        Set-ItemProperty -Path $tl -Name 'Date' -Value (Get-Date -Format 'MM/dd/yyyy')
        Set-ItemProperty -Path $tl -Name 'Description' -Value 'AttachLink watchdog deployment (machine)'

        # ---- 3) 登录计划任务：最高权限常驻看门狗（隐藏窗口，经 VBS 启动避免控制台闪现）----
        $action    = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + (Join-Path $machineDir 'watchdog_launcher.vbs') + '"')
        $trigger   = New-ScheduledTaskTrigger -AtLogOn -User $user
        $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
        # 长驻进程：不限执行时长（PT0S = 禁用时限）；失败自动重启最多 3 次；已运行则忽略新实例
        $settings  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Seconds 0) `
            -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        # 回读登记账户并明示（提权账户 ≠ 交互用户时的「假部署」要能一眼看穿）
        try { Write-Host ("  任务登记账户 : " + [string](Get-ScheduledTask -TaskName $taskName).Principal.UserId) } catch {}

        # ---- 4) 立即启动（无需注销重登）----
        Start-ScheduledTask -TaskName $taskName
        Start-Sleep -Seconds 2
        $state = (Get-ScheduledTask -TaskName $taskName).State

        # 版本核对：部署副本与 dist 必须一致（升级漏跑本脚本会留下旧版机器级副本）。
        # .NET 直调哈希：安装器 exe 的子 PowerShell 环境可能无法加载 Utility 模块
        $shaProv = [System.Security.Cryptography.SHA256]::Create()
        $fsA = [IO.File]::OpenRead((Join-Path $root 'dist\AttachLink.xlam'))
        try { $hDist = [BitConverter]::ToString($shaProv.ComputeHash($fsA)) -replace '-', '' } finally { $fsA.Close() }
        $fsB = [IO.File]::OpenRead((Join-Path $machineDir 'AttachLink.xlam'))
        try { $hDest = [BitConverter]::ToString($shaProv.ComputeHash($fsB)) -replace '-', '' } finally { $fsB.Close(); $shaProv.Dispose() }
        $verNote = if ($hDist -eq $hDest) { "与 dist 一致（SHA-256 前16位 $($hDist.Substring(0,16))）" } else { "警告：与 dist 不一致！升级时请务必重跑本脚本同步机器级副本" }

        Write-Host '机器级看门狗部署完成。'
        Write-Host ("  部署目录 : $machineDir")
        Write-Host ("  受信任位置(HKLM) : $tl")
        Write-Host ("  计划任务 : $taskName（登录自启，最高权限，当前任务状态: $state）")
        Write-Host ("  xlam 版本核对 : $verNote")
        Write-Host '验证：%TEMP%\AttachLinkWatchdog.log 应出现「看门狗启动」且无异常。'
        Write-Host '完全退出并重新打开鼎信诺后生效：打开底稿约 2~4 秒内右键菜单出现。'
    }
    else {
        # ---- 卸载：停任务 → 删任务 → 杀进程 → 清注册表（全版本）→ 删目录 → 按事实核验 ----
        Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        Stop-WatchdogProcesses
        Remove-LegacyShim
        # 受信任位置：16.0 / 15.0 / 14.0 全删（只删一个版本会让其它版本的旧键成为假成功的残留）
        foreach ($v in $officeVersions) {
            Remove-Item "HKLM:\SOFTWARE\Microsoft\Office\$v\Excel\Security\Trusted Locations\AttachLink" -Recurse -Force -ErrorAction SilentlyContinue
        }
        Remove-Item $machineDir -Recurse -Force -ErrorAction SilentlyContinue

        # 按事实核验：有残留 → 打印清单并 exit 8，不再无条件报「卸载完成」
        $residue = @(Get-MachineUninstallResidue -Task $taskName -Dir $machineDir)
        if ($residue.Count -gt 0) {
            Write-Host '⚠ 机器级卸载未完全成功，以下项目未能清除：'
            foreach ($r in $residue) { Write-Host ('  - ' + $r) }
            Write-Host '请关闭全部 Excel 后重新运行本脚本 -Uninstall；仍失败则按上方清单手动处理。'
            exit 8
        }
        Write-Host '机器级看门狗卸载完成（计划任务 / 部署目录 / HKLM 信任位置 / 看门狗进程均已核验清除；用户级安装不受影响，如需一并卸载请运行 install.ps1 -Uninstall）。'
    }
}
