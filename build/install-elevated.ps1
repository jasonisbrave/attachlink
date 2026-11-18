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
# 前置：普通用户级安装（install.ps1）已完成——本脚本只负责鼎信诺场景的
# 看门狗；日常双击打开的 Excel 由用户级安装直接覆盖。
# =====================================================================
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# 必须管理员
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { throw '本脚本需要管理员权限：请在管理员 PowerShell 中运行。' }

$machineDir = 'C:\Program Files\AttachLink'
$taskName   = 'AttachLink Watchdog'
$user       = "$env:USERDOMAIN\$env:USERNAME"

# 探测 Office 主版本（HKLM 受信任位置用）
$ver = $null
foreach ($v in @('16.0', '15.0', '14.0')) {
    if (Test-Path "HKCU:\Software\Microsoft\Office\$v\Excel") { $ver = $v; break }
}
if (-not $ver) {
    foreach ($v in @('16.0', '15.0', '14.0')) {
        if (Test-Path "HKLM:\Software\Microsoft\Office\$v\Excel") { $ver = $v; break }
    }
}
if (-not $ver) { throw '未检测到已安装的桌面版 Excel。' }
$tl = "HKLM:\SOFTWARE\Microsoft\Office\$ver\Excel\Security\Trusted Locations\AttachLink"

function Stop-WatchdogProcesses {
    # 按命令行精确匹配看门狗进程并结束（避免误伤其他 powershell/wscript）
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='wscript.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like '*AttachLink*watchdog*' } |
        ForEach-Object {
            try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop } catch {}
        }
}

function Remove-LegacyShim {
    # 旧 COM 引导器注册残留（该路线已废弃，直接删键；DLL 不存在时 RegAsm 无法反注册）
    foreach ($k in @(
        'HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\AttachLinkShim.Connect',
        'HKCU:\Software\Microsoft\Office\Excel\Addins\AttachLinkShim.Connect',
        'Registry::HKEY_CLASSES_ROOT\AttachLinkShim.Connect',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58}',
        'HKCU:\Software\Classes\AttachLinkShim.Connect',
        'HKCU:\Software\Classes\CLSID\{D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58}'
    )) { Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue }
}

if (-not $Uninstall) {
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
    # ---- 卸载：停任务 → 删任务 → 杀进程 → 清注册表 → 删目录 ----
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Stop-WatchdogProcesses
    Remove-LegacyShim
    Remove-Item $tl -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $machineDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host '机器级看门狗卸载完成（用户级安装不受影响，如需一并卸载请运行 install.ps1 -Uninstall）。'
}
