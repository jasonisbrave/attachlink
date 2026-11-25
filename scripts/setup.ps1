# =====================================================================
# AttachLink 安装 / 卸载 编排脚本（备用命令行入口）
# 主入口为 AttachLinkSetup.exe（安装器）；本脚本供 exe 被安全软件拦截时
# 手动运行，也可被其他工具编排调用：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\setup.ps1             安装
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\setup.ps1 -Uninstall  卸载
# =====================================================================
param(
    [switch]$Uninstall,
    [switch]$NonInteractive   # GUI/静默调用：所有交互点改为直接判定（Excel 在跑=失败退出；缺 dist=失败退出）
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$root     = Split-Path -Parent $PSScriptRoot
$buildDir = Join-Path $root 'build'
$distXlam = Join-Path $root 'dist\AttachLink.xlam'
$preXlam  = Join-Path $env:APPDATA 'Microsoft\AddIns\AttachLink.xlam'

function Get-BlockingExcelProcesses {
    # 存活的 Excel（WorkingSet64 > 1MB，过滤已退出但被句柄持有的墓碑残影）
    return @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $_.WorkingSet64 -gt 1MB })
}
function Test-AddInFileLocked([string]$Path) {
    # 加载项文件被占用判定：以「拒绝一切共享」的方式尝试打开——共享冲突或其它 IO 异常都算占用；
    # 文件不存在则不算占用。
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $fs.Close()
        return $false
    } catch { return $true }
}
function Get-UninstallExitCode([bool]$MachineResidue, [bool]$UserResidue, [bool]$UserListOpenFail) {
    # 卸载结论合成（按事实，优先级 8 > 10 > 9 > 0）：
    #   8 机器级有残留；10 用户级有残留（文件 / 注册表 / 收尾核验发现的任何残留）；
    #   9 仅用户级加载项列表登记未清净；0 全部核验干净。
    if ($MachineResidue) { return 8 }
    if ($UserResidue) { return 10 }
    if ($UserListOpenFail) { return 9 }
    return 0
}

Write-Host ''
Write-Host '==================================================='
if ($Uninstall) {
    Write-Host '  AttachLink 审计底稿附件链接加载项 —— 一键卸载'
} else {
    Write-Host '  AttachLink 审计底稿附件链接加载项 —— 一键安装'
}
Write-Host '==================================================='

if (-not $Uninstall) {
    # ---- 1. Excel 运行检查 ----
    while ($true) {
        $xl = Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $_.WorkingSet64 -gt 1MB }   # 过滤墓碑进程（已退出但被句柄持有的残影，内存趋零）——2026-11-21 实测误报修复
        if (-not $xl) { break }
        if ($NonInteractive) {
            Write-Host '>> 错误：检测到 Excel 正在运行。请先关闭所有 Excel（含鼎信诺）后重试。'
            exit 2
        }
        Write-Host ''
        Write-Host '>> 检测到 Excel 正在运行，安装前需要先关闭 Excel（否则文件会被占用）。'
        Write-Host '   （提示：如有其他软件在后台调用 Excel 无法关闭，可输入 S 跳过检查）'
        $ans = (Read-Host '   关闭所有 Excel 后直接回车继续；输入 S 跳过检查；输入 Q 退出').Trim().ToUpper()
        if ($ans -eq 'Q') { Write-Host '已取消安装。'; exit 1 }
        if ($ans -eq 'S') { break }
    }

    # ---- 2. 确保构建产物存在（分发包应自带；缺失则现场构建）----
    if (Test-Path $distXlam) {
        Write-Host ''
        Write-Host '>> 已找到安装包 dist\AttachLink.xlam'
    } else {
        Write-Host ''
        Write-Host '>> 未找到 dist\AttachLink.xlam，分发包可能不完整。'
        if ($NonInteractive) { exit 3 }
        Write-Host '   正常分发应自带预构建安装包；现场构建会临时开启「VBA 工程访问权限」，仅适合开发环境。'
        $ans = (Read-Host '   输入 B 强制从源码构建后继续；直接回车退出安装').Trim().ToUpper()
        if ($ans -ne 'B') {
            Write-Host '已退出安装。请向分发人索取包含 dist\AttachLink.xlam 的完整文件夹。'
            exit 1
        }
        & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $buildDir 'build.ps1')
        if ($LASTEXITCODE -ne 0) {
            throw '构建失败：请确认文件夹完整（含 src 与 build 目录）且本机已安装桌面版 Excel。'
        }
    }

    # ---- 3. 安装 ----
    Write-Host ''
    Write-Host '>> 正在安装：复制文件 -> 加入受信任位置 -> 注册启用加载项...'
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $buildDir 'install.ps1')
    if ($LASTEXITCODE -ne 0) { throw '安装失败，请把上方错误信息反馈给分发人。' }

    Write-Host ''
    Write-Host '==================================================='
    Write-Host '  安装完成！'
    Write-Host '  1. 完全退出并重新打开 Excel'
    Write-Host '  2. 在任意单元格上点右键，应能看到「插入附件链接（可多选）」'
    Write-Host '  -------------------------------------------------'
    Write-Host '  【用安装器的同事】点「一键安装」会自动完成下一步（部署鼎信诺环境）。'
    Write-Host '  【用脚本安装、且需要在本机以管理员运行的鼎信诺里编制底稿的同事】还需一步：'
    Write-Host '     关闭鼎信诺和 Excel 后，右键以管理员身份运行'
    Write-Host '     build\install-elevated.ps1（管理员 PowerShell 中执行：'
    Write-Host '     powershell -ExecutionPolicy Bypass -File build\install-elevated.ps1）'
    Write-Host '  -------------------------------------------------'
    Write-Host '  详细用法见文件夹中的 README.md'
    Write-Host '==================================================='
}
else {
    # ---- 0. 卸载前预检（council C1）：有存活的 Excel、或加载项文件被占用 → 不卸载 ----
    # 开着 Excel（或 WPS 表格等占用方）点卸载：文件删不掉，且旧会话退出时会把内存里的加载项登记
    # 写回注册表——界面却显示「卸载完成」。预检不通过绝不允许跳过。
    while ($true) {
        $liveXlPre = @(Get-BlockingExcelProcesses)
        $fileLocked = Test-AddInFileLocked $preXlam
        if ($liveXlPre.Count -eq 0 -and -not $fileLocked) { break }
        $reason = @()
        if ($liveXlPre.Count -gt 0) {
            $reason += ('检测到 Excel 正在运行（pid ' + (($liveXlPre | ForEach-Object { $_.Id }) -join '、') + '）')
        }
        if ($fileLocked) { $reason += "加载项文件被占用：$preXlam" }
        if ($NonInteractive) {
            Write-Host ('>> 错误：' + ($reason -join '；') + '。')
            Write-Host '>> 卸载未执行：请先关闭所有 Excel（含鼎信诺、WPS 表格）以及其它打开了该加载项的程序，然后重新运行卸载。'
            exit 2
        }
        Write-Host ''
        Write-Host ('>> 无法卸载：' + ($reason -join '；'))
        Write-Host '   请关闭所有 Excel（含鼎信诺、WPS 表格）以及其它打开了该加载项的程序。'
        $ans = (Read-Host '   关闭后直接回车重新检查；输入 Q 退出').Trim().ToUpper()
        if ($ans -eq 'Q') { Write-Host '已取消卸载。'; exit 1 }
    }

    Write-Host ''
    Write-Host '>> 正在一键卸载 AttachLink（用户级 + 机器级看门狗）...'
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $buildDir 'install.ps1') -Uninstall
    $userExit = $LASTEXITCODE
    $userLevelOpenFail = $false   # 仅「加载项列表登记未清净」（exit 9 的既有契约）
    $userResidue = $false         # 用户级按事实核验发现的任何其它残留（exit 10）
    if ($userExit -eq 9) {
        # oracle 三轮：子脚本「加载项列表清理未完全成功」需原样传播（GUI 退出码承诺 9），不得被 throw 吞成 1
        $userLevelOpenFail = $true
        Write-Host '注意：用户级卸载的加载项列表（OPEN / Add-in Manager）登记未能完全清理（见上方警告），继续机器级检查。'
    }
    elseif ($userExit -eq 10) {
        # 用户级有残留（副本 / 受信任位置 / 旧引导器 / 收尾仍有存活 Excel）：记下并按事实合成退出码，继续机器级检查
        $userResidue = $true
        Write-Host '注意：用户级卸载有残留（清单见上方），继续机器级检查。'
    }
    elseif ($userExit -ne 0) { throw '用户级卸载失败，请把上方错误信息反馈给分发人。' }
    if ($userExit -eq 0) {
        Write-Host ''
        Write-Host '用户级已卸载（日常 Excel 的加载项/受信任位置/副本已移除，且经核验无残留）。'
    }

    # ---- 机器级看门狗：检测并一并卸载（一键整合，2026-11-23）----
    # 旧版只打印手动提示，留给用户跑第二步。鼎信诺场景的同事只跑了用户级
    # 卸载时，看门狗继续把 Program Files 旧副本注入每个新 Excel 会话并重建
    # 右键菜单——「卸载后菜单仍在」的同事反馈即此（另有运行中会话内菜单
    # 保留到关闭的成分）。此处自动检测足迹，未提权则弹一次 UAC 完成卸载。
    $machineDir = 'C:\Program Files\AttachLink'
    $taskName   = 'AttachLink Watchdog'
    $tlHklm = $null
    foreach ($v in @('16.0', '15.0', '14.0')) {
        $k = "HKLM:\SOFTWARE\Microsoft\Office\$v\Excel\Security\Trusted Locations\AttachLink"
        if (Test-Path $k) { $tlHklm = $k; break }
    }
    $hasTask = [bool](Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)
    $hasDir  = Test-Path $machineDir
    $machineFailed = $false
    if (-not ($hasTask -or $hasDir -or $tlHklm)) {
        if (-not ($userResidue -or $userLevelOpenFail)) {
            Write-Host '未检测到机器级看门狗部署——本机无需第二步。'
        }
    }
    else {
        $machineFailed = $false
        $elevatedScript = Join-Path $buildDir 'install-elevated.ps1'
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($isAdmin) {
            & powershell -NoProfile -ExecutionPolicy Bypass -File $elevatedScript -Uninstall
            if ($LASTEXITCODE -ne 0) { Write-Host "机器级卸载退出码 $LASTEXITCODE，以下按实际状态核验。"; $machineFailed = $true }
        }
        else {
            Write-Host '检测到机器级看门狗部署（计划任务/Program Files/HKLM 信任位置），'
            Write-Host '即将弹出授权（UAC）以一并卸载，请在弹窗中点「是」...'
            try {
                # 单一参数行 + -File（提权传参坑：多元素拼接对含特殊字符命令静默失效）
                $p = Start-Process powershell -Verb RunAs -Wait -PassThru -WindowStyle Hidden `
                    -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -Uninstall' -f $elevatedScript)
                if ($p.ExitCode -ne 0) { Write-Host "提权卸载进程退出码 $($p.ExitCode)，以下按实际状态核验。"; $machineFailed = $true }
            } catch {
                Write-Host 'UAC 被取消或启动失败：机器级看门狗未卸载。'
                $machineFailed = $true
            }
        }
        # 事实核验（不依赖提权窗口输出，直接查残留）
        $leftTask = [bool](Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)
        $leftDir  = Test-Path $machineDir
        $leftTl   = $null
        foreach ($v in @('16.0', '15.0', '14.0')) {
            $k = "HKLM:\SOFTWARE\Microsoft\Office\$v\Excel\Security\Trusted Locations\AttachLink"
            if (Test-Path $k) { $leftTl = $k; break }
        }
        if ($leftTask -or $leftDir -or $leftTl) { $machineFailed = $true }
        if (-not $machineFailed) {
            Write-Host '机器级看门狗已卸载（计划任务/部署目录/HKLM 信任位置均已清除）。'
        }
        else {
            Write-Host '⚠ 机器级未完全卸载，请把以下清单反馈给分发人，或在管理员 PowerShell 手动执行：'
            Write-Host '  powershell -ExecutionPolicy Bypass -File build\install-elevated.ps1 -Uninstall'
            if ($leftTask) { Write-Host "  - 计划任务仍存在：$taskName" }
            if ($leftDir)  { Write-Host "  - 部署目录仍存在：$machineDir" }
            if ($leftTl)   { Write-Host "  - HKLM 信任位置仍存在：$leftTl" }
        }
    }

    # 正在运行的 Excel 会话内，菜单会保留到该会话关闭（#21 单实例驻留同族）
    $liveXl = Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.WorkingSet64 -gt 1MB }
    Write-Host ''
    # 按事实合成最终退出码（优先级 8 > 10 > 9 > 0）；「卸载完成」只在核验干净时说
    $finalCode = Get-UninstallExitCode $machineFailed $userResidue $userLevelOpenFail
    if ($finalCode -eq 8) {
        Write-Host '结果：卸载未完全成功（见上方残留清单），已打开会话中的菜单保留到其关闭。'
        exit 8
    }
    if ($finalCode -eq 10) {
        Write-Host '结果：机器级已处理，但用户级卸载有残留（清单见上方）。'
        Write-Host '      请关闭所有 Excel（含鼎信诺、WPS 表格）后重新运行卸载；仍失败则按清单手动处理。'
        exit 10
    }
    if ($finalCode -eq 9) {
        Write-Host '结果：机器级已卸载，但用户级加载项列表（OPEN / Add-in Manager）登记未能完全清理（见上方警告）。'
        exit 9
    }
    # 两级都卸干净：收尾清理看门狗日志（存在才删）
    $wdLog = Join-Path $env:TEMP 'AttachLinkWatchdog.log'
    if (Test-Path -LiteralPath $wdLog) { Remove-Item -LiteralPath $wdLog -Force -ErrorAction SilentlyContinue }
    if ($liveXl) {
        Write-Host '注意：检测到 Excel 正在运行——已打开会话中的右键菜单会保留到其关闭，'
        Write-Host '      完全退出所有 Excel 后重新打开即不再出现。'
    } else {
        Write-Host '卸载完成。重新打开 Excel 后右键菜单不再出现。'
    }
    exit 0
}
Write-Host ''
