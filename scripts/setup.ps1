# =====================================================================
# AttachLink 一键安装 / 一键卸载 编排脚本
# 由项目根目录的「一键安装.bat」「一键卸载.bat」调用，也可手动运行：
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
        $xl = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
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
    Write-Host '  【接收/复核归档底稿的同事】到这里就完成了。'
    Write-Host '  【需要在本机以管理员运行的鼎信诺里编制底稿的同事】还需一步：'
    Write-Host '     关闭鼎信诺和 Excel 后，右键以管理员身份运行'
    Write-Host '     build\install-elevated.ps1（管理员 PowerShell 中执行：'
    Write-Host '     powershell -ExecutionPolicy Bypass -File build\install-elevated.ps1）'
    Write-Host '  -------------------------------------------------'
    Write-Host '  详细用法见文件夹中的 README.md'
    Write-Host '==================================================='
}
else {
    Write-Host ''
    Write-Host '>> 正在卸载 AttachLink 加载项（用户级）...'
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $buildDir 'install.ps1') -Uninstall
    if ($LASTEXITCODE -ne 0) { throw '卸载失败，请把上方错误信息反馈给分发人。' }
    Write-Host ''
    Write-Host '卸载完成（用户级：日常 Excel 中的右键菜单/快捷键已移除）。'
    Write-Host '注意：本卸载不含机器级看门狗。若本机曾以管理员运行过'
    Write-Host '      build\install-elevated.ps1（鼎信诺场景），请另行在管理员'
    Write-Host '      PowerShell 运行：powershell -ExecutionPolicy Bypass -File'
    Write-Host '      build\install-elevated.ps1 -Uninstall'
}
Write-Host ''
