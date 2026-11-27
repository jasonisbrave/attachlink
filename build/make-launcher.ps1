# =====================================================================
# AttachLink 看门狗启动器构建：src\WatchdogLauncher.cs -> dist\AttachLink.WatchdogLauncher.exe
# ---------------------------------------------------------------------
# 用系统自带的 .NET Framework C# 编译器（csc.exe）编译，产物是 GUI 子系统程序（没有控制台窗口）。
# 只在开发机上编译：发布包里带的是预编译产物，不在用户机器上现场编译。
# 用法: powershell -ExecutionPolicy Bypass -File make-launcher.ps1 [-OutputDir ..\dist]
# 产物: <OutputDir>\AttachLink.WatchdogLauncher.exe（由 build.ps1 一并调用；也可以单独运行）
# =====================================================================
param(
    [string]$OutputDir = (Join-Path $PSScriptRoot '..\dist')
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$src = Join-Path $PSScriptRoot '..\src\WatchdogLauncher.cs'
if (-not (Test-Path -LiteralPath $src)) { throw "找不到源码: $src" }

$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$out = Join-Path $OutputDir 'AttachLink.WatchdogLauncher.exe'

$fw = Join-Path $env:WINDIR 'Microsoft.NET'
$csc = @((Join-Path $fw 'Framework64\v4.0.30319\csc.exe'), (Join-Path $fw 'Framework\v4.0.30319\csc.exe')) |
    Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $csc) { throw '找不到系统自带的 C# 编译器（.NET Framework 4.x 的 csc.exe）' }

# /target:winexe：GUI 子系统，无控制台窗口。（系统自带的这个 csc 不支持 /deterministic，两次构建的字节不保证相同；
# 版本一致性以构建时的哈希为准，不要跨构建比较。）
& $csc /nologo /target:winexe /platform:anycpu /optimize+ /warn:0 "/out:$out" "$src"
if ($LASTEXITCODE -ne 0) { throw "csc 编译失败（退出码 $LASTEXITCODE）" }

# 自检：PE 头的 Subsystem 必须是 2（GUI）。编成控制台程序的话，计划任务一启动就会在桌面上弹出窗口
$bytes = [System.IO.File]::ReadAllBytes($out)
if ($bytes.Length -lt 256 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) { throw '产物不是 PE 文件' }
$peOff = [BitConverter]::ToInt32($bytes, 0x3C)
$subsystem = [BitConverter]::ToUInt16($bytes, $peOff + 4 + 20 + 68)
if ($subsystem -ne 2) { throw ("产物的 PE Subsystem={0}（期望 2 = GUI 子系统）" -f $subsystem) }

$sha = [System.Security.Cryptography.SHA256]::Create()
$fs = [System.IO.File]::OpenRead($out)
try { $hash = ([System.BitConverter]::ToString($sha.ComputeHash($fs)) -replace '-', '').Substring(0, 16) } finally { $fs.Close() }

Write-Host ("构建完成: {0}（GUI 子系统，无控制台窗口；SHA-256 前16位 {1}）" -f $out, $hash)
