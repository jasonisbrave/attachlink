# =====================================================================
# AttachLink 看门狗（.NET 版）构建脚本 —— 预研基础第 5 步
# ---------------------------------------------------------------------
# 用法: powershell -ExecutionPolicy Bypass -File build\make-watchdog.ps1 [-OutputDir ..\dist]
# 产物: dist\AttachLink.Watchdog.exe（控制台程序；依赖 .NET Framework 4.8，Windows 10 1903 以后与
#       Windows 11 自带；由启动器以无窗口方式启动）
# 编译器: 本机「Visual Studio 生成工具 2026」自带的 Roslyn csc（经 vswhere 定位），直接引用系统
#       .NET Framework 目录里的程序集，不需要安装 SDK 或目标包。
#       /deterministic: 同样的源码产出同样的二进制（部署核对用）。
#       不在用户机器上编译——安装器打包的是这里的产物。
# 源码: src\WatchdogNet\*.cs（UTF-8 with BOM；/codepage:65001）。
# =====================================================================
param(
    [string]$OutputDir = (Join-Path $PSScriptRoot '..\dist')
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$root   = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $root 'src\WatchdogNet'
$sources = @(Get-ChildItem -LiteralPath $srcDir -Filter '*.cs' -File | Sort-Object Name | ForEach-Object { $_.FullName })
if ($sources.Count -eq 0) { throw "找不到源码: $srcDir" }

# ---- 定位 Roslyn csc（Visual Studio 生成工具 / MSBuild 自带）----
$csc = $null
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (Test-Path -LiteralPath $vswhere) {
    $installs = @(& $vswhere -all -products * -requires Microsoft.Component.MSBuild -property installationPath 2>$null | Where-Object { $_ })
    foreach ($inst in $installs) {
        foreach ($rel in 'MSBuild\Current\Bin\Roslyn\csc.exe', 'MSBuild\Current\Bin\amd64\Roslyn\csc.exe') {
            $cand = Join-Path $inst $rel
            if (Test-Path -LiteralPath $cand) { $csc = $cand; break }
        }
        if ($csc) { break }
    }
}
if (-not $csc) { throw '找不到 Roslyn csc（需要 Visual Studio 生成工具或 MSBuild；vswhere 定位不到）。没有它时只能退回系统自带的 csc.exe（只认 C# 5 语法），本脚本不自动降级。' }

$fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
foreach ($n in 'System.dll', 'System.Core.dll') {
    if (-not (Test-Path -LiteralPath (Join-Path $fw $n))) { throw "找不到 .NET Framework 程序集: $fw\$n" }
}

$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$out = Join-Path $OutputDir 'AttachLink.Watchdog.exe'
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }

# /noconfig: 不读编译器目录下的 csc.rsp（它按简单名引用程序集）；mscorlib 由编译器隐式引用
$cscArgs = @('/nologo', '/target:exe', '/platform:anycpu', '/optimize+', '/deterministic+', '/codepage:65001', '/warn:1', '/noconfig',
    "/r:$fw\System.dll", "/r:$fw\System.Core.dll", "/out:$out") + $sources
& $csc $cscArgs
if ($LASTEXITCODE -ne 0) { throw "csc 编译失败（退出码 $LASTEXITCODE）" }

$sha = (Get-FileHash -LiteralPath $out -Algorithm SHA256).Hash.Substring(0, 16)
Write-Host "构建完成: $out（.NET Framework 4.8 控制台程序；确定性构建；SHA-256 前16位 $sha）"
