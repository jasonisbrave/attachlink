# =====================================================================
# AttachLinkSetup.exe 构建脚本（分发方使用，非同事机器步骤）
# 前置：python + pyinstaller（本机已验证 6.21.0 + Python 3.14 + tkinter）
# 产物：release\AttachLink-<版本>\AttachLinkSetup.exe（单文件，内嵌全部载荷）
# 语义：版本号唯一来源 = src\AttachLink.bas 的 ATTACHLINK_VERSION 常量
# =====================================================================
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$root = Split-Path -Parent $PSScriptRoot
$srcBas = Join-Path $root 'src\AttachLink.bas'
if (-not (Test-Path $srcBas)) { throw "未找到 $srcBas" }

# 版本号从 .bas 唯一来源读取
$basText = [IO.File]::ReadAllText($srcBas, [Text.Encoding]::UTF8)
if ($basText -notmatch 'ATTACHLINK_VERSION\s+As\s+String\s*=\s*"([^"]+)"') { throw '未能从 AttachLink.bas 提取 ATTACHLINK_VERSION' }
$ver = $Matches[1]
Write-Host "版本：v$ver"

# 组装 payload（与 exe 内嵌布局一致）
$stage = Join-Path $env:TEMP "AttachLink_exe_payload"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path "$stage\dist", "$stage\scripts", "$stage\build" | Out-Null
$distXlam = Join-Path $root 'dist\AttachLink.xlam'
if (-not (Test-Path $distXlam)) { throw "未找到 $distXlam，请先运行 build.ps1" }
Copy-Item $distXlam "$stage\dist\AttachLink.xlam" -Force
Copy-Item (Join-Path $root 'scripts\setup.ps1')            "$stage\scripts\" -Force
Copy-Item (Join-Path $root 'scripts\watchdog.ps1')         "$stage\scripts\" -Force
Copy-Item (Join-Path $root 'scripts\watchdog_launcher.vbs') "$stage\scripts\" -Force
Copy-Item (Join-Path $root 'scripts\healthcheck.ps1')       "$stage\scripts\" -Force
Copy-Item (Join-Path $root 'build\install.ps1')            "$stage\build\" -Force
Copy-Item (Join-Path $root 'build\install-elevated.ps1')   "$stage\build\" -Force
Copy-Item (Join-Path $root 'packaging\run-elevated.ps1')      "$stage\" -Force
Copy-Item (Join-Path $root 'packaging\elevated-capture.ps1') "$stage\" -Force
[IO.File]::WriteAllText("$stage\VERSION.txt", $ver, [Text.UTF8Encoding]::new($false))
Write-Host "payload 就绪：$stage"

# pyinstaller（CLI 直调，免维护 spec 文件；--add-data 的 Windows 分隔符为分号；spec 定向到专用临时目录，不落到仓库根/CWD）
$work = Join-Path $root 'build\exe-build'
$out  = Join-Path $root 'build\exe-dist'
$specDir = Join-Path $env:TEMP 'AttachLink_exe_spec'
$releaseDir = Join-Path $root "release\AttachLink-$ver"
if (Test-Path $work)  { Remove-Item $work -Recurse -Force }
if (Test-Path $out)   { Remove-Item $out -Recurse -Force }
if (Test-Path $specDir) { Remove-Item $specDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $specDir | Out-Null
$addData = @(
    , @("$stage\VERSION.txt", "payload")
    , @("$stage\dist\AttachLink.xlam", "payload\dist")
    , @("$stage\scripts\setup.ps1", "payload\scripts")
    , @("$stage\scripts\watchdog.ps1", "payload\scripts")
    , @("$stage\scripts\watchdog_launcher.vbs", "payload\scripts")
    , @("$stage\scripts\healthcheck.ps1", "payload\scripts")
    , @("$stage\build\install.ps1", "payload\build")
    , @("$stage\build\install-elevated.ps1", "payload\build")
    , @("$stage\run-elevated.ps1", "payload")
    , @("$stage\elevated-capture.ps1", "payload")
)
$argList = @('-m', 'PyInstaller', '--noconfirm', '--clean', "--workpath", $work, "--distpath", $out,
    '--specpath', $specDir,
    '--onefile', '--noconsole', '--name', 'AttachLinkSetup')
foreach ($ad in $addData) { $argList += @('--add-data', "$($ad[0]);$($ad[1])") }
$argList += (Join-Path $root 'packaging\setup_gui.py')
& python @argList
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller 构建失败' }
# spec 已定向到专用临时目录（$specDir），仓库根与 CWD 不应再有 AttachLinkSetup.spec——残留即告警
foreach ($straySpec in @((Join-Path $root 'AttachLinkSetup.spec'), (Join-Path ([string](Get-Location)) 'AttachLinkSetup.spec'))) {
    if ($straySpec -and (Test-Path $straySpec)) { Write-Warning "发现残留 spec：$straySpec"; Remove-Item $straySpec -Force }
}

# 发布目录：exe + README（最终交付 zip 即此目录打包）
New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
Copy-Item (Join-Path $out 'AttachLinkSetup.exe') "$releaseDir\AttachLinkSetup.exe" -Force
Copy-Item (Join-Path $root 'README.md') "$releaseDir\README.md" -Force
Write-Host "构建完成：$releaseDir\AttachLinkSetup.exe"
Write-Host "自检：AttachLinkSetup.exe /selfcheck（应列出全部 OK）"
