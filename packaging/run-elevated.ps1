# run-elevated.ps1 —— 提权桥：以管理员运行目标脚本，日志与退出码回传
# 由 AttachLinkSetup.exe 调用：run-elevated.ps1 <target-script-relative-path> [args]
# 机制：UAC 启动 powershell -File elevated-capture.ps1（参数经引号拼接的单一参数行
#       传递，规避 Start-Process 多元素拼接对 -Command 管道/重定向的静默破坏）
# 输出：UAC 取消返回 1223；成功返回目标退出码，并把目标输出原样转发
param(
    [Parameter(Mandatory = $true)][string]$Target
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$parts = $Target.Split(' ', 2)
$scriptRel = $parts[0]
$scriptArgs = if ($parts.Count -gt 1) { $parts[1] } else { '' }
$scriptPath = Join-Path $PSScriptRoot $scriptRel
if (-not (Test-Path $scriptPath)) { Write-Output "run-elevated: 找不到 $scriptPath"; exit 4 }

$capture = Join-Path $PSScriptRoot 'elevated-capture.ps1'
if (-not (Test-Path $capture)) { Write-Output "run-elevated: 找不到 $capture"; exit 4 }

$log = Join-Path $env:TEMP 'AttachLinkSetup_elevated.log'
Remove-Item $log -Force -ErrorAction SilentlyContinue

# 单一参数行 + 内嵌引号：完全掌控传给提权进程的命令行
$argLine = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Target "{1}" -TargetArgs "{2}" -Log "{3}"' -f $capture, $scriptPath, $scriptArgs, $log
try {
    $p = Start-Process powershell -Verb RunAs -Wait -PassThru -WindowStyle Hidden -ArgumentList $argLine
} catch {
    Write-Output "run-elevated: UAC 被取消或启动失败 - $($_.Exception.Message)"
    exit 1223
}
if (Test-Path $log) {
    Get-Content $log -Encoding UTF8 | ForEach-Object { Write-Output $_ }
    Remove-Item $log -Force -ErrorAction SilentlyContinue
} else {
    Write-Output "run-elevated: 提权进程已运行（退出码 $($p.ExitCode)）但未产生日志"
}
exit $p.ExitCode
