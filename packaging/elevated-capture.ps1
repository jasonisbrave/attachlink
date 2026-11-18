# elevated-capture.ps1 —— 提权侧执行包装器（由 run-elevated.ps1 经 UAC 以管理员调用）
# 职责：运行目标脚本一次，把全部输出捕获到日志文件，退出码透传
param(
    [Parameter(Mandatory = $true)][string]$Target,
    [string]$TargetArgs = '',
    [Parameter(Mandatory = $true)][string]$Log
)
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
if ($TargetArgs) {
    & $Target $TargetArgs *>&1 | ForEach-Object { "$_" } | Out-File -FilePath $Log -Encoding UTF8
} else {
    & $Target *>&1 | ForEach-Object { "$_" } | Out-File -FilePath $Log -Encoding UTF8
}
exit $LASTEXITCODE
