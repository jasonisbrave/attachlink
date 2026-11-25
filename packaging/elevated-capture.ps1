# elevated-capture.ps1 —— 提权侧执行包装器（由 run-elevated.ps1 经 UAC 以管理员调用）
# 职责：运行目标脚本一次，把全部输出捕获到日志文件，退出码透传；
#       目标脚本抛出终止错误时，把「[异常] 消息 + 出错的文件与行号」追加进日志（否则日志里
#       只有出错前的输出，提权侧失败原因排查不到），退出码保持非 0。
param(
    [Parameter(Mandatory = $true)][string]$Target,
    [string]$TargetArgs = '',
    [Parameter(Mandatory = $true)][string]$Log
)
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$code = 0
try {
    if ($TargetArgs) {
        & $Target $TargetArgs *>&1 | ForEach-Object { "$_" } | Out-File -FilePath $Log -Encoding UTF8
    } else {
        & $Target *>&1 | ForEach-Object { "$_" } | Out-File -FilePath $Log -Encoding UTF8
    }
    $code = $LASTEXITCODE
} catch {
    "[异常] $($_.Exception.Message)" | Out-File -FilePath $Log -Encoding UTF8 -Append
    if ($_.InvocationInfo -and ($_.InvocationInfo.ScriptName -or $_.InvocationInfo.ScriptLineNumber -gt 0)) {
        "[异常] 位置：$($_.InvocationInfo.ScriptName) 行 $($_.InvocationInfo.ScriptLineNumber)" |
            Out-File -FilePath $Log -Encoding UTF8 -Append
    }
    if (-not $code) { $code = 1 }
}
exit $code
