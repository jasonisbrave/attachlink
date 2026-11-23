# =====================================================================
# AttachLink.xlam 安装 / 卸载脚本
# 安装: powershell -ExecutionPolicy Bypass -File install.ps1
# 卸载: powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall
# 动作: 复制 xlam 到 %APPDATA%\Microsoft\AddIns（Excel 库目录），
#       将该目录加入受信任位置（AllowSubFolders=0，最小信任范围），
#       并注册启用为 Excel 加载项
# 说明: 加载项登记（Options\OPENn）只存文件名、由 Excel 按库目录解析，
#       因此文件必须位于库目录根下，无法放入专用子目录（已实证）。
# =====================================================================
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$addInsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
$distXlam  = Join-Path $PSScriptRoot '..\dist\AttachLink.xlam'
$destXlam  = Join-Path $addInsDir 'AttachLink.xlam'
$legacySub = Join-Path $addInsDir 'AttachLink'          # 早期实验遗留的子目录（存在则清理）

# ---- 探测 Office 主版本 ----
$ver = $null
foreach ($v in @('16.0', '15.0', '14.0')) {
    if (Test-Path "HKCU:\Software\Microsoft\Office\$v\Excel") { $ver = $v; break }
}
if (-not $ver) {
    foreach ($v in @('16.0', '15.0', '14.0')) {
        if (Test-Path "HKLM:\Software\Microsoft\Office\$v\Excel") { $ver = $v; break }
    }
}
if (-not $ver) { $ver = '16.0' }

# 注意：Excel 实际读取的受信任位置在 ...\Excel\Security\Trusted Locations
#（"Trust Center" 只是界面术语，不是注册表键；旧版安装器曾误写到该路径，此处顺带清理）
$tl    = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations\AttachLink"
$tlBad = "HKCU:\Software\Microsoft\Office\$ver\Excel\Trust Center\Trusted Locations\AttachLink"

if (-not $Uninstall) {
    if (-not (Test-Path $distXlam)) { throw "未找到 $distXlam ，请先运行 build.ps1 构建。" }

    New-Item -ItemType Directory -Force -Path $addInsDir | Out-Null
    try {
        Copy-Item $distXlam $destXlam -Force
        Unblock-File $destXlam -ErrorAction SilentlyContinue   # 清除下载标记（MOTW），防止宏被静默拦截
    } catch {
        throw '复制文件失败（文件可能被打开的 Excel 占用），请先关闭 Excel 后重试。'
    }
    Remove-Item $legacySub -Recurse -Force -ErrorAction SilentlyContinue

    # ---- 加入受信任位置（仅 AddIns 库目录本身，AllowSubFolders=0 最小化范围）----
    Remove-Item $tlBad -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -Path $tl -Force | Out-Null
    Set-ItemProperty -Path $tl -Name 'Path' -Value "$addInsDir\"
    Set-ItemProperty -Path $tl -Name 'AllowSubFolders' -Value 0 -Type DWORD
    Set-ItemProperty -Path $tl -Name 'Date' -Value (Get-Date -Format 'MM/dd/yyyy')
    Set-ItemProperty -Path $tl -Name 'Description' -Value 'AttachLink add-in library folder'

    # ---- 清理历史遗留的 COM 引导器（该路线已废弃：自动化实例中 OnConnection
    #      从不触发；装载职责由看门狗承担。此处仅移除旧安装可能留下的痕迹）----
    foreach ($k in @(
        'HKCU:\Software\Microsoft\Office\Excel\Addins\AttachLinkShim.Connect',
        'HKCU:\Software\Classes\AttachLinkShim.Connect',
        'HKCU:\Software\Classes\CLSID\{D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58}'
    )) { Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue }
    Remove-Item (Join-Path $addInsDir 'AttachLinkShim.dll') -Force -ErrorAction SilentlyContinue

    # ---- 注册并启用加载项（注册失败时回滚受信任位置写入）----
    # 版本与哈希输出（支持排障用；升级核对两级是否同版）。
    # 用 .NET 直调而非 Get-FileHash：安装器 exe 的子 PowerShell 环境
    # 可能无法自动加载 Utility 模块（实测 CommandNotFoundException）
    $shaProv = [System.Security.Cryptography.SHA256]::Create()
    $fsHash = [IO.File]::OpenRead($destXlam)
    try { $xlamHash = ([BitConverter]::ToString($shaProv.ComputeHash($fsHash)) -replace '-', '').Substring(0, 16) }
    finally { $fsHash.Close(); $shaProv.Dispose() }
    Write-Host ("已安装用户级加载项：{0}（SHA-256 前16位 {1}）" -f $destXlam, $xlamHash)
    # Excel COM 创建有竞态（前一实例退出中会失败），重试 3 次
    $excel = $null
    foreach ($attempt in 1..3) {
        try {
            $excel = New-Object -ComObject Excel.Application
            break
        } catch {
            if ($attempt -eq 3) {
                throw '未检测到可用的桌面版 Excel（无法创建 Excel.Application），请先安装 Office Excel 再运行安装。'
            }
            Start-Sleep -Seconds 2
        }
    }
    try {
        # AddIns.Add 在没有任何打开工作簿的纯 COM 实例上会报错，先建一个临时工作簿
        $wbTmp = $excel.Workbooks.Add()
        $ai = $excel.AddIns.Add($destXlam, $false)
        $ai.Installed = $true
        if (-not $ai.Installed) { throw '加载项注册后未处于启用状态。' }
        $wbTmp.Close($false)
    } catch {
        Remove-Item $tl -Recurse -Force -ErrorAction SilentlyContinue
        throw "加载项注册失败，已回滚受信任位置写入。错误：$($_.Exception.Message)"
    } finally {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    Write-Host '安装完成: AttachLink 加载项已启用。'
    Write-Host '重新打开 Excel 后，在单元格上点击右键即可看到「插入附件链接（可多选）」。'
}
else {
    try {
        $excel = New-Object -ComObject Excel.Application
        try {
            $excel.AddIns.Item('AttachLink').Installed = $false
        } catch {}
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    } catch {
        Write-Host '未检测到 Excel，跳过加载项注销，直接清理文件。'
    }
    # ---- 清理加载项列表残留（OPEN / OPENn）----
    # Excel 的 Installed=False 只把 OPENn 标记为 /A 禁用，并不删除条目，
    # 导致「选项-加载项」里仍挂着 AttachLink（右键菜单的加载入口视觉残留）。
    # 按路径精确删除，第三方加载项一律保留；不依赖 Excel 是否可用。
    foreach ($v in @('16.0', '15.0', '14.0')) {
        $optKey = "HKCU:\Software\Microsoft\Office\$v\Excel\Options"
        if (-not (Test-Path $optKey)) { continue }
        $props = @((Get-ItemProperty -LiteralPath $optKey).PSObject.Properties |
                   Where-Object { $_.Name -like 'OPEN*' -and $_.Value -is [string] })
        $removed = 0
        foreach ($p in $props) {
            $entry = $p.Value.Trim().Trim('"')
            if ($entry -match '^(/[Aa]|\\[Aa])\s+') { $entry = $entry.Substring(2).Trim().Trim('"') }
            if (($entry -ieq $destXlam) -or ($entry -ieq 'AttachLink.xlam')) {
                Remove-ItemProperty -LiteralPath $optKey -Name $p.Name -ErrorAction SilentlyContinue
                $removed++
            }
        }
        if ($removed -gt 0) { Write-Host ("已清理 Office {0} 加载项列表残留条目 {1} 项。" -f $v, $removed) }
    }
    if (Test-Path $destXlam) {
        # Excel COM Quit 为异步：文件锁释放有延迟，重试删除（最多约 8 秒）
        $n = 0
        while ($n -lt 10) {
            try { Remove-Item $destXlam -Force -ErrorAction Stop; break }
            catch { Start-Sleep -Milliseconds 800; $n++ }
        }
    }
    Remove-Item $legacySub -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $tl -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $tlBad -Recurse -Force -ErrorAction SilentlyContinue
    # 清理 COM 引导器
    Remove-Item 'HKCU:\Software\Microsoft\Office\Excel\Addins\AttachLinkShim.Connect' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item 'HKCU:\Software\Classes\AttachLinkShim.Connect' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item 'HKCU:\Software\Classes\CLSID\{D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58}' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path $addInsDir 'AttachLinkShim.dll') -Force -ErrorAction SilentlyContinue
    Write-Host '卸载完成（含右键菜单入口：加载项列表残留已清理）。'
    Write-Host '若曾在别的路径手动加载过本加载项副本，请在 Excel「选项-加载项-管理: Excel 加载项-转到」中手动移除。'
}
