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

    # ---- 自动检测机器级部署是否需要同步（T6 落地：防止升级漏跑第二步）----
    $pfXlam = 'C:\Program Files\AttachLink\AttachLink.xlam'
    if (Test-Path $pfXlam) {
        try {
            $shaProv = [System.Security.Cryptography.SHA256]::Create()
            $fsA = [IO.File]::OpenRead($destXlam)
            try { $hLocal = [BitConverter]::ToString($shaProv.ComputeHash($fsA)) -replace '-', '' } finally { $fsA.Close() }
            $fsB = [IO.File]::OpenRead($pfXlam)
            try { $hPf = [BitConverter]::ToString($shaProv.ComputeHash($fsB)) -replace '-', '' } finally { $fsB.Close(); $shaProv.Dispose() }
            if ($hLocal -ne $hPf) {
                Write-Host ''
                Write-Host '==================================================================='
                Write-Host '  ⚠ 检测到鼎信诺管理员环境部署与用户级版本不一致（需同步更新）'
                Write-Host '  本机如需在鼎信诺中使用，请再执行一次「鼎信诺管理员环境部署」：'
                Write-Host '  AttachLinkSetup.exe → 点「鼎信诺管理员环境部署」按钮（弹 UAC）'
                Write-Host '  或管理员 PowerShell: powershell -ExecutionPolicy Bypass -File build\install-elevated.ps1'
                Write-Host '==================================================================='
            }
        } catch { }
    }
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
    # ---- 清理加载项列表残留（OPEN / OPENn，README E7）----
    # Excel 的 Installed=False 只把 OPENn 标记为禁用（/A 前缀）并不删除条目，
    # 「选项-加载项」里仍挂着 AttachLink。
    # 关键契约：OPEN 系列必须连续（OPEN, OPEN1, OPEN2…），Excel 读到第一个缺口即
    # 停止加载——删除中间条目后必须把剩余条目重排为连续序列，否则排在其后的
    # 第三方加载项不再加载（比视觉残留更糟）。重排只改键名、值原样保留
    # （第三方条目的 /R 等前缀不动）。（源自 PR #2 复核结论 + 补全实现）
    # oracle 二轮原子性要求：删除/写入 -ErrorAction Stop、仅真实删除计数、
    # 重排前快照可回滚、终态核验（连续 + 数量一致），失败 exit 9。
    $openCleanupFailed = $false
    foreach ($v in @('16.0', '15.0', '14.0')) {
        $optKey = "HKCU:\Software\Microsoft\Office\$v\Excel\Options"
        if (-not (Test-Path $optKey)) { continue }
        $props = @((Get-ItemProperty -LiteralPath $optKey).PSObject.Properties |
                   Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] })
        $removed = 0
        foreach ($p in $props) {
            $entry = $p.Value.Trim().Trim('"')
            if ($entry -match '^(/[Aa]|\\[Aa])\s+') { $entry = $entry.Substring(2).Trim().Trim('"') }
            if (($entry -ieq $destXlam) -or ($entry -ieq 'AttachLink.xlam') -or ($entry -like '*\AttachLink\AttachLink.xlam')) {
                try {
                    Remove-ItemProperty -LiteralPath $optKey -Name $p.Name -ErrorAction Stop
                    $removed++
                } catch {
                    Write-Warning ("OPEN 条目 {0} 删除失败：{1}（该条目保留，序列重排将绕过缺口）" -f $p.Name, $_.Exception.Message)
                    $openCleanupFailed = $true
                }
            }
        }
        if ($removed -gt 0) {
            # 快照（重排失败时的回滚基准：现存条目原名原值）
            $rest = @((Get-ItemProperty -LiteralPath $optKey).PSObject.Properties |
                      Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] } |
                      Sort-Object { [int]($_.Name -replace '\D', '') })
            $snapshot = @($rest | ForEach-Object { @{ Name = [string]$_.Name; Value = [string]$_.Value } })
            try {
                for ($j = $snapshot.Count - 1; $j -ge 0; $j--) {
                    Remove-ItemProperty -LiteralPath $optKey -Name $snapshot[$j].Name -ErrorAction Stop
                }
                for ($j = 0; $j -lt $snapshot.Count; $j++) {
                    $newName = 'OPEN'
                    if ($j -gt 0) { $newName = "OPEN$j" }
                    New-ItemProperty -LiteralPath $optKey -Name $newName -Value $snapshot[$j].Value -PropertyType String -Force -ErrorAction Stop | Out-Null
                }
            } catch {
                # 回滚：清残键后按快照原名写回（尽力而为）
                Write-Warning ("OPEN 序列重排异常，按快照回滚：{0}" -f $_.Exception.Message)
                Get-ItemProperty -LiteralPath $optKey -ErrorAction SilentlyContinue |
                    ForEach-Object { $_.PSObject.Properties } |
                    Where-Object { $_.Name -match '^OPEN(\d+)?$' } |
                    ForEach-Object { Remove-ItemProperty -LiteralPath $optKey -Name $_.Name -ErrorAction SilentlyContinue }
                foreach ($s in $snapshot) {
                    New-ItemProperty -LiteralPath $optKey -Name $s.Name -Value $s.Value -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
                }
                $openCleanupFailed = $true
            }
            # 终态核验：序列连续 + 数量与快照一致 + 第三方值逐项一致
            if (-not $openCleanupFailed) {
                $verify = @((Get-ItemProperty -LiteralPath $optKey).PSObject.Properties |
                            Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] } |
                            Sort-Object { [int]($_.Name -replace '\D', '') })
                $contigOk = ($verify.Count -eq $snapshot.Count)
                for ($j = 0; $contigOk -and $j -lt $verify.Count; $j++) {
                    $expect = 'OPEN'
                    if ($j -gt 0) { $expect = "OPEN$j" }
                    if ($verify[$j].Name -ne $expect -or "$($verify[$j].Value)" -ne $snapshot[$j].Value) { $contigOk = $false }
                }
                if ($contigOk) {
                    Write-Host ("已清理 Office {0} 加载项列表 OPEN 残留 {1} 项并重排序列（第三方条目保留，核验一致）。" -f $v, $removed)
                } else {
                    Write-Warning "Office $v OPEN 序列终态核验未通过（数量/连续性/值一致性），请人工核对加载项列表。"
                    $openCleanupFailed = $true
                }
            }
        }
    }
    # 终态残留核验（oracle 三轮）：无论删除/重排是否执行，最终不得残留任何指向 AttachLink 的 OPEN 条目
    foreach ($v in @('16.0', '15.0', '14.0')) {
        $optKey = "HKCU:\Software\Microsoft\Office\$v\Excel\Options"
        if (-not (Test-Path $optKey)) { continue }
        $leftovers = @((Get-ItemProperty -LiteralPath $optKey).PSObject.Properties |
            Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] } |
            Where-Object {
                $e = $_.Value.Trim().Trim('"')
                if ($e -match '^(/[Aa]|\\[Aa])\s+') { $e = $e.Substring(2).Trim().Trim('"') }
                ($e -ieq $destXlam) -or ($e -ieq 'AttachLink.xlam') -or ($e -like '*\AttachLink\AttachLink.xlam')
            })
        if ($leftovers.Count -gt 0) {
            Write-Warning ("Office {0} 仍残留 {1} 条指向 AttachLink 的 OPEN 条目，请人工在 Excel「选项-加载项」移除。" -f $v, $leftovers.Count)
            $openCleanupFailed = $true
        }
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
    if ($openCleanupFailed) {
        Write-Host '卸载完成（文件与注册表其余清理已执行），但加载项列表 OPEN 清理未完全成功（见上方警告）。'
        exit 9
    }
    Write-Host '卸载完成（含加载项列表 OPEN/OPENn 残留清理与序列重排）。'
    Write-Host '若曾在其他路径手动加载过本加载项副本，可在 Excel「选项-加载项-管理: Excel 加载项-转到」中手动移除。'
}
