# =====================================================================
# AttachLink.xlam 安装 / 卸载脚本
# 安装: powershell -ExecutionPolicy Bypass -File install.ps1
# 卸载: powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall
# 动作: 复制 xlam 到 %APPDATA%\Microsoft\AddIns（Excel 库目录），
#       将该目录加入受信任位置（AllowSubFolders=0，最小信任范围），
#       并注册启用为 Excel 加载项
# 说明: 加载项登记（Options\OPENn）只存文件名、由 Excel 按库目录解析，
#       因此文件必须位于库目录根下，无法放入专用子目录（已实证）。
# 加载项列表登记（2026-11-24）: Excel 的加载项列表按【文件名】识别条目，AddIns.Add(路径) 遇到同名条目时
#       直接返回已有条目、忽略传入的路径。机器级看门狗登记过 Program Files 副本的机器上，列表里的
#       「AttachLink.xlam」指向那份副本（登记在 Options\OPENn 或 Add-in Manager）：文件已随机器级卸载
#       删除时，用户级安装会绑定到这条悬空登记，报「加载项注册后未处于启用状态」；文件还在时则
#       「安装成功」但登记的不是库目录副本。现安装前先清掉指向其他路径的登记，卸载时一并清理。
#       详见下方「加载项列表登记的识别与清理」。
# 宿主识别（v1.4.3）: Excel 与 WPS 共存的机器上，若 WPS 开启了「兼容第三方系统和
#       软件」，Excel.Application 这个 ProgID 会被解析到 WPS 表格——旧版会把加载项
#       「注册」进一个隐藏的 WPS 实例，真正的 Excel 侧什么都没装。现按 Application.Path
#       判别宿主（WPS 的 Name/Version 均伪装成 Excel），命中 WPS 即改用 Excel 固定
#       CLSID 直接激活。诊断：powershell -ExecutionPolicy Bypass -File install.ps1 -ProbeHost
# =====================================================================
param(
    [switch]$Uninstall,
    # 诊断模式：只识别「Excel.Application」实际落到哪个宿主，不安装/卸载。退出码 0 = Microsoft Excel
    [switch]$ProbeHost,
    # 首选激活用的 ProgID（默认 Excel.Application；诊断时可换成 ket.Application 以演练「被 WPS 接管」分支）
    [string]$ExcelProgId = 'Excel.Application'
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# ---- Excel 宿主识别（Excel 与 WPS 共存环境）----
$excelClsid = '00024500-0000-0000-C000-000000000046'   # Excel.Application 的 CLSID：Excel 自身注册，不随 ProgID 改写而变
function Test-IsNotRealExcel($app) {
    # 判定「不能当作 Microsoft Excel 使用」：Application.Path 含 kingsoft / wps（WPS 表格的 Name、Version 都
    # 伪装成 Excel，Path 才如实），或 Path 读不到 / 为空。读不到时同样返回 $true——无法确认是 Excel 就不往里装
    # （fail-closed）：调用方会丢弃该实例、改走 Excel 自身 CLSID 激活，仍无法确认则明确报错。
    $p = ''
    try { $p = [string]$app.Path } catch { $p = '' }
    return [bool]((-not $p) -or ($p -match '(?i)kingsoft|wps'))
}function Release-Host($app) {
    try { $app.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
}
function New-RealExcel {
    # 返回真正的 Microsoft Excel Application 对象（调用方负责 Quit + 释放）
    $app = $null
    # Excel COM 创建有竞态（前一实例退出中会失败），重试 3 次
    foreach ($attempt in 1..3) {
        try { $app = New-Object -ComObject $ExcelProgId; break }
        catch {
            if ($attempt -eq 3) { throw '未检测到可用的桌面版 Excel（无法创建 Excel.Application），请先安装 Office Excel 再运行安装。' }
            Start-Sleep -Seconds 2
        }
    }
    if (-not (Test-IsNotRealExcel $app)) { return $app }

    # 「Excel.Application」被解析到 WPS（或宿主身份无法确认）：丢弃该实例，改经 Excel 自身 CLSID 激活
    $wpsPath = ''
    try { $wpsPath = [string]$app.Path } catch {}
    Release-Host $app
    $app = $null
    Write-Host ''
    Write-Host "  注意：「Excel.Application」当前解析到 WPS 表格（$wpsPath），已改用 Excel 自身 CLSID 直接激活。"
    Write-Host '  根因通常是 WPS「配置工具 → 高级 → 兼容设置」勾选了「WPS Office 兼容第三方系统和软件」；'
    Write-Host '  该选项会让所有依赖 Excel COM 的软件（含鼎信诺）拿到 WPS，建议关闭。'
    try {
        $t = [type]::GetTypeFromCLSID([guid]$excelClsid)
        if (-not $t) { throw 'Excel 的 COM 类未注册' }
        $app = [Activator]::CreateInstance($t)
    } catch {
        throw ('未检测到可用的桌面版 Excel：「Excel.Application」未解析到 Microsoft Excel，且按 CLSID 也无法启动它。' +
            "请先安装 Office Excel，或在 WPS 中关闭「兼容第三方系统和软件」后重试。（$($_.Exception.Message)）")
    }
    if (Test-IsNotRealExcel $app) {
        Release-Host $app
        throw '未检测到可用的桌面版 Excel：按 CLSID 激活得到的仍不是 Microsoft Excel，请确认本机已安装 Office Excel。'
    }
    return $app
}

if ($ProbeHost) {
    # 诊断模式：不安装/卸载，不写注册表，不复制文件
    $app = $null
    try {
        $app = New-RealExcel
        Write-Host "宿主识别: Microsoft Excel  ProgID=$ExcelProgId  路径=$($app.Path)  版本=$($app.Version)"
        exit 0
    } catch {
        Write-Host "宿主识别失败: $($_.Exception.Message)"
        exit 1
    } finally {
        if ($app) { Release-Host $app }
    }
}

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

# 加载项列表登记（Options\OPENn、Add-in Manager）按 Office 主版本分开存放，逐个版本检查
$officeVersions = @('16.0', '15.0', '14.0')

# ---- 加载项列表登记的识别与清理（安装 / 卸载共用）----
# 背景（2026-11-24 实测）：Excel 的加载项列表按【文件名】识别条目——AddIns.Add(路径) 遇到同名条目时直接
# 返回已有条目、忽略传入的路径。同名条目的持久化登记有两处，而且 Excel 都是到退出时才落盘：
#   Options\OPENn     已启用的条目；库目录副本只存文件名，其他路径存带引号的完整路径，禁用的带 /A 前缀
#   Add-in Manager    显式添加过、未启用的非库目录副本，值名即完整路径（已启用的条目被置 Installed=False
#                     后，退出时就从 OPENn 挪到这里）
# （历史根因：v1.6.0 起看门狗只装载、不再登记，下面这条链路不再发生；但清理逻辑对 v1.6.0 之前留下的存量登记仍然必要）
# 机器级看门狗把 Program Files 副本 AddIns.Add 进列表并置 Installed 的会话退出后，OPENn 里就多了一条指向
# 那份副本的登记；卸载的第一步（Installed=False）让它挪进 Add-in Manager，文件随后被机器级卸载删除——
# 成了悬空登记。之后用户级安装的 AddIns.Add 就绑定到它、Installed 置位静默失败（「卸载后重装」失败的
# 根因）；文件还在时更隐蔽——安装「成功」，登记的却不是库目录副本。
function ConvertFrom-AddInEntry([string]$raw) {
    # OPENn 值 → 加载项条目：去首尾空白与引号，去掉 /A（禁用）前缀
    $entry = $raw.Trim().Trim('"')
    if ($entry -match '^(/[Aa]|\\[Aa])\s+') { $entry = $entry.Substring(2).Trim().Trim('"') }
    return $entry
}
function Test-IsAttachLinkEntry([string]$entry) {
    # 是否 AttachLink 的登记——按文件名判断（与 Excel 识别加载项列表条目的规则一致）：
    # 库目录副本（文件名或完整路径）与任何其他路径的同名副本都算
    return [bool]($entry -match '(?i)(^|[\\/])AttachLink\.xlam$')
}
function Test-IsLibraryEntry([string]$entry) {
    # 是否库目录副本自身的登记：Excel 按库目录解析的文件名，或登记成了库目录的完整路径
    return [bool](($entry -ieq $destXlam) -or ($entry -ieq 'AttachLink.xlam'))
}
function Test-SamePath($a, $b) {
    try { return ([IO.Path]::GetFullPath($a) -ieq [IO.Path]::GetFullPath($b)) } catch { return $false }
}
function Remove-OpenEntries([string]$OptKey, [scriptblock]$ShouldRemove) {
    # 从 Options 的 OPEN 序列里删除满足 $ShouldRemove 的条目（入参为归一化后的条目），剩余条目重排为连续序列。
    # 关键契约：OPEN 系列必须连续（OPEN, OPEN1, OPEN2…），Excel 读到第一个缺口即停止加载——删除中间条目后
    # 必须把剩余条目重排为连续序列，否则排在其后的第三方加载项不再加载（比视觉残留更糟）。重排只改键名、
    # 值原样保留（第三方条目的 /R 等前缀不动）。（源自 PR #2 复核结论 + 补全实现）
    # oracle 二轮原子性要求：删除/写入 -ErrorAction Stop、仅真实删除计数、重排前快照可回滚、
    # 终态核验（连续 + 数量一致）。返回 @{ Removed = 真实删除数; Failed = 删除 / 重排 / 终态核验任一失败 }
    $removed = 0; $failed = $false
    $props = @((Get-ItemProperty -LiteralPath $OptKey).PSObject.Properties |
               Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] })
    foreach ($p in $props) {
        $entry = ConvertFrom-AddInEntry $p.Value
        if (-not (& $ShouldRemove $entry)) { continue }
        try {
            Remove-ItemProperty -LiteralPath $OptKey -Name $p.Name -ErrorAction Stop
            $removed++
        } catch {
            Write-Warning ("OPEN 条目 {0} 删除失败（{1}）：{2}（该条目保留，序列重排将绕过缺口）" -f $p.Name, $OptKey, $_.Exception.Message)
            $failed = $true
        }
    }
    if ($removed -gt 0) {
        # 快照（重排失败时的回滚基准：现存条目原名原值）
        $rest = @((Get-ItemProperty -LiteralPath $OptKey).PSObject.Properties |
                  Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] } |
                  Sort-Object { [int]($_.Name -replace '\D', '') })
        $snapshot = @($rest | ForEach-Object { @{ Name = [string]$_.Name; Value = [string]$_.Value } })
        try {
            for ($j = $snapshot.Count - 1; $j -ge 0; $j--) {
                Remove-ItemProperty -LiteralPath $OptKey -Name $snapshot[$j].Name -ErrorAction Stop
            }
            for ($j = 0; $j -lt $snapshot.Count; $j++) {
                $newName = 'OPEN'
                if ($j -gt 0) { $newName = "OPEN$j" }
                New-ItemProperty -LiteralPath $OptKey -Name $newName -Value $snapshot[$j].Value -PropertyType String -Force -ErrorAction Stop | Out-Null
            }
        } catch {
            # 回滚：清残键后按快照原名写回（尽力而为）
            Write-Warning ("OPEN 序列重排异常，按快照回滚（{0}）：{1}" -f $OptKey, $_.Exception.Message)
            Get-ItemProperty -LiteralPath $OptKey -ErrorAction SilentlyContinue |
                ForEach-Object { $_.PSObject.Properties } |
                Where-Object { $_.Name -match '^OPEN(\d+)?$' } |
                ForEach-Object { Remove-ItemProperty -LiteralPath $OptKey -Name $_.Name -ErrorAction SilentlyContinue }
            foreach ($s in $snapshot) {
                New-ItemProperty -LiteralPath $OptKey -Name $s.Name -Value $s.Value -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
            }
            $failed = $true
        }
        # 终态核验：序列连续 + 数量与快照一致 + 第三方值逐项一致
        if (-not $failed) {
            $verify = @((Get-ItemProperty -LiteralPath $OptKey).PSObject.Properties |
                        Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] } |
                        Sort-Object { [int]($_.Name -replace '\D', '') })
            $contigOk = ($verify.Count -eq $snapshot.Count)
            for ($j = 0; $contigOk -and $j -lt $verify.Count; $j++) {
                $expect = 'OPEN'
                if ($j -gt 0) { $expect = "OPEN$j" }
                if ($verify[$j].Name -ne $expect -or "$($verify[$j].Value)" -ne $snapshot[$j].Value) { $contigOk = $false }
            }
            if (-not $contigOk) {
                Write-Warning ("OPEN 序列终态核验未通过（数量/连续性/值一致性），请人工核对加载项列表：{0}" -f $OptKey)
                $failed = $true
            }
        }
    }
    return @{ Removed = $removed; Failed = $failed }
}
function Remove-AddInManagerEntries([string]$AmKey, [scriptblock]$ShouldRemove) {
    # 从 Add-in Manager 键里删除值名满足 $ShouldRemove 的登记。返回 @{ Gone = 实际删除的值名; Failed = 有删除失败 }。
    # 该键的值名即加载项完整路径（REG_SZ，数据为空串）；键不存在时什么都不做。
    $gone = @(); $failed = $false
    if (-not (Test-Path -LiteralPath $AmKey)) { return @{ Gone = $gone; Failed = $failed } }
    foreach ($name in @((Get-Item -LiteralPath $AmKey).GetValueNames())) {
        if (-not (& $ShouldRemove $name)) { continue }
        try {
            Remove-ItemProperty -LiteralPath $AmKey -Name $name -ErrorAction Stop
            $gone += $name
        } catch {
            Write-Warning ("Add-in Manager 登记 {0} 删除失败（{1}）：{2}" -f $name, $AmKey, $_.Exception.Message)
            $failed = $true
        }
    }
    return @{ Gone = $gone; Failed = $failed }
}
function Clear-AddInListRegistrations([ValidateSet('All', 'OtherPaths')][string]$Mode = 'All', [string]$HiveRoot = 'HKCU:\Software\Microsoft\Office') {
    # 清理 Excel 加载项列表里的 AttachLink 登记，逐个 Office 主版本处理两处持久化位置：
    #   Add-in Manager（值名即完整路径）——直接删除；
    #   Options 的 OPEN 序列——删除后重排为连续序列（见 Remove-OpenEntries）。
    # 最后对两处各做一遍独立的终态核验：无论删除 / 重排是否执行，终态都不得残留目标条目。
    #   Mode=All         卸载用：所有 AttachLink 登记
    #   Mode=OtherPaths  安装前用：只清指向其他路径的副本，库目录副本自身的登记保留
    # 返回 $true = 全部清理干净；$false = 有删除失败、重排失败或核验时仍有残留——调用方必须据此报失败，
    # 不能带着残留继续（安装时残留的旧登记会让 Excel 绑定到错误副本）。
    $target = 'AttachLink 登记'
    $ShouldRemove = { param($e) Test-IsAttachLinkEntry $e }
    if ($Mode -eq 'OtherPaths') {
        $target = '指向其他路径的 AttachLink 登记'
        $ShouldRemove = { param($e) (Test-IsAttachLinkEntry $e) -and -not (Test-IsLibraryEntry $e) }
    }
    $failed = $false
    foreach ($v in $officeVersions) {
        $amKey = Join-Path $HiveRoot "$v\Excel\Add-in Manager"
        if (Test-Path -LiteralPath $amKey) {
            $am = Remove-AddInManagerEntries $amKey $ShouldRemove
            if ($am.Gone.Count -gt 0) { Write-Host ("已清理 Office {0} Add-in Manager：{1} {2} 项（{3}）" -f $v, $target, $am.Gone.Count, ($am.Gone -join '、')) }
            if ($am.Failed) { $failed = $true }
        }
        $optKey = Join-Path $HiveRoot "$v\Excel\Options"
        if (Test-Path $optKey) {
            $r = Remove-OpenEntries $optKey $ShouldRemove
            if ($r.Failed) { $failed = $true }
            elseif ($r.Removed -gt 0) { Write-Host ("已清理 Office {0} OPEN 序列：{1} {2} 项，已重排序列（第三方条目保留，核验一致）。" -f $v, $target, $r.Removed) }
        }
    }
    # 终态残留核验（oracle 三轮）：删除失败的条目、没被上面的循环覆盖到的版本，都在这里被看见
    foreach ($v in $officeVersions) {
        $optKey = Join-Path $HiveRoot "$v\Excel\Options"
        if (Test-Path $optKey) {
            $leftovers = @((Get-ItemProperty -LiteralPath $optKey).PSObject.Properties |
                Where-Object { $_.Name -match '^OPEN(\d+)?$' -and $_.Value -is [string] } |
                Where-Object { [bool](& $ShouldRemove (ConvertFrom-AddInEntry $_.Value)) })
            if ($leftovers.Count -gt 0) {
                Write-Warning ("Office {0} OPEN 序列仍有 {1} 条{2}未能清理，请人工在 Excel「选项-加载项」移除。" -f $v, $leftovers.Count, $target)
                $failed = $true
            }
        }
        $amKey = Join-Path $HiveRoot "$v\Excel\Add-in Manager"
        if (Test-Path -LiteralPath $amKey) {
            $leftAm = @((Get-Item -LiteralPath $amKey).GetValueNames() | Where-Object { [bool](& $ShouldRemove $_) })
            if ($leftAm.Count -gt 0) {
                Write-Warning ("Office {0} Add-in Manager 仍有 {1} 条{2}未能清理，请人工删除：{3}" -f $v, $leftAm.Count, $target, ($leftAm -join '、'))
                $failed = $true
            }
        }
    }
    return (-not $failed)
}
function Release-Com($obj) {
    # 立即释放 RCW 持有的全部 COM 引用（循环到引用计数归零）。宿主要等外部引用全部释放才真正退出——
    # 脚本若把集合 / 条目等中间对象留到进程结束再放，实例会拖成「关了还在运行」的残影。
    if ($null -eq $obj) { return }
    try { while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) -gt 0) { } } catch {}
}
function Get-ExcelProcessIds {
    # 当前全部 EXCEL 进程号。用 .NET 直调：没有 Excel 在运行时返回空数组（Get-Process -Name 找不到进程时会报
    # 一条非终止错误，加上 -ErrorAction Stop 就把「本来就没有 Excel」变成了失败）；枚举本身失败则抛出，
    # 由调用方按「无法确认」处理。
    return @([System.Diagnostics.Process]::GetProcessesByName('EXCEL') | ForEach-Object { $_.Id })
}
function Test-ProcessAlive([int]$Id) {
    # 进程是否还在。只把「该进程号已不存在」判为已退出；其余查询错误原样抛出，由调用方按「无法确认」处理——
    # 吞掉它们会把「查不到」误当成「已退出」（Get-Process -Id … -ErrorAction SilentlyContinue 就会这样）。
    try {
        $null = [System.Diagnostics.Process]::GetProcessById($Id)
        return $true
    } catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        if ($ex -is [ArgumentException]) { return $false }
        throw
    }
}
function Test-ProcessLive([int]$Id) {
    # 进程是否「存活」：存在且不是墓碑残影（WorkingSet64 趋零的已退出残影不算存活）。
    # 查询错误原样抛出——把「查不到」当成「已退出」会误判 ExitConfirmed。
    try {
        $p = [System.Diagnostics.Process]::GetProcessById($Id)
        return ($p.WorkingSet64 -gt 1MB)
    } catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        if ($ex -is [ArgumentException]) { return $false }
        throw
    }
}
function Get-BlockingExcelProcessIds {
    # 当前存活的 EXCEL 进程号（WorkingSet64 > 1MB，过滤已退出但被句柄持有的墓碑残影）。
    # 卸载预检（setup.ps1）与卸载收尾核验共用。
    return @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $_.WorkingSet64 -gt 1MB } | ForEach-Object { $_.Id })
}
function Test-AddInFileLocked([string]$Path) {
    # 加载项文件是否被占用：以「拒绝一切共享」的方式尝试打开——共享冲突或其它 IO 异常都算占用；
    # 文件不存在则不算占用（没有可删的东西）。卸载预检用：被占用时卸载删不掉文件，且占用方退出时
    # 可能把内存里的加载项登记写回注册表，造成「卸载完成」的假象。
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $fs.Close()
        return $false
    } catch { return $true }
}
function Wait-HostExit([int[]]$ProcIds, [int]$TimeoutSeconds = 15, [int]$PollMilliseconds = 200, [scriptblock]$IsAlive = { param($id) Test-ProcessAlive $id }) {
    # 等这些实例真正退出：Excel 到退出时才把加载项列表（OPENn / Add-in Manager）落盘，早于它清理注册表会被
    # 它的落盘覆盖回去。
    # 返回 $true = 全部确认已退出；$false = 没有可等的对象、超时或查询出错。调用方拿到 $false 必须跳过
    # 加载项列表的注册表清理（见卸载分支），不能当作已退出——核验只能看到当时的注册表，看不到 Excel
    # 内存里还没落盘的登记。
    if ($null -eq $ProcIds -or @($ProcIds).Count -eq 0) { return $false }
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ($true) {
        try {
            $alive = @($ProcIds | Where-Object { [bool](& $IsAlive $_) })
        } catch {
            Write-Warning ("无法确认 Excel 是否已退出：{0}" -f $_.Exception.Message)
            return $false
        }
        if ($alive.Count -eq 0) { return $true }
        if ((Get-Date) -ge $deadline) {
            Write-Warning ("Excel（pid {0}）{1} 秒内没有退出，加载项列表可能尚未落盘。" -f ($alive -join '、'), $TimeoutSeconds)
            return $false
        }
        Start-Sleep -Milliseconds $PollMilliseconds
    }
}
function Get-NewExcelProcessIds([int[]]$Before, [scriptblock]$GetExcelIds = { Get-ExcelProcessIds }) {
    # 相对创建实例前的快照，新出现的 EXCEL 进程号（正常只有刚拉起的那个；New-RealExcel 丢弃的 WPS 实例是
    # et.exe，不在其列）。期间恰好有别的 Excel 启动时一并返回，由调用方全部等完——宁可多等也不认错。
    # 返回 $null = 没有快照或枚举失败（做不了差集，调用方按「无法确认」处理）；空数组 = 确实没有新进程。
    # 用进程差集而不是 Application.Hwnd + Add-Type：用户级安装 / 卸载不该依赖 C# 编译器。
    if ($null -eq $Before) { return $null }
    try {
        $ids = @(& $GetExcelIds)
    } catch {
        Write-Warning ("无法枚举 Excel 进程：{0}" -f $_.Exception.Message)
        return $null
    }
    return ,@($ids | Where-Object { $Before -notcontains $_ })
}
function Invoke-AddInUnregister([scriptblock]$CreateExcel = { New-RealExcel }, [scriptblock]$GetExcelIds = { Get-ExcelProcessIds }, [scriptblock]$IsAlive = { param($id) Test-ProcessAlive $id }, [scriptblock]$IsLive = { param($id) Test-ProcessLive $id }, [int]$TimeoutSeconds = 15, [int]$PollMilliseconds = 200) {
    # 卸载第 1 步：让一个 Excel 实例把 AttachLink 置为 Installed=False，然后确认相关实例都已退出。
    # 返回 @{ Unregistered = 是否真的执行了注销; ExitConfirmed = 是否确认实例已退出（可以接着清注册表）; Error = 失败原因 }
    #
    # ExitConfirmed 只有两种情况为真：
    #   · 取得实例并完成释放（注销本身可能失败：只要相关实例都已退出，接着清注册表就是安全的），
    #     且创建期间新出现的 EXCEL 进程全部等到了退出；
    #   · 没取得实例，但创建期间确实没有新的 EXCEL 进程（例如本机没有 Excel，或 Excel.Application 落到了 WPS）。
    # 其余一律不确认：没取得实例却出现了新进程且没等到它退出；没有创建前的快照或进程枚举失败（做不了差集）。
    # New-RealExcel 失败不等于没有拉起 Excel——它可能在创建实例之后才失败（宿主身份无法确认）并对该实例
    # 调过 Quit，而 Quit 是异步的：那个实例退出时同样会把内存里的加载项列表落盘。
    # COM 对象用完即释放（循环到引用计数归零）：引用不放则实例要等本脚本进程结束才退，成了「关了 Excel
    # 仍在运行」的残影。
    $before = $null
    try { $before = @(& $GetExcelIds) } catch { Write-Warning ("无法枚举 Excel 进程：{0}" -f $_.Exception.Message) }
    $excel = $null; $errorText = ''; $unregistered = $false; $confirmed = $false
    try { $excel = & $CreateExcel } catch { $errorText = $_.Exception.Message }
    if ($excel) {
        $own = Get-NewExcelProcessIds $before $GetExcelIds
        $addins = $null; $ai = $null
        try {
            $addins = $excel.AddIns
            $ai = $addins.Item('AttachLink')
            $ai.Installed = $false
            $unregistered = $true
        } catch { $errorText = $_.Exception.Message } finally {
            Release-Com $ai; Release-Com $addins
        }
        try { $excel.Quit() } catch {}
        Release-Com $excel
        $excel = $null
        $confirmed = Wait-HostExit $own -TimeoutSeconds $TimeoutSeconds -PollMilliseconds $PollMilliseconds -IsAlive $IsAlive
    } else {
        $new = Get-NewExcelProcessIds $before $GetExcelIds
        if ($null -eq $new) {
            $confirmed = $false
        } elseif ($new.Count -eq 0) {
            $confirmed = $true
        } else {
            Write-Warning ("未取得可用的 Excel 实例，但创建期间出现了 Excel 进程（pid {0}），等它退出。" -f ($new -join '、'))
            $confirmed = Wait-HostExit $new -TimeoutSeconds $TimeoutSeconds -PollMilliseconds $PollMilliseconds -IsAlive $IsAlive
        }
    }
    # council C2：原有确认只等「创建期间新出现」的 PID——卸载开始前就已存在的 Excel 若仍在运行，
    # 它退出时会把内存里的加载项登记写回注册表，抢在其后清理会被覆盖回去（「卸载干净」假象）。
    # 因此 $before 里任何 PID 仍然存活（或存活状态无法确认）都不能确认退出。
    if ($confirmed -and $null -ne $before -and @($before).Count -gt 0) {
        try {
            $stale = @($before | Where-Object { [bool](& $IsLive $_) })
        } catch {
            Write-Warning ("无法确认卸载开始前已存在的 Excel 是否仍在运行：{0}" -f $_.Exception.Message)
            $stale = $null
        }
        if ($null -eq $stale) {
            $confirmed = $false
        } elseif ($stale.Count -gt 0) {
            $confirmed = $false
            $msg = ('卸载开始前就已打开的 Excel（pid ' + ($stale -join '、') + '）仍在运行，请关闭后重试')
            Write-Warning $msg
            if ($errorText) { $errorText += '；' + $msg } else { $errorText = $msg }
        }
    }
    return @{ Unregistered = [bool]$unregistered; ExitConfirmed = [bool]$confirmed; Error = $errorText }
}
function Test-DisabledItemsBlob([byte[]]$bytes) {
    # Excel「禁用的项目」（Resiliency\DisabledItems）的一个二进制值里是否列有 AttachLink.xlam。
    # 值内含条目路径的 UTF-16 文本，按完整文件名匹配：前后只能是路径分隔符 / 字符串结尾（NUL）——
    # AttachLinkShim.dll、MyAttachLink.xlam、AttachLink.xlam.bak 都不算。
    if ($null -eq $bytes -or $bytes.Length -eq 0) { return $false }
    try { return [bool]([Text.Encoding]::Unicode.GetString($bytes) -match '(?i)(^|[\\/])AttachLink\.xlam(?=[\x00]|$)') } catch { return $false }
}
function Get-DisabledItemsHint {
    # Excel「禁用的项目」里是否列有 AttachLink.xlam。列在其中时 Installed 置位同样静默失败——
    # 给失败信息补一条可操作提示。返回提示文本；没有则为空串。
    foreach ($v in $officeVersions) {
        $k = "HKCU:\Software\Microsoft\Office\$v\Excel\Resiliency\DisabledItems"
        if (-not (Test-Path -LiteralPath $k)) { continue }
        try {
            $key = Get-Item -LiteralPath $k
            foreach ($n in @($key.GetValueNames())) {
                if (Test-DisabledItemsBlob ($key.GetValue($n))) {
                    return ('Excel 已把 AttachLink 列入「禁用的项目」（Office ' + $v + '）。' +
                        '请在 Excel「文件 → 选项 → 加载项 → 管理：禁用的项目 → 转到」里选中 AttachLink 点「启用」后重试。')
                }
            }
        } catch { }
    }
    return ''
}
function Get-UninstallResidue([bool]$ListCleaned, [string]$XlamPath = $destXlam, [string]$LegacyDir = $legacySub,
    [string]$HiveRoot = 'HKCU:\Software\Microsoft\Office', [string[]]$OfficeVers = $officeVersions,
    [string[]]$ComShimKeys = @(
        'HKCU:\Software\Microsoft\Office\Excel\Addins\AttachLinkShim.Connect',
        'HKCU:\Software\Microsoft\Office\Excel\AddinsData\AttachLinkShim.Connect',
        'HKCU:\Software\Classes\AttachLinkShim.Connect',
        'HKCU:\Software\Classes\CLSID\{D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58}'
    ),
    [scriptblock]$GetLiveExcel = { Get-BlockingExcelProcessIds }) {
    # 卸载收尾核验（council C1/C4）：删除动作全部做完之后按事实收集残留清单，每条形如
    # 「名称：路径（原因 / 处置）」。返回空数组 = 全部干净；有任何一行 = 卸载未完全成功（exit 10）。
    # 全部判定可注入（临时文件 / 临时注册表键 / 假 scriptblock），便于单元测试。
    $residue = @()
    if (Test-Path -LiteralPath $XlamPath) {
        $residue += "加载项副本：$XlamPath（文件未能删除，可能被 Excel / WPS 表格或其它程序占用；处置：关闭全部 Excel 后重新运行卸载）"
    }
    if (Test-Path -LiteralPath $LegacyDir) {
        $residue += "旧子目录：$LegacyDir（目录未能删除；处置：关闭全部 Excel 后重新运行卸载，仍失败请手动删除该目录）"
    }
    # 受信任位置：卸载的清理范围覆盖安装分支可能写过的所有键（16.0 / 15.0 / 14.0 × Security 与 Trust Center 两处）
    foreach ($v in @($OfficeVers)) {
        foreach ($sub in @('Security\Trusted Locations\AttachLink', 'Trust Center\Trusted Locations\AttachLink')) {
            $k = Join-Path $HiveRoot "$v\Excel\$sub"
            if (Test-Path -LiteralPath $k) {
                $residue += "受信任位置登记：$k（注册表键未能删除；处置：请手动删除该键）"
            }
        }
    }
    # 旧 COM 引导器（路线已废弃）：现有清理清单 + AddinsData 缓存键（本机实存 LoadCount=10）+ AddIns 目录下的 DLL
    foreach ($k in @($ComShimKeys)) {
        if (Test-Path -LiteralPath $k) {
            $residue += "旧 COM 引导器登记：$k（注册表键未能删除；处置：请手动删除该键）"
        }
    }
    $shimDll = Join-Path (Split-Path -Parent $XlamPath) 'AttachLinkShim.dll'
    if (Test-Path -LiteralPath $shimDll) {
        $residue += "旧 COM 引导器文件：$shimDll（文件未能删除；处置：请手动删除该文件）"
    }
    if (-not $ListCleaned) {
        $residue += '加载项列表登记：Options\OPENn 或 Add-in Manager 仍有 AttachLink 登记未能清理（见上方警告；处置：请在 Excel「选项-加载项」中手动移除，或关闭所有 Excel 后重新运行卸载）'
    }
    $live = @(& $GetLiveExcel)
    if ($live.Count -gt 0) {
        $residue += ('仍有存活的 Excel（pid ' + ($live -join '、') + '）：卸载期间启动的会话退出时可能把加载项登记写回注册表（处置：关闭全部 Excel 后重新运行卸载）')
    }
    return $residue
}

if (-not $Uninstall) {
    if (-not (Test-Path $distXlam)) { throw "未找到 $distXlam ，请先运行 build.ps1 构建。" }

    New-Item -ItemType Directory -Force -Path $addInsDir | Out-Null
    try {
        Copy-Item $distXlam $destXlam -Force
        Unblock-File $destXlam -ErrorAction SilentlyContinue   # 清除下载标记（MOTW），防止宏被静默拦截
    } catch {
        # 占用方诊断：Excel 与 WPS 共存时，WPS 表格启动会沿用 Excel 的加载项注册把 AttachLink 一并加载，
        # 此时只关 Excel 不够（2026-11-23 实测：WPS 12.1 启动即加载库目录副本）
        $who = '文件可能被打开的 Excel 占用，请先关闭 Excel 后重试。'
        if (Get-Process -Name et -ErrorAction SilentlyContinue) {
            $who = '检测到 WPS 表格（et.exe）正在运行——WPS 启动时可能已自动加载 AttachLink 并占用该文件。请同时关闭 WPS 表格后重试。'
        }
        throw ('复制文件失败。' + $who)
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
    # 经宿主识别取得真正的 Excel（Excel.Application 被 WPS 接管时自动改走 CLSID；创建竞态重试见函数内）
    $excel = $null
    try {
        # ---- 先清掉指向其他路径的同名登记（机理见文件头「加载项列表登记」）----
        # 不清的话下方 AddIns.Add 会绑定到那份条目：文件已不在 → Installed 置位静默失败；文件还在 →
        # 登记的不是库目录副本。「卸载后重装」、以及此前只卸过用户级（机器级仍在）的机器会命中。
        # 库目录副本自身的登记（文件名 / 库目录完整路径）保留。清不干净就直接失败：带着旧登记拉起 Excel，
        # 绑定到哪一份就不受控了。（已清掉的登记不随后续的注册失败回滚——旧登记指向的本来就不是库目录
        # 副本，重跑安装即可继续。）
        if (-not (Clear-AddInListRegistrations -Mode OtherPaths)) {
            throw '加载项列表里指向其他路径的 AttachLink 登记未能清理干净（见上方警告）。'
        }

        $excel = New-RealExcel
        # AddIns.Add 在没有任何打开工作簿的纯 COM 实例上会报错，先建一个临时工作簿
        $wbTmp = $excel.Workbooks.Add()
        $ai = $excel.AddIns.Add($destXlam, $false)
        # AddIns.Add 按文件名绑定已有条目：确认绑定的是库目录副本。登记路径读不到时用 Path + Name 兜底，
        # 仍读不到则中止——无法确认绑定到哪一份时，不继续启用。
        $boundTo = ''
        try { $boundTo = [string]$ai.FullName } catch { $boundTo = '' }
        if (-not $boundTo) { try { $boundTo = Join-Path ([string]$ai.Path) ([string]$ai.Name) } catch { $boundTo = '' } }
        if (-not $boundTo) { throw '无法确认加载项列表里的「AttachLink.xlam」绑定到了哪一份（登记路径读不到），已中止，未启用加载项。请重新运行安装；仍失败请把日志反馈给分发人。' }
        if (-not (Test-SamePath $boundTo $destXlam)) {
            throw ('加载项列表里的「AttachLink.xlam」指向另一份副本（' + $boundTo + '），未能改为库目录副本。' +
                '请在 Excel「文件 → 选项 → 加载项 → 管理：Excel 加载项 → 转到」里取消勾选并移除该条目后重试。')
        }
        $ai.Installed = $true
        if (-not $ai.Installed) { throw ('加载项注册后未处于启用状态。' + (Get-DisabledItemsHint)) }
        $wbTmp.Close($false)
    } catch {
        Remove-Item $tl -Recurse -Force -ErrorAction SilentlyContinue
        throw "加载项注册失败，已回滚受信任位置写入。错误：$($_.Exception.Message)"
    } finally {
        if ($excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
    }
    Write-Host '安装完成: AttachLink 加载项已启用。'
    Write-Host '重新打开 Excel 后，在单元格上点击右键即可看到「插入附件链接（可多选）」。'

    # ---- 自动检测机器级部署是否需要同步（T6 落地：防止升级漏跑第二步）----
    $pfXlam = 'C:\Program Files\AttachLink\AttachLink.xlam'
    $pfWd   = 'C:\Program Files\AttachLink\watchdog.ps1'
    # v1.4.3 起看门狗带宿主校验（日志标记「仅附着 EXCEL.EXE」）；旧版会把 xlam 注入 WPS 表格
    $wdOutdated = $false
    if (Test-Path $pfWd) {
        try { $wdOutdated = -not ([IO.File]::ReadAllText($pfWd, [Text.Encoding]::UTF8).Contains('仅附着 EXCEL.EXE')) } catch {}
    }
    # 半部署：机器级目录里有 xlam 却没有看门狗脚本（部署中断 / 被杀软误删），鼎信诺场景会静默失效
    $wdMissing = (Test-Path $pfXlam) -and -not (Test-Path $pfWd)
    if ((Test-Path $pfXlam) -or $wdOutdated -or $wdMissing) {
        try {
            $hLocal = ''; $hPf = ''
            if (Test-Path $pfXlam) {
                $shaProv = [System.Security.Cryptography.SHA256]::Create()
                $fsA = [IO.File]::OpenRead($destXlam)
                try { $hLocal = [BitConverter]::ToString($shaProv.ComputeHash($fsA)) -replace '-', '' } finally { $fsA.Close() }
                $fsB = [IO.File]::OpenRead($pfXlam)
                try { $hPf = [BitConverter]::ToString($shaProv.ComputeHash($fsB)) -replace '-', '' } finally { $fsB.Close(); $shaProv.Dispose() }
            }
            if ($hLocal -ne $hPf -or $wdOutdated -or $wdMissing) {
                Write-Host ''
                Write-Host '==================================================================='
                if ($wdMissing) {
                    Write-Host '  ⚠ 机器级部署不完整：C:\Program Files\AttachLink\ 里有 AttachLink.xlam，却没有 watchdog.ps1（鼎信诺里不会出现右键菜单）'
                    Write-Host '  请重新运行安装器点「一键安装」补全：'
                } elseif ($wdOutdated) {
                    Write-Host '  ⚠ 检测到机器级看门狗仍为旧版：它会把加载项注入 WPS 表格（WPS 卡在「正在打开文档」或反复弹「同名文件已打开」）'
                    Write-Host '  v1.4.3 起看门狗只附着 EXCEL.EXE——请重新运行安装器点「一键安装」完成更新：'
                } else {
                    Write-Host '  ⚠ 检测到鼎信诺管理员环境部署与用户级版本不一致（需同步更新）'
                    Write-Host '  本机如需在鼎信诺中使用，请再点一次「一键安装」（覆盖安装加载项并部署鼎信诺环境，弹一次 UAC）：'
                }
                Write-Host '  AttachLinkSetup.exe →「一键安装」'
                Write-Host '  或管理员 PowerShell: powershell -ExecutionPolicy Bypass -File build\install-elevated.ps1'
                Write-Host '==================================================================='
            }
        } catch { }
    }
}
else {
    # ---- 1. 让 Excel 注销（Installed=False），并确认相关实例已退出（何时算「确认」见 Invoke-AddInUnregister）----
    $un = Invoke-AddInUnregister
    if (-not $un.Unregistered) {
        Write-Host ('未能完成加载项注销（未取得 Excel 实例，或访问加载项集合失败）。' + $(if ($un.Error) { '（' + $un.Error + '）' } else { '' }))
    }

    # ---- 2. 清理加载项列表登记（OPEN / OPENn 与 Add-in Manager，README E7）----
    # Excel 的 Installed=False 不保证登记被清掉：既有实现记录的是「只把 OPENn 标记为禁用（/A 前缀）并不删除
    # 条目」；本机实测库目录条目会在它退出时自行消失，非库目录条目（最常见的是机器级看门狗登记的
    # Program Files 副本）则挪进 Add-in Manager——文件被删后就是悬空条目，之后重新安装会绑定到它
    #（「卸载后重装失败」的根因，见文件头）。
    # 必须等第 1 步确认实例已退出之后再清：它退出时会把内存里的加载项列表写回注册表，抢在前面清会被
    # 覆盖回去，还可能留下「卸载成功」的假象。确认不了就不清，最后如实报失败。
    $listCleaned = $false
    if ($un.ExitConfirmed) {
        $listCleaned = Clear-AddInListRegistrations -Mode All
    } else {
        Write-Warning '未能确认用于注销的 Excel 实例已经退出，本次没有清理加载项列表登记（OPEN / Add-in Manager）。请关闭所有 Excel 后重新运行卸载。'
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
    # 受信任位置：遍历全部 Office 版本、两处可能写过的路径，逐个删、逐个核验（核验见 Get-UninstallResidue）
    foreach ($v in $officeVersions) {
        foreach ($k in @(
            "HKCU:\Software\Microsoft\Office\$v\Excel\Security\Trusted Locations\AttachLink",
            "HKCU:\Software\Microsoft\Office\$v\Excel\Trust Center\Trusted Locations\AttachLink"
        )) { Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue }
    }
    # 清理 COM 引导器
    Remove-Item 'HKCU:\Software\Microsoft\Office\Excel\Addins\AttachLinkShim.Connect' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item 'HKCU:\Software\Microsoft\Office\Excel\AddinsData\AttachLinkShim.Connect' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item 'HKCU:\Software\Classes\AttachLinkShim.Connect' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item 'HKCU:\Software\Classes\CLSID\{D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58}' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path $addInsDir 'AttachLinkShim.dll') -Force -ErrorAction SilentlyContinue

    # ---- 3. 按事实核验：有残留就如实报「未完全成功」（exit 10），全部干净才说卸载完成 ----
    $residue = @(Get-UninstallResidue -ListCleaned $listCleaned)
    $listResidue = @($residue | Where-Object { $_ -like '加载项列表登记*' })
    $otherResidue = @($residue | Where-Object { $_ -notlike '加载项列表登记*' })
    if ($otherResidue.Count -gt 0) {
        Write-Host '⚠ 卸载未完全成功，以下项目未能清除：'
        foreach ($r in $residue) { Write-Host ('  - ' + $r) }
        Write-Host '请关闭所有 Excel（含鼎信诺、WPS 表格）后重新运行卸载；仍失败则按上方清单手动处理。'
        exit 10
    }
    if (-not $listCleaned) {
        Write-Host '卸载完成（文件与注册表其余清理已执行且核验无残留），但加载项列表（OPEN / Add-in Manager）登记未能完全清理（见上方警告）。'
        exit 9
    }
    Write-Host '卸载完成（含加载项列表 OPEN/OPENn 残留清理、序列重排与 Add-in Manager 登记清理，副本 / 受信任位置 / 旧引导器残留均已核验清除）。'
    Write-Host '若曾在其他路径手动加载过本加载项副本，可在 Excel「选项-加载项-管理: Excel 加载项-转到」中手动移除。'
}
