# =====================================================================
# AttachLink 看门狗启动器回归（v1.6.3）
# 对象：build\make-launcher.ps1 的产物 dist\AttachLink.WatchdogLauncher.exe
# ---------------------------------------------------------------------
# 验证启动器作为计划任务直接动作的几项性质。全部用本测试自己的临时目录与进程；
# 看门狗脚本换成一段心跳脚本（把自己的进程号和父进程号写进日志，按标志文件以 7 或 0 退出）。
#   A. 产物：存在；PE 子系统为 GUI（2）；带 .NET 头的 AnyCPU 程序
#   B. 无窗口：启动器与它拉起的 PowerShell 都没有可见窗口，也没有拉起 Windows Terminal
#   C. 父子关系：看门狗的父进程就是启动器
#   D. 拉起策略：看门狗连续以 7 退出 → 重启两次（第 1/3、第 2/3）→ 第三次后启动器以 7 退出；日志三行「已结束」
#   E. 正常结束：看门狗以 0 退出 → 启动器不重启、返回 0
#   F. 不留孤儿：结束启动器后看门狗随之结束（作业对象）
#   G. 参数与部署自检：不认识的参数 → 5；旁边没有 watchdog.ps1 → 2
# 可选：-WithTaskIntegration 追加真机段（约 6 分钟；会注册一个当前用户的临时计划任务，结束后注销）：
#   H. 计划任务（按触发启动、带 1 分钟重复）以启动器为动作：无窗口、状态 Running；看门狗崩溃后
#      启动器 8 秒内拉起、任务保持 Running；连续三次后启动器放弃、任务以非 0 结束；下一次重复把任务
#      拉回来；结束启动器 → 看门狗随之结束、下一次重复再拉回来；Stop-ScheduledTask 后两者都不在。
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test-launcher.ps1 [-WithTaskIntegration]
#       期望末行 RESULT: PASS（需要 dist\AttachLink.WatchdogLauncher.exe，先运行 build\build.ps1）
#       退出码：通过 0，有失败 1
# 红线：只结束本测试自己启动的进程（心跳脚本记下的进程号与它们的父进程号），只注册 / 注销自己的临时任务；
#       不碰 wps.exe、et.exe、sjp7.exe 与机器上常驻的那个启动器（它由计划任务管理）；
#       启动器一律带 -LogFile 指向临时目录，不写进真实的看门狗日志。
# 编码：含中文，必须保存为 UTF-8 with BOM（WORKLOG 根因 #11）。
# =====================================================================
param([switch]$WithTaskIntegration)
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$repo = Split-Path -Parent $PSScriptRoot
$exe  = Join-Path $repo 'dist\AttachLink.WatchdogLauncher.exe'
$base = Join-Path $env:TEMP ('AttachLinkTestLauncher_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $base | Out-Null

$script:ok = $true
function Check($cond, $label) { if (-not $cond) { $script:ok = $false; Write-Host "FAIL: $label" } else { Write-Host "pass: $label" } }

Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class AttachLinkLauncherTestWin {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  public static List<string> Visible() {
    var r = new List<string>();
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); var c = new StringBuilder(128); GetClassName(h, c, 128);
      if (IsWindowVisible(h)) r.Add(h.ToInt64() + "|" + p + "|" + c.ToString()); return true; }, IntPtr.Zero);
    return r;
  }
}
'@

# 本测试自己启动的进程号（看门狗与它们的父进程，即启动器）；收尾只按这张表结束
$script:own = New-Object 'System.Collections.Generic.HashSet[int]'
function Remember-Own([int[]]$Pids) { foreach ($p in $Pids) { if ($p -gt 0) { [void]$script:own.Add($p) } } }

# 布置一个「部署目录」：启动器 + 心跳脚本（顶替 watchdog.ps1）+ 空的 AttachLink.xlam
function New-Stage([string]$Dir, [switch]$WithoutScript) {
    New-Item -ItemType Directory -Force -Path $Dir | Out-Null
    Copy-Item -LiteralPath $exe -Destination (Join-Path $Dir 'AttachLink.WatchdogLauncher.exe') -Force
    [IO.File]::WriteAllText((Join-Path $Dir 'AttachLink.xlam'), '')
    if (-not $WithoutScript) {
        $fake = @'
param([string]$XlamPath = '')
$ErrorActionPreference = 'Continue'
$LogFile = Join-Path $PSScriptRoot 'heartbeat.log'
while ($true) {
    try {
        $ppid = (Get-CimInstance Win32_Process -Filter "ProcessId=$PID").ParentProcessId
        [IO.File]::AppendAllText($LogFile, ('HB {0:HH:mm:ss.fff} pid={1} ppid={2}' -f (Get-Date), $PID, $ppid) + "`r`n")
    } catch {}
    if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'exit7.flag')) { exit 7 }
    if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'exit0.flag')) { exit 0 }
    Start-Sleep -Milliseconds 400
}
'@
        [IO.File]::WriteAllText((Join-Path $Dir 'watchdog.ps1'), $fake, (New-Object Text.UTF8Encoding($true)))
    }
    return (Join-Path $Dir 'AttachLink.WatchdogLauncher.exe')
}
# 读某个部署目录里第 $FromLine 行之后的心跳：pid、ppid
function Get-Hb([string]$Dir, [int]$FromLine) {
    $f = Join-Path $Dir 'heartbeat.log'
    if (-not (Test-Path -LiteralPath $f)) { return @() }
    $lines = [IO.File]::ReadAllLines($f)
    if ($lines.Count -le $FromLine) { return @() }
    return @($lines[$FromLine..($lines.Count - 1)] | Where-Object { $_ -match 'pid=(\d+) ppid=(\d+)' } |
        ForEach-Object { [pscustomobject]@{ pid = [int]$Matches[1]; ppid = [int]$Matches[2] } })
}
function Get-HbLineCount([string]$Dir) { $f = Join-Path $Dir 'heartbeat.log'; if (Test-Path -LiteralPath $f) { return [IO.File]::ReadAllLines($f).Count } else { return 0 } }
# 等某个部署目录出现心跳，返回第一条（pid、ppid）；超时返回 $null（至少查一次）。
# 只认进程还活着的心跳，并且不看 $ExcludePid 那个进程的心跳：被结束的看门狗在断气前还可能写下最后一行，
# 不排除的话它会被当成「新拉起的看门狗」。
function Wait-Hb([string]$Dir, [int]$FromLine, [int]$Seconds, [int]$ExcludePid = 0) {
    $t0 = Get-Date
    while ($true) {
        $h = @(Get-Hb $Dir $FromLine | Where-Object { $_.pid -ne $ExcludePid -and (Test-Alive ([int]$_.pid)) })
        if ($h.Count -gt 0) { return $h[0] }
        if (((Get-Date) - $t0).TotalSeconds -ge $Seconds) { return $null }
        Start-Sleep -Milliseconds 200
    }
}
# 观察 $Seconds 秒：启动器 / 看门狗名下，或控制台 / 终端窗口类，有没有新出现的可见窗口。返回描述行数组
function Watch-Windows([int[]]$Pids, [int]$Seconds) {
    $before = @([AttachLinkLauncherTestWin]::Visible() | ForEach-Object { $_.Split('|')[0] })
    $found = New-Object 'System.Collections.Generic.HashSet[string]'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $Seconds) {
        foreach ($w in [AttachLinkLauncherTestWin]::Visible()) {
            $f = $w.Split('|')
            if ($before -contains $f[0]) { continue }
            if (($Pids -contains [int]$f[1]) -or ($f[2] -match 'CASCADIA_HOSTING_WINDOW_CLASS|ConsoleWindowClass')) { [void]$found.Add($w) }
        }
        Start-Sleep -Milliseconds 20
    }
    return @($found)
}
function Test-Alive([int]$ProcId) { return [bool](Get-Process -Id $ProcId -ErrorAction SilentlyContinue) }
function Wait-Gone([int]$ProcId, [int]$Seconds) {
    $t0 = Get-Date
    while (((Get-Date) - $t0).TotalSeconds -lt $Seconds) { if (-not (Test-Alive $ProcId)) { return $true }; Start-Sleep -Milliseconds 250 }
    return $false
}
function Get-NewTerminals([int[]]$Before) { return @(Get-Process -Name WindowsTerminal -ErrorAction SilentlyContinue | Where-Object { $Before -notcontains $_.Id } | ForEach-Object { $_.Id }) }

'=== A. 产物 ==='
if (-not (Test-Path -LiteralPath $exe)) {
    Check $false ("找不到 {0}，先运行 build\build.ps1" -f $exe)
} else {
    $b = [IO.File]::ReadAllBytes($exe)
    $isPe = ($b.Length -gt 256 -and $b[0] -eq 0x4D -and $b[1] -eq 0x5A)
    Check $isPe '产物是 PE 文件'
    if ($isPe) {
        $peOff = [BitConverter]::ToInt32($b, 0x3C)
        $subsystem = [BitConverter]::ToUInt16($b, $peOff + 4 + 20 + 68)
        $machine = [BitConverter]::ToUInt16($b, $peOff + 4)
        $magic = [BitConverter]::ToUInt16($b, $peOff + 24)
        $ddOff = 0; if ($magic -eq 0x20B) { $ddOff = $peOff + 24 + 112 } else { $ddOff = $peOff + 24 + 96 }
        $clrRva = [BitConverter]::ToUInt32($b, $ddOff + 14 * 8)
        Check ($subsystem -eq 2) 'PE 子系统 = 2（GUI）：计划任务启动它时没有控制台窗口'
        Check ($machine -eq 0x14C -and $clrRva -ne 0) 'AnyCPU 的 .NET 程序（带 CLR 头）'
    }
}

'=== B / C. 启动：无窗口、父子关系 ==='
$stage1 = Join-Path $base 's1'
$lx1 = New-Stage $stage1
$termBefore = @(Get-Process -Name WindowsTerminal -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$L1 = Start-Process -FilePath $lx1 -ArgumentList '-LogFile "launcher.log" -RestartDelaySeconds 1' -WorkingDirectory $stage1 -PassThru
Remember-Own @($L1.Id)
$hb1 = Wait-Hb $stage1 0 15
if ($null -eq $hb1) {
    Check $false '启动器拉起了看门狗（15 秒内出现心跳）'
} else {
    Remember-Own @([int]$hb1.pid)
    Check ($L1.Id -eq [int]$hb1.ppid) ('看门狗的父进程就是启动器（看门狗 pid={0} 父={1} 启动器 pid={2}）' -f $hb1.pid, $hb1.ppid, $L1.Id)
    $wins = Watch-Windows @($L1.Id, [int]$hb1.pid) 3
    Check ($wins.Count -eq 0) ('启动器与看门狗名下没有可见窗口，也没有控制台 / 终端窗口类的新窗口（3 秒，约 20 毫秒一采）' + $(if ($wins.Count -gt 0) { '；看到：' + ($wins -join ' ; ') } else { '' }))
    Check ((Get-NewTerminals $termBefore).Count -eq 0) '没有拉起 Windows Terminal'

    '=== D. 拉起策略 ==='
    Set-Content -LiteralPath (Join-Path $stage1 'exit7.flag') -Value 'x' -Encoding ASCII
    $tD = Get-Date
    $exited = $L1.WaitForExit(40000)
    $all1 = @(Get-Hb $stage1 0)
    $kids1 = @($all1 | Where-Object { $_.ppid -eq $L1.Id } | ForEach-Object { $_.pid } | Sort-Object -Unique)
    Remember-Own ([int[]]$kids1)
    Check $exited '看门狗连续以 7 退出后，启动器结束'
    Check ($exited -and $L1.ExitCode -eq 7) ('启动器的退出码 = 看门狗的退出码（实际 {0}，期望 7）' -f $(if ($exited) { $L1.ExitCode } else { 'n/a' }))
    Check ($kids1.Count -eq 3) ('启动器一共拉起 3 个看门狗：首次 + 两次重启（实际 {0}：{1}）' -f $kids1.Count, ($kids1 -join ','))
    $ll = Join-Path $stage1 'launcher.log'
    $llText = ''; if (Test-Path -LiteralPath $ll) { $llText = [IO.File]::ReadAllText($ll, [Text.Encoding]::UTF8) }
    Check (([regex]::Matches($llText, '看门狗已结束（退出码=7')).Count -eq 3) '启动器日志：三行「看门狗已结束（退出码=7…）」'
    Check ($llText.Contains('秒后重启看门狗（第 1/3 次）') -and $llText.Contains('秒后重启看门狗（第 2/3 次）') -and -not $llText.Contains('第 3/3 次')) '启动器日志：重启两次（第 1/3、第 2/3），没有第 3/3 次'
    Check ($llText.Contains('连续 3 次没能稳定运行，启动器退出（退出码=7）')) '启动器日志：第三次后放弃并带出退出码'
    Check (((Get-Date) - $tD).TotalSeconds -ge 2) ('两次重启之间各等 1 秒（总耗时 {0:N1} 秒）' -f ((Get-Date) - $tD).TotalSeconds)
    Check (@($kids1 | Where-Object { Test-Alive $_ }).Count -eq 0) '启动器结束后没有残留的看门狗'
    Remove-Item -LiteralPath (Join-Path $stage1 'exit7.flag') -Force -ErrorAction SilentlyContinue

    '=== E. 正常结束 ==='
    $fromE = Get-HbLineCount $stage1
    $L1e = Start-Process -FilePath $lx1 -ArgumentList '-LogFile "launcher.log" -RestartDelaySeconds 1' -WorkingDirectory $stage1 -PassThru
    Remember-Own @($L1e.Id)
    $hbE = Wait-Hb $stage1 $fromE 15
    if ($null -eq $hbE) { Check $false '第二次启动：15 秒内出现心跳' } else {
        Remember-Own @([int]$hbE.pid)
        Set-Content -LiteralPath (Join-Path $stage1 'exit0.flag') -Value 'x' -Encoding ASCII
        $exitedE = $L1e.WaitForExit(20000)
        $kidsE = @(Get-Hb $stage1 $fromE | Where-Object { $_.ppid -eq $L1e.Id } | ForEach-Object { $_.pid } | Sort-Object -Unique)
        Remember-Own ([int[]]$kidsE)
        Check ($exitedE -and $L1e.ExitCode -eq 0) ('看门狗以 0 结束：启动器返回 0、不重启（实际 {0}）' -f $(if ($exitedE) { $L1e.ExitCode } else { 'n/a' }))
        Check ($kidsE.Count -eq 1) ('只拉起过一个看门狗（实际 {0}）' -f $kidsE.Count)
        Remove-Item -LiteralPath (Join-Path $stage1 'exit0.flag') -Force -ErrorAction SilentlyContinue
    }

    '=== F. 不留孤儿 ==='
    $fromF = Get-HbLineCount $stage1
    $L1f = Start-Process -FilePath $lx1 -ArgumentList '-LogFile "launcher.log" -RestartDelaySeconds 1' -WorkingDirectory $stage1 -PassThru
    Remember-Own @($L1f.Id)
    $hbF = Wait-Hb $stage1 $fromF 15
    if ($null -eq $hbF) { Check $false '第三次启动：15 秒内出现心跳' } else {
        Remember-Own @([int]$hbF.pid)
        Start-Sleep -Seconds 2
        Stop-Process -Id $L1f.Id -Force -ErrorAction SilentlyContinue
        $goneF = Wait-Gone ([int]$hbF.pid) 8
        Check $goneF '结束启动器后，看门狗 8 秒内随之结束（作业对象）'
    }
}

'=== G. 参数与部署自检 ==='
$lxNoScript = New-Stage (Join-Path $base 's2') -WithoutScript
$pNo = Start-Process -FilePath $lxNoScript -ArgumentList '-LogFile "launcher.log"' -WorkingDirectory (Join-Path $base 's2') -PassThru
Remember-Own @($pNo.Id)
[void]$pNo.WaitForExit(15000)
Check ($pNo.HasExited -and $pNo.ExitCode -eq 2) ('旁边没有 watchdog.ps1 → 退出码 2（实际 {0}）' -f $(if ($pNo.HasExited) { $pNo.ExitCode } else { 'n/a' }))
$stage3 = Join-Path $base 's3'
$lx3 = New-Stage $stage3
$pBad = Start-Process -FilePath $lx3 -ArgumentList '-LogFile "launcher.log" -Bogus' -WorkingDirectory $stage3 -PassThru
Remember-Own @($pBad.Id)
[void]$pBad.WaitForExit(15000)
Check ($pBad.HasExited -and $pBad.ExitCode -eq 5) ('不认识的参数 → 退出码 5（实际 {0}）' -f $(if ($pBad.HasExited) { $pBad.ExitCode } else { 'n/a' }))
Check (@(Get-Hb $stage3 0).Count -eq 0) '参数不对时没有启动看门狗'

if ($WithTaskIntegration) {
    '=== H. 计划任务真机段（按触发启动 + 1 分钟重复）==='
    $stageH = Join-Path $base 'sH'
    $lxH = New-Stage $stageH
    $taskName = 'AttachLink Test Launcher ' + [guid]::NewGuid().ToString('N').Substring(0, 8)
    $registered = $false
    try {
        # 启动器日志用绝对路径：计划任务启动进程时当前目录不是部署目录（测试里也不能依赖它）
        $actionH = New-ScheduledTaskAction -Execute $lxH -Argument ('-LogFile "{0}" -RestartDelaySeconds 1' -f (Join-Path $stageH 'launcher.log'))
        # 重复间隔 1 分钟、时长 1 天（[TimeSpan]::MaxValue 会被任务计划程序判为超出范围）；窗口结束时不停任务
        $triggerH = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(30) -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 1)
        $triggerH.Repetition.StopAtDurationEnd = $false
        $principalH = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
        $settingsH = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Seconds 0) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
        $regH = Register-ScheduledTask -TaskName $taskName -Action $actionH -Trigger $triggerH -Principal $principalH -Settings $settingsH -Force -ErrorAction Stop
        $registered = $true
        $visBeforeH = @([AttachLinkLauncherTestWin]::Visible() | ForEach-Object { $_.Split('|')[0] })
        $foundH = New-Object 'System.Collections.Generic.HashSet[string]'
        $hbH = $null
        $t0H = Get-Date
        while (((Get-Date) - $t0H).TotalSeconds -lt 110 -and $null -eq $hbH) {
            $hs = @(Get-Hb $stageH 0)
            if ($hs.Count -gt 0) { $hbH = $hs[0] } else {
                foreach ($w in [AttachLinkLauncherTestWin]::Visible()) {
                    $f = $w.Split('|')
                    if (($visBeforeH -notcontains $f[0]) -and ($f[2] -match 'CASCADIA_HOSTING_WINDOW_CLASS|ConsoleWindowClass')) { [void]$foundH.Add($w) }
                }
                Start-Sleep -Milliseconds 20
            }
        }
        if ($null -eq $hbH) {
            Check $false '触发后 110 秒内启动器拉起了看门狗'
        } else {
            Remember-Own @([int]$hbH.pid, [int]$hbH.ppid)
            Start-Sleep -Seconds 3
            Check ((Get-Process -Id ([int]$hbH.ppid) -ErrorAction SilentlyContinue).ProcessName -eq 'AttachLink.WatchdogLauncher') ('按触发启动：看门狗 pid={0} 的父进程是本测试的启动器' -f $hbH.pid)
            $winsH = @(Watch-Windows @([int]$hbH.ppid, [int]$hbH.pid) 2)
            Check ($winsH.Count -eq 0 -and $foundH.Count -eq 0) '任务启动时没有出现控制台 / 终端窗口'
            $taskH = Get-ScheduledTask -TaskName $taskName
            Check ("$($taskH.State)" -eq 'Running') ('看门狗运行期间任务状态 = Running（实际 {0}）' -f $taskH.State)
            # 崩溃三次：前两次启动器 8 秒内拉起、任务保持 Running；第三次后放弃
            $launcherH = [int]$hbH.ppid
            $lastChild = [int]$hbH.pid
            for ($n = 1; $n -le 3; $n++) {
                if (-not (Test-Alive $lastChild)) { Check $false ('第 {0} 次崩溃前没有在运行的看门狗' -f $n); break }
                $victimH = $lastChild
                $swKillH = [Diagnostics.Stopwatch]::StartNew()
                Stop-Process -Id $victimH -Force -ErrorAction SilentlyContinue
                # 先等它确实退出再记行数（否则它临终写下的心跳会被当成新拉起的看门狗）；8 秒从结束它的那一刻算起
                [void](Wait-Gone $victimH 5)
                $fromN = Get-HbLineCount $stageH
                $newH = Wait-Hb $stageH $fromN ([int][Math]::Ceiling(8 - $swKillH.Elapsed.TotalSeconds)) -ExcludePid $victimH
                if ($n -lt 3) {
                    if ($null -eq $newH) { Check $false ('第 {0} 次崩溃后 8 秒内启动器拉起了看门狗' -f $n) } else {
                        Remember-Own @([int]$newH.pid)
                        Check ([int]$newH.ppid -eq $launcherH) ('第 {0} 次崩溃后 8 秒内由同一个启动器拉起（新看门狗 pid={1}）' -f $n, $newH.pid)
                        $stN = Get-ScheduledTask -TaskName $taskName
                        Check ("$($stN.State)" -eq 'Running') ('第 {0} 次崩溃后任务仍是 Running' -f $n)
                        $lastChild = [int]$newH.pid
                    }
                } else {
                    Check ($null -eq $newH) '第三次崩溃后启动器放弃：8 秒内没有新的看门狗'
                }
            }
            Start-Sleep -Seconds 4
            $stEnd = Get-ScheduledTask -TaskName $taskName; $infoEnd = Get-ScheduledTaskInfo -TaskName $taskName
            Check ("$($stEnd.State)" -ne 'Running' -and $infoEnd.LastTaskResult -ne 0 -and $infoEnd.LastTaskResult -ne 267009) ('启动器放弃后任务结束，结果码非 0（State={0} LastTaskResult={1}）' -f $stEnd.State, $infoEnd.LastTaskResult)
            $llH = Join-Path $stageH 'launcher.log'; $llHText = ''; if (Test-Path -LiteralPath $llH) { $llHText = [IO.File]::ReadAllText($llH, [Text.Encoding]::UTF8) }
            Check ($llHText.Contains('没能稳定运行，启动器退出')) '启动器日志记了放弃'
            # 下一次重复把任务拉回来（≤ 150 秒）
            $fromR = Get-HbLineCount $stageH
            $backH = Wait-Hb $stageH $fromR 150
            if ($null -eq $backH) { Check $false '启动器放弃后 150 秒内，任务重复把启动器拉回来' } else {
                Remember-Own @([int]$backH.pid, [int]$backH.ppid)
                Check ([int]$backH.ppid -ne $launcherH) ('任务重复拉回了新的启动器（pid={0}）与看门狗（pid={1}）' -f $backH.ppid, $backH.pid)
                # 结束启动器 → 看门狗随之结束 → 下一次重复再拉回来
                Stop-Process -Id ([int]$backH.ppid) -Force -ErrorAction SilentlyContinue
                Check (Wait-Gone ([int]$backH.pid) 8) '结束启动器后看门狗随之结束'
                $fromR2 = Get-HbLineCount $stageH
                $backH2 = Wait-Hb $stageH $fromR2 150
                if ($null -eq $backH2) { Check $false '结束启动器后 150 秒内，任务重复再把启动器拉回来' } else {
                    Remember-Own @([int]$backH2.pid, [int]$backH2.ppid)
                    Check ([int]$backH2.ppid -ne [int]$backH.ppid) ('任务重复又拉回了新的启动器（pid={0}）' -f $backH2.ppid)
                    Stop-ScheduledTask -TaskName $taskName
                    $goneL = Wait-Gone ([int]$backH2.ppid) 15
                    $goneC = Wait-Gone ([int]$backH2.pid) 15
                    Check ($goneL -and $goneC) 'Stop-ScheduledTask 后启动器与看门狗都不在（15 秒内）'
                }
            }
        }
    } catch {
        Check $false ("H 段执行异常: " + $_.Exception.Message.Split("`n")[0])
    } finally {
        if ($registered) {
            try { Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue } catch {}
            Start-Sleep -Seconds 3
            try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop } catch { Check $false ('注销临时计划任务失败: ' + $_.Exception.Message.Split("`n")[0]) }
        }
    }
}

'--- 清理 ---'
Start-Sleep -Seconds 2
foreach ($p in @($script:own)) { if (Test-Alive $p) { try { Stop-Process -Id $p -Force -ErrorAction Stop; Write-Host ("    已清理残留测试进程 pid={0}" -f $p) } catch {} } }
try { Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue } catch {}
Write-Host ("RESULT: " + $(if ($script:ok) { 'PASS' } else { 'FAIL' }))
if ($script:ok) { exit 0 } else { exit 1 }
