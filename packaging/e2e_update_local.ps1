# =====================================================================
# AttachLink 一键更新 —— 本地端到端验证（不碰外网、不做真实安装）
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File packaging\e2e_update_local.ps1
#       （可选 -Version / -Python / -Port，默认版本取自 src\AttachLink.bas）
# 步骤:
#   1) 运行 build\make-manifest.ps1 生成发布 zip 与 latest.json
#   2) 把 exe/zip/清单复制到临时目录，改写清单 url 指向本地回环地址并用发布私钥重新签名
#      （安装器的「--manifest-url 指向回环」测试模式允许 http 并放宽下载地址白名单，但清单验签无旁路）
#   3) 起本地 http.server（记录 PID，结束时结束它）
#   4) 依次验证：/check-update（读清单）→ /check-update（构造 v9.9.9 → exit 20）→
#      /update --update-dry-run（下载/校验/解压）→ /selfcheck --updated-from --wait-pid 交接模拟
# 期望末行 RESULT: PASS
# =====================================================================
param(
    [string]$Python = 'C:\Python314\python.exe',
    [string]$Version = '',
    [int]$Port = 0
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$script:ok = $true
function Check($cond, $label) {
    if ($cond) { Write-Host "pass: $label" } else { $script:ok = $false; Write-Host "FAIL: $label" }
}
function Get-Sha256Hex([string]$Path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [IO.File]::OpenRead($Path)
        try { ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() }
        finally { $stream.Dispose() }
    } finally { $sha.Dispose() }
}

$root = Split-Path -Parent $PSScriptRoot
# 点源 build\make-manifest.ps1：复用 New-ManifestSignature 给本地回环清单重签名
# （篡改 url/version 后原签名即失效；v1.5.0 起安装器对清单验签，本地测试模式同样验签、无旁路）
. (Join-Path $root 'build\make-manifest.ps1')

$srcBas = Join-Path $root 'src\AttachLink.bas'
if (-not $Version) {
    $basText = [IO.File]::ReadAllText($srcBas, [Text.Encoding]::UTF8)
    if ($basText -notmatch 'ATTACHLINK_VERSION\s+As\s+String\s*=\s*"([^"]+)"') { throw '未能从 AttachLink.bas 提取版本号' }
    $Version = $Matches[1]
}
$releaseDir = Join-Path $root "release\AttachLink-$Version"
$exe = Join-Path $releaseDir 'AttachLinkSetup.exe'
if (-not (Test-Path $exe)) { throw "未找到 $exe（请先跑 build\build.ps1 → build\make-exe.ps1）" }

# 占位进程 / 服务进程的 PID 记录在这里，finally 里只结束它们
$script:serverProc = $null
$script:placeholder = $null
$serveDir = Join-Path $env:TEMP "AttachLink_e2e_serve"
$handDir  = Join-Path $env:TEMP "AttachLink_e2e_handoff"

# 发布私钥：只在此读取用于给本地回环清单签名（与 make-manifest.ps1 同一纪律）
$keyFile = Join-Path $env:USERPROFILE '.attachlink\release-key.xml'
$script:haveKey = Test-Path $keyFile
if ($script:haveKey) {
    $script:keyXml = [IO.File]::ReadAllText($keyFile)
} else {
    Write-Host "未找到发布私钥 $keyFile——4a/4b/4c（签名正向链路）将 SKIP；4e 负向用例照常执行"
}

function Sign-LocalManifest($m) {
    # 改写过 url/version 的清单必须重新签名（签名覆盖 version/url/size/两个哈希/notes_url）
    $m.sig_alg = 'RS256-SHA256-PKCS1v15:v1'
    $m.signature = New-ManifestSignature -Version ([string]$m.version) -Url ([string]$m.url) `
        -SizeZip $m.size_zip -Sha256Zip ([string]$m.sha256_zip) -Sha256Exe ([string]$m.sha256_exe) `
        -NotesUrl ([string]$m.notes_url) -KeyXml $script:keyXml
    return $m
}

try {
    # ---- 1) 生成发布 zip 与 latest.json ----
    powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path (Split-Path -Parent $PSScriptRoot) 'build\make-manifest.ps1')
    if ($LASTEXITCODE -ne 0) { throw 'make-manifest.ps1 失败' }
    $zipPath = Join-Path $root "release\AttachLink-$Version.zip"
    $manifestPath = Join-Path $releaseDir 'latest.json'
    Check (Test-Path $zipPath) "make-manifest：发布 zip 已生成"
    Check (Test-Path $manifestPath) "make-manifest：latest.json 已生成"
    $m0 = Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Check ($m0.version -eq $Version) "make-manifest：清单 version = $Version"
    Check ($m0.sha256_zip -eq (Get-Sha256Hex $zipPath)) "make-manifest：清单 sha256_zip 与 zip 实际哈希一致"
    Check ($m0.sha256_exe -eq (Get-Sha256Hex $exe)) "make-manifest：清单 sha256_exe 与 exe 实际哈希一致"
    if ((Get-Item $zipPath).Length -ne $m0.size_zip) { throw 'zip 大小与清单不符' }

    # ---- 2) 本地服务目录 + 回环清单 ----
    if ($Port -eq 0) {
        $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
        $listener.Start()
        $Port = ([Net.IPEndPoint]$listener.LocalEndpoint).Port
        $listener.Stop()
    }
    Remove-Item $serveDir -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $serveDir | Out-Null
    Copy-Item $zipPath $serveDir -Force
    $mLocal = Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $mLocal.url = "http://127.0.0.1:$Port/AttachLink-$Version.zip"
    if ($script:haveKey) { $mLocal = Sign-LocalManifest $mLocal }
    $localManifest = Join-Path $serveDir 'latest.json'
    [IO.File]::WriteAllText($localManifest, (ConvertTo-Json -InputObject $mLocal), [Text.UTF8Encoding]::new($false))
    # 构造一个「未来版本」清单，验证 /check-update 的「有新版本 → exit 20」分支与 /update 的下载链路
    $mNew = Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $mNew.version = '9.9.9'
    $mNew.url = "http://127.0.0.1:$Port/AttachLink-$Version.zip"   # 下载也走本地（回环清单的放宽只对本地测试生效）
    if ($script:haveKey) { $mNew = Sign-LocalManifest $mNew }
    $newManifest = Join-Path $serveDir 'latest-new.json'
    [IO.File]::WriteAllText($newManifest, (ConvertTo-Json -InputObject $mNew), [Text.UTF8Encoding]::new($false))

    # ---- 3) 启动本地 http.server（记录 PID）----
    $script:serverProc = Start-Process -FilePath $Python `
        -ArgumentList @('-m', 'http.server', "$Port", '--bind', '127.0.0.1', '--directory', $serveDir) `
        -WindowStyle Hidden -PassThru
    Write-Host "本地服务 PID=$($script:serverProc.Id) 端口=$Port 目录=$serveDir"
    $ready = $false
    foreach ($i in 1..30) {
        try {
            $resp = Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:$Port/latest.json" -TimeoutSec 2
            if ($resp.StatusCode -eq 200) { $ready = $true; break }
        } catch { Start-Sleep -Milliseconds 300 }
    }
    Check $ready "本地服务就绪（GET latest.json → 200）"

    if ($script:haveKey) {
        # ---- 4a) /check-update：读本地清单并解析（当前版本 = 清单版本 → exit 0）----
        $p1 = Start-Process -FilePath $exe -Wait -PassThru `
            -ArgumentList @('/check-update', '--manifest-url', "http://127.0.0.1:$Port/latest.json")
        Check ($p1.ExitCode -eq 0) "check-update：本地回环清单已重签名、可读可解析（exit=$($p1.ExitCode)，当前 v$Version 与清单同版 → 必须恰好为 0）"

        # ---- 4b) /check-update：清单版本 9.9.9 → 有新版本 → exit 20 ----
        $p1b = Start-Process -FilePath $exe -Wait -PassThru `
            -ArgumentList @('/check-update', '--manifest-url', "http://127.0.0.1:$Port/latest-new.json")
        Check ($p1b.ExitCode -eq 20) "check-update：发现新版本 v9.9.9（exit=$($p1b.ExitCode)）"

        # ---- 4c) /update --update-dry-run：下载 → 校验 → 解压（不安装；用「未来版本」清单触发下载）----
        # v1.5.0：更新临时目录按次隔离（根目录下每次运行一个 GUID 子目录，并发更新互不干扰）
        $updRoot = Join-Path $env:TEMP 'AttachLinkSetup_update'
        Remove-Item $updRoot -Recurse -Force -ErrorAction SilentlyContinue
        $p2 = Start-Process -FilePath $exe -Wait -PassThru `
            -ArgumentList @('/update', '--update-dry-run', '--manifest-url', "http://127.0.0.1:$Port/latest-new.json")
        Check ($p2.ExitCode -eq 0) "update dry-run：exit=$($p2.ExitCode)"
        $subDirs = @(Get-ChildItem $updRoot -Directory -ErrorAction SilentlyContinue)
        Check ($subDirs.Count -eq 1) "update dry-run：临时目录按次隔离（恰好一个本次 GUID 子目录，实际 $($subDirs.Count) 个）"
        $updDir = if ($subDirs.Count -ge 1) { $subDirs[0].FullName } else { $updRoot }
        $zipDown = Join-Path $updDir "AttachLink-$($mNew.version).zip"   # 文件名按清单版本命名（此处 9.9.9）
        $exeDown = Join-Path $updDir 'AttachLinkSetup.exe'
        Check (Test-Path $zipDown) "update dry-run：发布 zip 已下载到临时目录"
        Check ((Get-Item $zipDown).Length -eq $m0.size_zip) "update dry-run：zip 大小与清单一致"
        Check ((Get-Sha256Hex $zipDown) -eq $m0.sha256_zip) "update dry-run：zip SHA-256 与清单一致"
        Check ((Get-Sha256Hex $exeDown) -eq $m0.sha256_exe) "update dry-run：解压出的安装器 SHA-256 与清单一致"
    } else {
        Write-Host "SKIP 4a/4b/4c：无私钥无法构造合法的本地签名清单（make-manifest.ps1 同样要求私钥）"
    }

    # ---- 4e) 负向：清单完整性 ----
    # 4e-1) 篡改已签清单的 sha256_exe（保留签名）→ /check-update 必须 21，日志含「签名校验失败」
    $mBad = Get-Content $newManifest -Raw -Encoding UTF8 | ConvertFrom-Json
    $h0 = [string]$mBad.sha256_exe
    $firstHex = if ($h0[0] -eq '0') { '1' } else { '0' }
    $mBad.sha256_exe = $firstHex + $h0.Substring(1)
    $badManifest = Join-Path $serveDir 'latest-bad.json'
    [IO.File]::WriteAllText($badManifest, (ConvertTo-Json -InputObject $mBad), [Text.UTF8Encoding]::new($false))
    $updRoot = Join-Path $env:TEMP 'AttachLinkSetup_update'
    $pE1 = Start-Process -FilePath $exe -Wait -PassThru `
        -ArgumentList @('/check-update', '--manifest-url', "http://127.0.0.1:$Port/latest-bad.json")
    Check ($pE1.ExitCode -eq 21) "负向-篡改哈希：/check-update 拒绝（exit=$($pE1.ExitCode)，期望 21）"
    $logFile = Join-Path $env:TEMP 'AttachLinkSetup.log'
    $logBad = if (Test-Path $logFile) { Get-Content $logFile -Raw -Encoding UTF8 } else { '' }
    Check ($logBad -match '签名校验失败') "负向-篡改哈希：日志含「签名校验失败」（镜像被篡改或损坏的处置路径）"
    # 4e-2) 同一被篡改清单跑 /update dry-run：必须失败（22）且没有发生任何下载
    Remove-Item $updRoot -Recurse -Force -ErrorAction SilentlyContinue
    $pE2 = Start-Process -FilePath $exe -Wait -PassThru `
        -ArgumentList @('/update', '--update-dry-run', '--manifest-url', "http://127.0.0.1:$Port/latest-bad.json")
    Check ($pE2.ExitCode -eq 22) "负向-篡改哈希：/update dry-run 失败（exit=$($pE2.ExitCode)，期望 22）"
    $zipsLeft = @(Get-ChildItem $updRoot -Recurse -Filter '*.zip' -File -ErrorAction SilentlyContinue)
    Check ($zipsLeft.Count -eq 0) "负向-篡改哈希：验签失败 → 未发生任何下载（临时目录 zip 数=$($zipsLeft.Count)）"
    # 4e-3) 完全不带 signature 字段的清单 → 同样拒绝（21）
    $mNoSig = Get-Content $newManifest -Raw -Encoding UTF8 | ConvertFrom-Json
    $mNoSig.PSObject.Properties.Remove('signature')
    $noSigManifest = Join-Path $serveDir 'latest-nosig.json'
    [IO.File]::WriteAllText($noSigManifest, (ConvertTo-Json -InputObject $mNoSig), [Text.UTF8Encoding]::new($false))
    $pE3 = Start-Process -FilePath $exe -Wait -PassThru `
        -ArgumentList @('/check-update', '--manifest-url', "http://127.0.0.1:$Port/latest-nosig.json")
    Check ($pE3.ExitCode -eq 21) "负向-无签名字段：/check-update 拒绝（exit=$($pE3.ExitCode)，期望 21）"

    # ---- 4d) 交接模拟：--updated-from + --wait-pid ----
    Remove-Item $handDir -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $handDir | Out-Null
    $oldExe = Join-Path $handDir 'AttachLinkSetup.exe'
    Copy-Item $exe $oldExe -Force
    # 把其中一个字节改掉，使「旧 exe 副本」与新版不同（便于证明替换真的发生；该副本不会被运行）
    $bytes = [IO.File]::ReadAllBytes($oldExe)
    $flip = [int]([Math]::Floor($bytes.Length / 2))
    $bytes[$flip] = $bytes[$flip] -bxor 0xFF
    [IO.File]::WriteAllBytes($oldExe, $bytes)
    $oldShaBefore = Get-Sha256Hex $oldExe
    Check ($oldShaBefore -ne $m0.sha256_exe) "交接模拟：旧 exe 副本已被改到与新版不同"

    # 占位进程：模拟「还没退出的旧安装器」（8 秒后自行退出；替换等待上限 15 秒 > 8 秒）
    $script:placeholder = Start-Process -FilePath 'powershell' `
        -ArgumentList @('-NoProfile', '-Command', 'Start-Sleep -Seconds 8') -PassThru -WindowStyle Hidden
    $phPid = $script:placeholder.Id
    Write-Host "占位进程 PID=$phPid（Start-Sleep 8）"
    $p3 = Start-Process -FilePath $exe -Wait -PassThru `
        -ArgumentList @('/selfcheck', '--updated-from', $oldExe, '--wait-pid', "$phPid")
    Write-Host "selfcheck(带交接) exit=$($p3.ExitCode)（/selfcheck 的结果如实记录）"
    Check ($p3.ExitCode -eq 0) "交接模拟：/selfcheck 自身通过（exit=$($p3.ExitCode)）"

    # 等占位进程自然退出（只结束自己记录的这个 PID）
    if (-not $script:placeholder.WaitForExit(25000)) {
        Stop-Process -Id $phPid -Force -ErrorAction SilentlyContinue
    }
    $script:placeholder = $null
    $oldShaAfter = Get-Sha256Hex $oldExe
    Check ($oldShaAfter -eq $m0.sha256_exe) "交接模拟：占位进程退出后旧 exe 副本已被替换为新版（哈希一致）"
    Check (-not (Test-Path "$oldExe.bak")) "交接模拟：无 .bak 残留"
    $logText = ''
    $logFile = Join-Path $env:TEMP 'AttachLinkSetup.log'
    if (Test-Path $logFile) { $logText = Get-Content $logFile -Raw -Encoding UTF8 }
    Check ($logText -match '\[交接\]') "交接模拟：日志含 [交接] 替换记录"
    Check ($logText -match ('旧安装器进程（PID ' + $phPid + '）已退出')) "交接模拟：日志记录了等待旧进程退出"
}
finally {
    # 只结束自己创建并记录了 PID 的进程
    if ($script:serverProc -and -not $script:serverProc.HasExited) {
        Stop-Process -Id $script:serverProc.Id -Force -ErrorAction SilentlyContinue
    }
    if ($script:placeholder -and -not $script:placeholder.HasExited) {
        Stop-Process -Id $script:placeholder.Id -Force -ErrorAction SilentlyContinue
    }
    Remove-Item $serveDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $handDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path $env:TEMP 'AttachLinkSetup_update') -Recurse -Force -ErrorAction SilentlyContinue   # 根目录连同本次 GUID 子目录一并清场
    # exe 运行会把日志快照写到发布目录，清掉（发布目录只留 exe / README / latest.json / *.log 规则外文件）
    Remove-Item (Join-Path $releaseDir 'AttachLinkSetup.log') -Force -ErrorAction SilentlyContinue
}

if ($script:ok) { Write-Host 'RESULT: PASS' } else { Write-Host 'RESULT: FAIL'; exit 1 }
