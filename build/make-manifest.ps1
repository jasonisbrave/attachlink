# =====================================================================
# AttachLink 发布清单与分发包生成脚本（发布方使用，v1.5.0 起）
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File build\make-manifest.ps1
#       可选: -ReleaseDir <发布目录>  -Version <版本号>（默认从 src\AttachLink.bas 读取）
#              -KeyFile <私钥 XML>（默认 %USERPROFILE%\.attachlink\release-key.xml）
# 产物:
#   release\AttachLink-<版本>.zip     只含 AttachLinkSetup.exe + README.md 两个顶层条目（无目录前缀）
#   release\AttachLink-<版本>\latest.json  发布清单（UTF-8 无 BOM；随版本作为 Release 附件上传；含签名）
# 说明:
#   - zip 用 .NET System.IO.Compression 逐条目打包；目录里的 *.log 忽略，其它意外文件打印警告
#   - SHA-256 用 .NET 直调计算（不用 Get-FileHash）；结束时打印 zip / exe / latest.json 三者哈希
#   - 清单按「RS256-SHA256-PKCS1v15:v1」用发布私钥签名（build\new-release-key.ps1 生成密钥对）；
#     写完后自动调用 packaging\setup_update.py --verify-manifest 做互操作自检，不通过直接失败
#   - New-ManifestSignature / Get-Sha256Hex 可供其它脚本点源调用（如 e2e_update_local.ps1 造签名测试清单）
# =====================================================================
param(
    [string]$ReleaseDir = "",
    [string]$Version = "",
    [string]$KeyFile = "$env:USERPROFILE\.attachlink\release-key.xml"
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$script:root = Split-Path -Parent $PSScriptRoot

# ---- SHA-256（.NET 直调，不用 Get-FileHash）----
function Get-Sha256Hex([string]$Path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [IO.File]::OpenRead($Path)
        try {
            ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant()
        } finally {
            $stream.Dispose()
        }
    } finally {
        $sha.Dispose()
    }
}

# ---- 清单签名（与 packaging\setup_update.py 的 manifest_signing_text 逐字节一致）----
function New-ManifestSignature {
    param(
        [Parameter(Mandatory=$true)] [string]$Version,
        [Parameter(Mandatory=$true)] [string]$Url,
        [AllowNull()]                $SizeZip,     # 为空则省行
        [Parameter(Mandatory=$true)] [string]$Sha256Zip,
        [Parameter(Mandatory=$true)] [string]$Sha256Exe,
        [AllowNull()] [string]$NotesUrl,           # 为空则签空值行
        [Parameter(Mandatory=$true)] [string]$KeyXml
    )
    # 被签名正文：固定行序、"`n" 换行（不是 "`r`n"）、行尾也带 "`n"；size_zip 缺失整行省略；
    # published 不参与签名。UTF-8 无 BOM（[Text.UTF8Encoding]::new($false)）
    $lines = @()
    $lines += 'AttachLink manifest v1'
    $lines += "version=$Version"
    $lines += "url=$Url"
    if ($null -ne $SizeZip) { $lines += "size_zip=$SizeZip" }
    $lines += "sha256_zip=$($Sha256Zip.ToLowerInvariant())"
    $lines += "sha256_exe=$($Sha256Exe.ToLowerInvariant())"
    $lines += "notes_url=$(if ($NotesUrl) { $NotesUrl } else { '' })"
    $text = ($lines -join "`n") + "`n"
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($text)

    $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
    try {
        $rsa.FromXmlString($KeyXml)
        $shaOid = '2.16.840.1.101.3.4.2.1'   # SHA-256 的 OID
        try {
            # .NET Framework 4.x 下该重载可用（与 python 侧 pow(s,e,n) 验证互操作）
            $sig = $rsa.SignData($bytes, $shaOid)
        } catch [System.Security.Cryptography.CryptographicException] {
            # 个别运行环境不支持 SignData(字节, OID) 重载时：先算 SHA-256 再 SignHash（等价签名）
            $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
            $sig = $rsa.SignHash($hash, $shaOid)
        }
        return [Convert]::ToBase64String($sig)
    } finally {
        $rsa.Dispose()
    }
}

# ---- 主流程（点源时不执行）----
function Invoke-MakeManifest {
    $srcBas = Join-Path $script:root 'src\AttachLink.bas'
    if (-not (Test-Path $srcBas)) { throw "未找到 $srcBas" }

    # 版本号默认从 .bas 唯一来源读取（与 make-exe.ps1 同一正则）
    if (-not $Version) {
        $basText = [IO.File]::ReadAllText($srcBas, [Text.Encoding]::UTF8)
        if ($basText -notmatch 'ATTACHLINK_VERSION\s+As\s+String\s*=\s*"([^"]+)"') { throw '未能从 AttachLink.bas 提取 ATTACHLINK_VERSION' }
        $Version = $Matches[1]
    }
    if (-not $ReleaseDir) { $ReleaseDir = Join-Path $script:root "release\AttachLink-$Version" }
    Write-Host "版本：v$Version"
    Write-Host "发布目录：$ReleaseDir"

    $exe    = Join-Path $ReleaseDir 'AttachLinkSetup.exe'
    $readme = Join-Path $ReleaseDir 'README.md'
    if (-not (Test-Path $exe))    { throw "未找到 $exe（请先运行 build\make-exe.ps1）" }
    if (-not (Test-Path $readme)) { throw "未找到 $readme（发布目录应含 README.md）" }

    # 发布前版本一致性（v1.5.1 起）：README 的 Release 徽章与配置表里的版本号必须与本次发布一致——
    # 这个 README 会原样进 zip，写错版本就是交付物自相矛盾（一级复核在 v1.5.1 发现过配置表漏改）
    $readmeText = [IO.File]::ReadAllText($readme, [Text.Encoding]::UTF8)
    $badge = [regex]::Match($readmeText, '(?m)^\[!\[Release v([0-9][0-9A-Za-z.\-]*)\]\(https://img\.shields\.io/badge/Release-v([0-9][0-9A-Za-z.\-]*)-[0-9A-Fa-f]+\)\]')
    if (-not $badge.Success) { throw "README.md 顶部找不到 Release 徽章行，无法核对版本号。" }
    if ($badge.Groups[1].Value -ne $Version -or $badge.Groups[2].Value -ne $Version) {
        throw "README.md 顶部的 Release 徽章是 v$($badge.Groups[1].Value)：与本次发布（v$Version）不一致，请先更新 README 再打包。"
    }
    $row = [regex]::Match($readmeText, '(?m)^\|\s*`ATTACHLINK_VERSION`\s*\|\s*`"([^"]+)"`\s*\|')
    if (-not $row.Success) { throw "README.md 配置表里找不到 ATTACHLINK_VERSION 行，无法核对版本号。" }
    if ($row.Groups[1].Value -ne $Version) {
        throw "README.md 配置表的版本号（$($row.Groups[1].Value)）与本次发布（v$Version）不一致，请先更新 README 再打包。"
    }

    # 意外文件检查：*.log 忽略不打包，其它列出警告（不入包）
    $expected = @('AttachLinkSetup.exe', 'README.md', 'latest.json')
    $unexpected = Get-ChildItem $ReleaseDir -File |
        Where-Object { $_.Name -notin $expected -and $_.Extension -ne '.log' }
    foreach ($f in $unexpected) {
        Write-Warning ("发布目录存在意外文件（不打包，请人工确认）：{0}" -f $f.Name)
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    # ---- 打 zip：只含两个文件，条目不带目录前缀 ----
    $zipPath = Join-Path $script:root "release\AttachLink-$Version.zip"
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    $fs = [IO.File]::Open($zipPath, [IO.FileMode]::CreateNew)
    $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($f in @($exe, $readme)) {
            $entryName = [IO.Path]::GetFileName($f)   # 顶层条目，无目录前缀
            # PS 5.1 下 CreateEntryFromFile 是扩展方法，实例方法语法解析不到，须直调静态扩展
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f, $entryName, [IO.Compression.CompressionLevel]::Optimal)
        }
    } finally {
        $zip.Dispose()
        $fs.Dispose()
    }
    # 复核：zip 里必须恰好两个条目
    $check = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        if ($check.Entries.Count -ne 2) { throw "zip 条目数异常（应为 2，实际 $($check.Entries.Count)）" }
        $names = @($check.Entries | ForEach-Object { $_.FullName })
        Write-Host ("zip 条目：{0}" -f ($names -join '、'))
    } finally {
        $check.Dispose()
    }

    $shaZip = Get-Sha256Hex $zipPath
    $shaExe = Get-Sha256Hex $exe
    $sizeZip = (Get-Item $zipPath).Length

    # ---- 签名：私钥只在本脚本（及其点源调用方）内读取，不入仓库、不进安装包 ----
    if (-not (Test-Path $KeyFile)) {
        throw "未找到发布私钥 $KeyFile。先运行 build\new-release-key.ps1 生成密钥，" +
              "并把打印的公钥嵌入 packaging\setup_update.py 后重新构建安装器。"
    }
    $keyXml = [IO.File]::ReadAllText($KeyFile)
    $signature = New-ManifestSignature -Version $Version `
        -Url "https://github.com/jasonisbrave/attachlink/releases/download/v$Version/AttachLink-$Version.zip" `
        -SizeZip $sizeZip -Sha256Zip $shaZip -Sha256Exe $shaExe `
        -NotesUrl "https://github.com/jasonisbrave/attachlink/releases/tag/v$Version" `
        -KeyXml $keyXml

    # ---- 写 latest.json（UTF-8 无 BOM；published 取当天日期；附签名字段）----
    $manifest = [ordered]@{
        version    = $Version
        published  = (Get-Date -Format 'yyyy-MM-dd')
        url        = "https://github.com/jasonisbrave/attachlink/releases/download/v$Version/AttachLink-$Version.zip"
        size_zip   = $sizeZip
        sha256_zip = $shaZip
        sha256_exe = $shaExe
        notes_url  = "https://github.com/jasonisbrave/attachlink/releases/tag/v$Version"
        sig_alg    = 'RS256-SHA256-PKCS1v15:v1'
        signature  = $signature
    }
    $json = ConvertTo-Json -InputObject $manifest
    $manifestPath = Join-Path $ReleaseDir 'latest.json'
    [IO.File]::WriteAllText($manifestPath, $json + "`n", [Text.UTF8Encoding]::new($false))

    # ---- 互操作自检：安装器侧的纯 Python 验签必须接受 PowerShell 签出的签名 ----
    $py = Join-Path $script:root 'packaging\setup_update.py'
    $checkOut = & C:\Python314\python.exe $py --verify-manifest $manifestPath 2>&1
    $checkCode = $LASTEXITCODE
    if ($checkCode -ne 0 -or $checkOut -notcontains 'OK') {
        throw "签名互操作自检失败（退出码 $checkCode）：$checkOut"
    }
    Write-Host "签名自检：$checkOut（安装器纯 Python 验签通过）"

    $shaJson = Get-Sha256Hex $manifestPath
    Write-Host "latest.json 已写出：$manifestPath"
    Write-Host "zip         SHA-256：$shaZip"
    Write-Host "exe         SHA-256：$shaExe"
    Write-Host "latest.json SHA-256：$shaJson"
    # 公钥指纹 = 私钥文件 SHA-256 前 16 位（与 build\new-release-key.ps1 打印的一致，核对备份时逐字比对）。
    # 运行时现算、不硬编码：换钥后这里如果还印着旧指纹，发布 SOP 里「看到指纹才算成功」这一步就会误导人
    $keyFingerprint = (Get-Sha256Hex $KeyFile).Substring(0, 16).ToLower()
    Write-Host "latest.json 已用发布密钥签名（公钥指纹 $keyFingerprint）"
    Write-Host "发布附件：$zipPath 与 $manifestPath（二者都要作为 Release 资产上传）"
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-MakeManifest
}
