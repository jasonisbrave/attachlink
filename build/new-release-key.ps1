# =====================================================================
# AttachLink 发布签名密钥生成工具（一次性；发布方使用，不随安装包分发）
# ---------------------------------------------------------------------
# 用途：生成「一键更新」版本清单（latest.json）的签名密钥。
#   - 私钥：RSA 参数 XML，保存到 %USERPROFILE%\.attachlink\release-key.xml
#           （目录 ACL 收紧为仅当前用户；不入仓库、不进安装包）。
#   - 公钥：模数（十六进制）+ 指数，打印出来后嵌入 packaging\setup_update.py
#           的 RELEASE_KEY_N / RELEASE_KEY_E，安装器用它验证清单签名。
# 算法：RSA 2048，PKCS#1 v1.5 + SHA-256（PowerShell 侧用 .NET RSACryptoServiceProvider
#       签名，安装器侧用纯 Python 大整数运算验证，无第三方依赖）。
#
# ■ 私钥务必备份到安全的位置（密码管理器 / 加密盘）：
#   私钥丢失 = 之后发布的版本无法通过旧安装器的签名验证，用户只能手动下载过渡；
#   私钥泄露 = 持有者可以签发恶意更新，必须换钥匙并发布带新公钥的版本。
# 已存在私钥时默认不改动（确需重新生成加 -Force，并清楚上述后果）。
# 用法（Windows PowerShell 5.1）：powershell -NoProfile -ExecutionPolicy Bypass -File build\new-release-key.ps1 [-Force]
# 编码：本文件含中文，必须保存为 UTF-8 with BOM。
# =====================================================================
param([switch]$Force)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$keyDir = Join-Path $env:USERPROFILE '.attachlink'
$keyFile = Join-Path $keyDir 'release-key.xml'

if ((Test-Path -LiteralPath $keyFile) -and -not $Force) {
    Write-Host "已存在私钥，未改动：$keyFile"
} else {
    New-Item -ItemType Directory -Force -Path $keyDir | Out-Null
    $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider 2048
    try {
        $xml = $rsa.ToXmlString($true)
        [IO.File]::WriteAllText($keyFile, $xml, [Text.UTF8Encoding]::new($false))
    } finally { $rsa.Clear() }
    # 目录 ACL：去掉继承，只留当前用户完全控制
    $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    & icacls.exe $keyDir /inheritance:r /grant:r "${me}:(OI)(CI)F" | Out-Null
    Write-Host "已生成私钥：$keyFile（仅 $me 可访问）"
}

# 读回并打印公钥
$rsa2 = New-Object System.Security.Cryptography.RSACryptoServiceProvider
try {
    $rsa2.FromXmlString([IO.File]::ReadAllText($keyFile, [Text.Encoding]::UTF8))
    $p = $rsa2.ExportParameters($false)
} finally { $rsa2.Clear() }
$nHex = ([System.BitConverter]::ToString($p.Modulus) -replace '-', '').ToLower()
$nHex = $nHex.TrimStart('0')
$e = [long]0
$exp = [byte[]]$p.Exponent
[array]::Reverse($exp)
foreach ($b in $exp) { $e = $e * 256 + $b }
$fp = (Get-FileHash -LiteralPath $keyFile -Algorithm SHA256).Hash.Substring(0, 16).ToLower()   # 小写：与 build\make-manifest.ps1 现算并打印的指纹逐字一致
Write-Host ""
Write-Host "公钥指纹（私钥文件 SHA-256 前 16 位，便于日后核对备份）：$fp"
Write-Host "把下面两行嵌入 packaging\setup_update.py："
Write-Host ("RELEASE_KEY_N = `"{0}`"" -f $nHex)
Write-Host ("RELEASE_KEY_E = {0}" -f $e)
Write-Host ""
Write-Host "■ 请立即把 $keyFile 备份到安全位置（不要放进仓库或任何同步盘的公开目录）。"
