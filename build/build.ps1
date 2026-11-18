# =====================================================================
# AttachLink.xlam 一键构建脚本
# 用法: powershell -ExecutionPolicy Bypass -File build.ps1
# 产物: ..\dist\AttachLink.xlam
# 说明:
#   - 临时开启「信任访问 VBA 工程对象模型」(AccessVBOM)：在 try 内进行，
#     任何异常都会走 finally 还原；支持 Office 2010(14.0)~2016+(16.0)
#   - .bas 源码按 UTF-8 读取，自动转为系统 ANSI（GBK）后导入，中文不乱码；
#     ThisWorkbook 代码经 AddFromString 以 COM BSTR(UTF-16) 直传 VBE，无需转码
#   - 构建完成后自动打开产物调用自检函数，编译错误会在这一步暴露
# =====================================================================
param(
    [string]$OutputDir = (Join-Path $PSScriptRoot '..\dist')
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$srcBas = Join-Path $PSScriptRoot '..\src\AttachLink.bas'
$srcTw  = Join-Path $PSScriptRoot '..\src\AttachLink.ThisWorkbook.vba'
if (-not (Test-Path $srcBas)) { throw "找不到源码: $srcBas" }
if (-not (Test-Path $srcTw))  { throw "找不到源码: $srcTw" }

$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$outXlam = Join-Path $OutputDir 'AttachLink.xlam'

# 读取源码（UTF-8）；.bas 需转系统 ANSI——VBA IDE 导入要求 ANSI，中文系统即 GBK
# （历史说明：曾在此编译 COM 引导器 AttachLinkShim.dll，该路线已废弃，源码归档于
#   src\AttachLinkShim.cs，装载职责改由 scripts\watchdog.ps1 外部看门狗承担）
$codeBas = [System.IO.File]::ReadAllText($srcBas, [System.Text.Encoding]::UTF8)
$codeTw  = [System.IO.File]::ReadAllText($srcTw,  [System.Text.Encoding]::UTF8)
$tempBas = Join-Path $env:TEMP 'AttachLink_build.bas'
[System.IO.File]::WriteAllText($tempBas, $codeBas, [System.Text.Encoding]::GetEncoding(0))

$excel = $null
$touched = @()
try {
    # ---- 临时开启 AccessVBOM（记录原值，finally 中还原）----
    foreach ($ver in @('16.0', '15.0', '14.0')) {
        $key = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
        if (Test-Path $key) {
            $saved = $null
            try { $saved = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction Stop).AccessVBOM } catch {}
            Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWORD
            $touched += ,@($key, $saved)
            Write-Host ("AccessVBOM 原始状态[{0}]: {1}" -f $key, $(if ($null -ne $saved) { $saved } else { '未设置' }))
        }
    }
    if ($touched.Count -eq 0) {
        # HKCU 无记录时按已安装的 Office 探测主版本；确实没有 Excel 则直接报错
        $detected = $null
        foreach ($ver in @('16.0', '15.0', '14.0')) {
            if (Test-Path "HKLM:\Software\Microsoft\Office\$ver\Excel") { $detected = $ver; break }
        }
        if (-not $detected) {
            foreach ($ver in @('16.0', '15.0', '14.0')) {
                if (Test-Path "HKLM:\Software\Wow6432Node\Microsoft\Office\$ver\Excel") { $detected = $ver; break }
            }
        }
        if (-not $detected) { throw '未检测到已安装的桌面版 Excel，无法构建。' }
        $key = "HKCU:\Software\Microsoft\Office\$detected\Excel\Security"
        New-Item -Path $key -Force | Out-Null
        Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWORD
        $touched += ,@($key, $null)
        Write-Host ("AccessVBOM 原始状态[{0}]: 未设置（按已装 Office {1} 新建键）" -f $key, $detected)
    }

    # ---- 组装加载项工程 ----
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $vbproj = $wb.VBProject
    [void]$vbproj.VBComponents.Import($tempBas)
    $twComp = $vbproj.VBComponents.Item('ThisWorkbook')
    [void]$twComp.CodeModule.AddFromString($codeTw)

    if (Test-Path $outXlam) { Remove-Item $outXlam -Force }
    $wb.SaveAs($outXlam, 55)   # 55 = xlOpenXMLAddIn（.xlam 格式；18 才是老式 .xla 的 xlAddIn）
    $wb.Close($false)

    # ---- 冒烟自检：打开产物并调用自检函数（触发模块编译，编译错误在此暴露）----
    $wbTest = $excel.Workbooks.Open($outXlam)
    $ok = $excel.Run('AttachLink_SelfTest')
    $wbTest.Close($false)
    if (-not $ok) { throw '自检失败: AttachLink_SelfTest 返回 False' }

    Write-Host "构建完成: $outXlam"
    Write-Host '自检通过: 模块编译正常。'
    Write-Host '下一步安装: powershell -ExecutionPolicy Bypass -File install.ps1'
}
finally {
    if ($null -ne $excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    foreach ($t in $touched) {
        $key = $t[0]; $saved = $t[1]
        try {
            if ($null -eq $saved) {
                Remove-ItemProperty -Path $key -Name AccessVBOM -ErrorAction Stop
                Write-Host "AccessVBOM 已还原为未设置: $key"
            } else {
                Set-ItemProperty -Path $key -Name AccessVBOM -Value $saved -Type DWORD
                Write-Host "AccessVBOM 已还原为 $saved : $key"
            }
        } catch {
            Write-Warning ("AccessVBOM 还原失败 [{0}]: {1}" -f $key, $_.Exception.Message)
        }
    }
    Remove-Item $tempBas -ErrorAction SilentlyContinue
}
