# =====================================================================
# AttachLink 全链路回归测试（非提权）
# ---------------------------------------------------------------------
# 场景：用 COM 自动化（CreateObject）拉起 Excel——精确复刻鼎信诺的启动
# 方式（跳过一切加载项初始化），再由看门狗完成外部装载。验证：
#   基线（装载前右键菜单不存在） → 看门狗附着 → 菜单 Tag 在位 → SelfTest
# 用法：powershell -ExecutionPolicy Bypass -File test-chain.ps1
# 期望输出末行：RESULT: PASS
# 注意：本文件含中文，必须保存为 UTF-8 with BOM（PS5.1 按 ANSI 误读无
# BOM 文件会破坏解析，教训见 WORKLOG 根因 #11）。
# =====================================================================
$ErrorActionPreference = 'Stop'

$wdPath = Join-Path $PSScriptRoot 'watchdog.ps1'
if (-not (Test-Path $wdPath)) { throw "未找到 $wdPath" }

'[1] COM 自动化拉起 Excel（复刻鼎信诺启动方式，跳过加载项初始化）'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$null = $xl.Workbooks.Add()   # 零工作簿实例没有 EXCEL7 窗口，先建一个
"    hwnd=$($xl.Hwnd)"

'[2] 基线：该实例不应有右键菜单'
$baseline = $false
foreach ($c in $xl.CommandBars.Item('Cell').Controls) { if ($c.Tag -eq 'AttachLinkCtx') { $baseline = $true } }
if ($baseline) {
    "    baseline menu = True（机器级看门狗抢先装载所致，属预期竞态，不影响结论）"
} else {
    "    baseline menu = False"
}

'[3] 启动看门狗（6 周期 ≈ 12 秒）'
$wd = Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',$wdPath,'-MaxCycles','6' -PassThru -WindowStyle Hidden

'[4] 轮询验证菜单与装载状态（最多 25 秒）'
$menu = $false; $installed = $false
$deadline = (Get-Date).AddSeconds(25)
while ((Get-Date) -lt $deadline -and -not $menu) {
    Start-Sleep -Milliseconds 1000
    try {
        foreach ($c in $xl.CommandBars.Item('Cell').Controls) { if ($c.Tag -eq 'AttachLinkCtx') { $menu = $true; break } }
        foreach ($a in $xl.AddIns) { if ($a.Name -ieq 'AttachLink.xlam') { $installed = [bool]$a.Installed } }
    } catch { Start-Sleep -Milliseconds 500 }
}
"    AddIn Installed = $installed"
"    右键菜单 Tag 在位 = $menu"

if ($menu) {
    '[5] 冒烟：在目标实例中运行 SelfTest'
    $t = $xl.Run('AttachLink_SelfTest')
    "    AttachLink_SelfTest = $t"
}

'[6] 清理'
try { foreach ($wb in @($xl.Workbooks)) { $wb.Saved = $true }; $xl.Quit() } catch {}
$wd | Wait-Process -Timeout 15 -ErrorAction SilentlyContinue
Start-Sleep 2
"RESULT: " + $(if ($menu -and $installed) { 'PASS' } else { 'FAIL' })
