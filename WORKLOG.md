# AttachLink 工作日志（交接文档）

> 最后更新：2026-11-19 凌晨（v1.0.0 发布态）
> 状态：**✅ v1.0.0 全量交付就绪——功能（鼎信诺编制/归档迁移/内容基线）+ AttachLinkSetup.exe 安装器 + 交付 SOP；正式包 = `release\AttachLink-1.0.0.zip`（exe + README）；oracle 交付复核 P0 三项已全部落地，待次日同事端实测收尾；后续改进见《开发交接.md》**
> 本文档面向接手者（人或 AI 会话），记录全部根因结论、实证数据、当前卡点与续作步骤。阅读本文件即可继续，无需重新考古。

---

## 一、项目背景

审计底稿无纸化：在 Excel 底稿中插入附件超链接。原方案（非 COM 网页加载项）三缺陷：与鼎信诺 COM 加载项冲突、只能用同级附件文件夹、不支持多选。

**最终选定架构**：`AttachLink.xlam`（VBA，走 Excel 加载项通道）+ **外部看门狗**（PowerShell 常驻，负责在鼎信诺拉起的 Excel 中装载 xlam）。底稿文件保持无宏。

- 主力环境：MS Excel 16.0 64 位（Click-to-Run）+ 鼎信诺审计系统 7100（sjp7.exe，**以管理员运行**）
- 交付/归档形态：Excel 直接交付 → 超链接是最终形态 → 相对路径策略成立

## 二、已实证的根因链（全部本机实测，非推测）

| # | 结论 | 关键证据 |
|---|---|---|
| 1 | **COM 自动化（CreateObject）拉起的 Excel 不加载任何东西**：.xlam（OPEN 注册）、XLSTART、COM 加载项三条自动初始化路径全部跳过 | 干净自动化实例实测：xlam 文件未锁定、XLSTART 副本未加载、COMAddIns 存在（Count=8）但无任何激活迹象；AddIns 集合可枚举且 AttachLink installed=True 但未真正装载 |
| 2 | **高完整性（管理员）进程忽略 HKCU per-user COM 类注册**（微软文档行为） | 用户鼎信诺为管理员运行 → Excel 高完整性：UIPI 拒绝中完整性进程的点击（kimi-cu 点击失败报告 integrity=high）、进程路径读取被拒、GetActiveObject MK_E_UNAVAILABLE |
| 3 | **外部附着可行且已验证**：EXCEL7 窗口 + `AccessibleObjectFromWindow(OBJID_NATIVEOM, IID_IDispatch)` 可从外部取得 Excel Application | 实测发现 3 个实例窗口、取得 Application（Hwnd/Workbooks 正常）、AddIns.Add 装载 Installed=True |
| 4 | **自动化实例 + 打开工作簿后，ROT/GetActiveObject 可用**（同完整性） | in-proc 与跨进程 attach 均 ok；此前 MK_E_UNAVAILABLE 是针对管理员 Excel 的完整性问题 |
| 5 | **零工作簿实例中 `AddIns.Add` 报 0x800A03EC「不能取得类 AddIns 的 Add 属性」**；先 `Workbooks.Add()` 再 Add 即可绕过 | install.ps1 早期踩坑 + 引导器日志再次证实 |
| 6 | `AddIns.Add` 对同名条目幂等（返回既有条目），可安全重复调用 | 多次实测 |
| 7 | xlam 宏信任链正常：AutomationSecurity=msoAutomationSecurityLow + 受信任位置（HKCU `%APPDATA%\Microsoft\AddIns\`）→ 全新自动化实例中 `Workbooks.Open(xlam)` 后 `AttachLink_SelfTest=True` | trustdiag.ps1 实测 |
| 8 | **LB=2 语义**：启动加载标志但未连接；Excel 写 2 后不再自动重试（需手动置回 3） | 多轮观察 |
| 9 | 鼎信诺自己的加载项（ZZZ.DinkumExcel.Connect.V6600）CLSID 的 InprocServer32 为**空**，无法常规 COM 激活 → 鼎信诺的加载项是 sjp7 在拉起 Excel 后**程序式装载**的（我们无法借用其时机） | 注册表解剖 |
| 10 | Excel 启动路径谜团（未解，但已绕开）：LB=3 时自动化实例中引导器对象被创建（cctor 日志）但 OnConnection 从不调用 → LB 写 2。DC1ACE6A/B65AD801 两种 IID、IDispatch/Dual 两种形态、CommandLineSafe=1 全试过，均如此。**唯一成功路径是外部 `Connect=true` / 外部 `AddIns.Add`**（AfterStartup 模式 OnConnection 正常进入） | shim 日志（cctor 有、OnConnection 无）+ 8140 实例手动 Connect 成功日志 |
| 11 | **看门狗 Add-Type「失败」的真根因是文件编码**：watchdog.ps1 曾存为 UTF-8 无 BOM，PS 5.1 按 ANSI(GBK) 误读，C# 注释行 `//（用反射而非 dynamic…` 的乱码字节恰好吞掉行尾换行，把下一行方法签名 `public static object GetApplicationFromWindow(...)` 并入注释行 → csc 连锁编译错误。此前怀疑的 dynamic/守卫真值/缓存全部无关（wdiag.ps1 当时能成功只是因为它没踩同一个字节对齐）。**规则：含中文的 .ps1 必须 UTF-8 with BOM** | 修复 BOM 后 Add-Type 成功；csc 临时文件 o50ivn15.0.cs 第 55 行直接可见被吞的签名 |
| 12 | `AddIns` 集合 `Installed=True` 只反映**登记状态**，不代表已装载（自动化实例可「已登记勾选但从未装载」）。强制装载的可靠手段是 `Workbooks.Open(xlam)`：IsAddin 工作簿无可见窗口、不切换用户活动窗口、已装载时为 no-op。仅靠 Installed 标志位循环重试永远失败（Run Initialize 因宏不存在而抛错，曾被空 catch 吞掉） | 全链路测试：基线 Installed=True 但菜单无 → 看门狗 Open+Run 后菜单在位、SelfTest=True |
| 13 | Open(xlam) 后 EXCEL7 枚举数 +1（隐藏的加载项工作簿也占一个 EXCEL7 文档窗口），无害；$seen 按句柄去重后不再处理 | probe 实测 count 1→2 |
| 14 | **Excel 另存（SaveAs）会把超链接地址重基为相对新位置的 `..\` 路径**：编制阶段插入的 `附件\x`，经鼎信诺归档另存到新目录后变成 `..\..\..\Dinkum\账套\...\附件\x`（仍指回账套原件）——这就是「归档件链接指回原地址」的机制，也意味着只按「绝对路径」识别附件链接会漏掉归档件形态 | 归档迁移夹具实测：以绝对路径 Hyperlinks.Add 后 SaveAs，落盘地址变 `../zht/14370003000/附件/a.pdf` |
| 15 | **「SaveAs(xlam) 报 0x800A03EC / 产物宏不可用」的根因是 Declare 声明位置非法**：VBA 要求模块级 Declare 必须位于第一条过程之前；曾放在模块中段，导致 SaveAs 触发的全项目编译失败，而按需编译的 Run 却正常（VBE 默认 CompileOnDemand 只编译被调过程邻域），极具迷惑性。oracle 会诊定位，单变量实验闭合全部矩阵（声明在前的极简模块全好 / 去声明全好 / 中段声明必坏） | probe_decl / probe_mini_save / 变体二分全矩阵；声明移至顶部声明区后构建自检通过 |
| 16 | **CAPI 三连坑（SHA-256 内容哈希落地过程）**：① .NET COM 路线死路：SHA256Managed 在 .NET 4.x COM 表面无 ComputeHash_2、SHA256CryptoServiceProvider 未注册（探针实测）；② `PROV_RSA_FULL` 不支持 `CALG_SHA_256`，报 NTE_BAD_ALG_ID(0x80090008)，必须 `PROV_RSA_AES(=24)`；③ **VBA 十六进制符号陷阱：`&H800C` ≥ &H8000 被解析为负 Integer 符号扩展成 0xFFFF800C**，必须写 `&H800C&`。另：`⚠` 等 Unicode 符号在 .bas 的 GBK 转码中降级为 `?`（警示符改用 GB2312 自带的 `■`） | hashdiag 仪表日志（LastDllError=-2146893816）；修复后 VBA CAPI 哈希与 PowerShell Get-FileHash 随机二进制对拍 MATCH |
| 17 | Excel COM 竞态两则：① 杀掉曾打开 dist xlam 的 Excel 后，墓碑句柄需数秒释放，立即重建会 SaveAs 静默失败/报错（等 8~10 秒再建）；② 构建/安装后立即跑 COM 测试偶发 0x800AC472（Excel busy），重试即过 | 今晚多次实测 |
| 18 | **封装安装器过程的三连坑（2026-11-19）**：① VBA 标识符大小写不敏感——`Public Function AttachLink_Version` 与 `Public Const ATTACHLINK_VERSION` 重名 → Ambiguous name 编译错误（又一个「按需 Run 正常、SaveAs 全量编译才炸」的迷惑形态）；② 安装器 exe 的子 PowerShell 环境可能无法自动加载 Utility 模块（继承的 PSModulePath 异常）→ `Get-FileHash` 报 CommandNotFoundException——脚本内哈希一律 .NET 直调；③ `Start-Process -ArgumentList` 多元素按空格拼接且不加引号，`-Command` 值含管道/重定向时**静默失效（子进程 exit 0 零副作用）**——提权传递必须用单一参数行 + 内嵌引号 + `-File` 包装器（run-elevated/elevated-capture 重构）；另 install.ps1 卸载分支的 Excel `Quit()` 异步导致文件锁晚释放 → 删文件加重试循环 | exe /deploy-elevated 首版 exit=0 但零效果复现 + 二分定位（T1/T2 提权实验、test-inner 去提权复现）；重构后全链 exit=0 且哈希核对一致 |

> 结论：**放弃进程内 COM 引导器主线，改为「外部看门狗」**（oracle 两次会诊亦建议外部附着路线 D/B）。COM 引导器源码保留于 `src/AttachLinkShim.cs` 仅作归档，不再编译注册。

## 三、当前架构（最终形态）

```
普通双击打开的 Excel ──► xlam 经 OPEN 注册自动装载（原有机制，正常工作）
鼎信诺拉起的 Excel   ──► 看门狗（外部 PowerShell 常驻）检测 EXCEL7 窗口
                          └► AccessibleObjectFromWindow 取 Application
                          └► AddIns.Add(xlam).Installed=true
                          └► Application.Run('AttachLink_Initialize')   ← 显式补初始化（时序已错过）
                          └► 验证右键菜单 Tag 在位
```

- xlam 内 VBA 不变：`AttachLink.bas`（插入/清单/修复/清除/保存归一化）+ ThisWorkbook 的 `Workbook_Open` 兜底初始化 + `SheetBeforeRightClick` 菜单自愈
- 提权链：看门狗以计划任务 RunLevel=Highest 运行 → 可服务管理员权限的鼎信诺 Excel；xlam 放 `C:\Program Files\AttachLink\`（管理员控制目录，规避提权宏劫持面），HKLM 受信任位置指向该目录

## 四、已验证可用的环节（可放心复用）

1. **EXCEL7 枚举与附着**：`EnumWindows`(XLMAIN) → `EnumChildWindows`(EXCEL7) → `AccessibleObjectFromWindow(OBJID_NATIVEOM, IID_IDispatch)` → Window 对象 → 反射取 `.Application`。实测发现全部实例并取得对象。
2. **装载逻辑**：AddIns.Add（幂等）+ Installed=true + 显式 `Application.Run('AttachLink_Initialize')`（幂等）+ 菜单 Tag 验证。手动路径全链路绿（除零工作簿坑已修）。
3. **宏信任**：全新自动化实例中 `Workbooks.Open(xlam)` → `AttachLink_SelfTest=True`。
4. **GetActiveObject 对「已打开工作簿的自动化实例」可用**（同完整性）——已实测 in-proc + 跨进程。
5. install.ps1 / install-elevated.ps1 语法 OK；xlam 构建（编译自检）稳定。

## 五、看门狗卡点（已解决 ✅ 2026-11-18 深夜）

原卡点「Add-Type 抛『无法加载一个或多个类型』」**已定位并修复**，见根因表 #11/#12：

1. **编码修复**：watchdog.ps1 转存 UTF-8 with BOM（此前无 BOM 被 PS5.1 按 GBK 误读，中文注释乱码吞换行，方法签名被并入注释行，csc 连锁报错）。同时在文件头注释中固化编码要求。
2. **错误处理补强**：Add-Type 加 `-ErrorAction Stop`（Continue 下编译错误为非终止错误、不进 catch——这解释了此前「有时 catch 到有时 catch 不到」的混乱史）；catch 内输出完整异常链 + LoaderExceptions；致命失败直接 exit 1（交给计划任务 RestartCount 重启）。
3. **装载逻辑修复**：Provision-App 改为「确保登记（AddIns.Add 幂等 + Installed）→ **Workbooks.Open(xlam) 强制装载**」；Run Initialize 的空 catch 改为记录日志。
4. **验证**：非提权全链路 PASS×2（构建→安装→自动化拉起 Excel→看门狗附着→菜单 Tag 在位→SelfTest=True）。

> 附带教训：watchdog_launcher.vbs 同样不能含非 ASCII（WSH 按 ANSI 读无 BOM 的 vbs，注释乱码吞换行会让第 2 行代码静默失效）——已改纯 ASCII 注释。

## 六、文件清单与状态

| 文件 | 状态 | 说明 |
|---|---|---|
| `src/AttachLink.bas` | ✅ 完成 | 主模块；含 Workbook_Open 兜底初始化（自动化实例 Auto_Open 被跳过的兜底）、SheetBeforeRightClick 菜单自愈、保存归一化 |
| `src/AttachLink.ThisWorkbook.vba` | ✅ 完成 | WithEvents Application 事件桥 |
| `src/AttachLinkShim.cs` | 📦 归档 | COM 引导器（已证不可靠路线），保留不作主线，build 不再编译 |
| `scripts/watchdog.ps1` | ✅ 完成 | **UTF-8 BOM 已修复 + Provision-App 强制装载改造 + 错误处理补强**；全链路验证 PASS |
| `scripts/test-chain.ps1` | ✅ 新增 | 全链路回归测试（复刻鼎信诺启动方式；2026-11-18 从仓库位置跑 PASS） |
| `scripts/test-migrate.ps1` | ✅ 新增 | 归档迁移功能测试（夹具复刻账套→归档场景；2026-11-18 PASS） |
| `scripts/test-hash.ps1` | ✅ 新增 | 内容基线生命周期测试（登记→篡改警示→防洗白→重建→复检，12 断言；2026-11-18 PASS） |
| `packaging/setup_gui.py` + `run-elevated.ps1` + `elevated-capture.ps1` | ✅ 新增 | **AttachLinkSetup.exe 安装器**（PyInstaller 单文件，tkinter 界面：标准安装/鼎信诺管理员部署/两级卸载 + 静默参数 + /selfcheck；业务逻辑走内嵌 ps1 单一事实源） |
| `build/make-exe.ps1` | ✅ 新增 | exe 构建（版本号唯一来源 = bas 的 ATTACHLINK_VERSION；产物 → release\AttachLink-<版本>\{exe, README}） |
| `release/AttachLink-1.0.0/` | ✅ 交付物 | 正式交付目录（exe + README，zip 即发）。验证：/selfcheck 9/9、/install exit=0（OPEN1 注册回写）、/deploy-elevated exit=0（UAC 提权链 + 哈希核对一致） |
| `src/AttachLink.bas` | ✅ 新增功能 | **归档迁移（批量）**：附件夹整体移动 + 归档目录全底稿链接批量改写（处理绝对与 `../` 相对两种指回形态，根因 #14）；移动语义=只删空目录、同名以账套原件覆盖、逐文件失败不中断（2026-11-18 需求定稿：移动而非复制，避免双份存储）。**内容基线（SHA-256）**：清单/体检新增内容基线列（WinCrypt CAPI，根因 #16 三坑修复后与 .NET 对拍 MATCH）；基线存 `附件\_attachlink_baseline.csv` 随附件夹迁移/同步；篡改/同名覆盖红字警示且不洗白基线 |
| `scripts/watchdog_launcher.vbs` | ✅ 完成 | 计划任务隐藏启动器（纯 ASCII，指向 Program Files 路径） |
| `build/build.ps1` | ✅ 完成 | xlam 构建稳定；COM 引导器编译段已删除（2026-11-18） |
| `build/install.ps1` | ✅ 完成 | HKCU 用户级安装；COM 引导器注册段已改为「清理历史残留」（2026-11-18） |
| `build/install-elevated.ps1` | ✅ 完成 | **已改造为机器级看门狗部署**：Program Files 目录 + 登录计划任务（RunLevel=Highest, 当前用户, PT0S 无时限, IgnoreNew, 失败重启×3）+ HKLM 受信任位置 + 旧 COM shim 注册清理 + 装完立即启动；管理员实跑两轮（部署 21:30 / 哈希对齐同步 23:24），含 xlam 与 dist 哈希核对输出 |
| `dist/AttachLink.xlam` | ✅ 最新 | 清理后重新构建，编译自检通过（2026-11-18） |
| `dist/AttachLinkShim.dll` | 🗑 已删除 | 2026-11-18 |
| `README.md` | ✅ 完成 | 共存章节已按看门狗架构重写（根因记录/提权注意事项/安全说明），目录结构与安装方式四已补 |
| `开发交接.md` | ✅ 新增 | 技术实现路径总览 + 痛点改进任务书（P0~P2，面向后续开发，2026-11-18 生成） |
| `AGENTS.md` | ✅ 已登记 | 子项目表 `excel-attach-addin/` 行 |

## 七、用户级需求确认记录（定稿，勿再问）

- 附件文件夹：底稿同级 `附件\`，插入时无则自动建、有则弹窗核对（每文件夹每会话一次，`NOTICE_ONCE_PER_FOLDER` 可改每次）
- 冲突弹窗（覆盖/自动重命名/跳过）；格数不够向下扩展；显示名不含扩展名；清单按 sheet 分组编号（`货币资金-01`）
- 快捷键 Ctrl+Shift+A（插入）/ Ctrl+Shift+K（清单）；右键菜单三入口
- 存量旧链接不管；WPS/网页版不支持；底稿必须已保存
- 鼎信诺归档勾选与否均可：普通超链接不受公式值化影响（README 有五分钟验证流程）

## 八、待办（接手即做）

已完成（2026-11-18 深夜）：~~1 修看门狗~~✅、~~3 清理 COM 残留~~✅、~~4 README 重写~~✅、~~提权部署 install-elevated~~✅（管理员运行成功：Program Files + 计划任务 + HKLM 信任位置；任务看门狗已常驻，日志持续刷新）。

**提权附着终验已过**：以管理员自动化拉起 Excel（鼎信诺等价场景）→ 基线无菜单 → 计划任务看门狗（Highest）40 秒内完成装载 → 菜单在位 + SelfTest=True。证据：`%TEMP%\opencode\elevated-attach-test.log`（RESULT: PASS）。

剩余（唯一）：

1. ~~**鼎信诺真实流程终验**~~ ✅ **2026-11-18 用户实测通过：鼎信诺中可正常使用加载项**（用户另行继续归档与同步测试）
2. oracle 第二轮会诊结果若返回，回填第九节（其针对的「启动路径 OnConnection 之谜」已被根因 #11/#12 的进展部分取代，架构决策已落地）
3. （可选演进，双重复核遗留 O1/O2/O3/O7/O11 等，见 README 复核表）

排查小抄（鼎信诺终验若失败）：
- 看门狗日志 `%TEMP%\AttachLinkWatchdog.log` 是否有「已装载并初始化」；若「附着失败」为完整性问题，确认计划任务是否 Highest（`Export-ScheduledTask -TaskName 'AttachLink Watchdog'` 查 RunLevel）
- 从中完整性 shell 用 Win32_Process 查不到高完整性进程的 CommandLine 属正常（UIPI），勿误判看门狗未运行；以日志是否持续刷新为准
- schtasks 中文字段系统下用 `Export-ScheduledTask` / `Get-ScheduledTask` 查询，勿用英文字段名过滤

## 九、oracle 会诊要点存档

- 第一轮：确认双 IID 尝试方向、提出高完整性忽略 HKCU（已证实）、CommandLineSafe、分阶段 Connect 取 HRESULT（催生了 0x800A03EC 的发现）
- 第二轮（后台待回）：带着「启动路径创建对象但不调 OnConnection + LB→2 + 外部附着已验证」全套证据，询问启动现象解释与架构决策（看门狗 vs 继续 vs 外置工具）
- 关键安全提醒：提权进程勿从用户可写路径装载代码；HKLM CodeBase 指向 Program Files

## 十、本机资产状态快照（2026-11-18 深夜最终态）

- 机器级部署**已完成**：`C:\Program Files\AttachLink\`（AttachLink.xlam + watchdog.ps1 + watchdog_launcher.vbs）✅；计划任务 `AttachLink Watchdog`（登录自启、Highest、当前用户）✅ 常驻中；HKLM\...\16.0\Excel\Security\Trusted Locations\AttachLink → Program Files\AttachLink\ ✅
- 旧 COM shim 注册（HKLM/HKCU/HKCR）已在部署时清理 ✅
- HKCU\...\Excel\Addins\AttachLinkShim.Connect **已删除**（COM 引导器注册清理完毕）
- HKCU\...\Excel\Security\Trusted Locations\AttachLink = `%APPDATA%\Microsoft\AddIns\`（AllowSubFolders=0）✅
- HKCU\...\Excel\Options OPEN1 = "AttachLink.xlam"（交互式 Excel 自动装载用）✅
- `%APPDATA%\Microsoft\AddIns\AttachLink.xlam` = 清理后重新构建安装的最新版 ✅
- dist\AttachLink.xlam = 最新构建 ✅；dist\AttachLinkShim.dll 已删除
- 用户已确认可运行宏；底稿 Excel 交付即归档，无需导回鼎信诺

## 十一、验证记录附录（2026-11-18 深夜）

- 全链路测试 ×3 PASS：COM 自动化拉起 Excel（基线：Installed=True 但无菜单，精确复刻鼎信诺状态）→ 看门狗周期 1 附着 → Workbooks.Open(xlam) 强制装载 → Run Initialize → 菜单 Tag 在位（独立侧验证）→ SelfTest=True（回归脚本：`scripts\test-chain.ps1`，非提权）
- **提权终验 PASS**：管理员自动化拉起 Excel × 计划任务看门狗（Highest）→ 40 秒内装载 → 菜单在位 + SelfTest=True（鼎信诺等价场景）
- 探针实验：EXCEL7 枚举在最小化/还原下稳定；Open(xlam) 后 +1（隐藏加载项工作簿，无害）
- build.ps1（清理后）+ install.ps1（清理后）真实跑通：构建自检通过、安装成功
- install-elevated.ps1（重写后）管理员真实部署成功：任务 Ready→Running、看门狗日志持续刷新（用 Program Files xlam）
- **归档迁移功能**（2026-11-18 深夜追加）：夹具测试 PASS（仓库脚本 `scripts\test-migrate.ps1`）——移动语义（源夹移除/账套根保留/同名覆盖以源为准/未引用文件一并移动/多出旧文件保留）、链接改写（`../` 指回与绝对两种形态 → `附件\文件名`，2 底稿 3 链接）、看门狗回归不受影响（test-chain PASS）
- **内容基线 SHA-256**（2026-11-18 深夜追加）：生命周期 12 断言 PASS（`scripts\test-hash.ps1`）；哈希正确性对拍：VBA WinCrypt CAPI ≡ PowerShell Get-FileHash（7777 字节随机二进制 MATCH）；迁移/看门狗回归 PASS。落地过程即根因 #15/#16（Declare 位置 / CAPI 三坑），oracle 会诊记录见第八节排查小抄上方
- 机器级 Program Files 副本 2026-11-18 23:24 已与 dist/用户级三处哈希对齐（含归档迁移+内容基线终版）
- **用户真实环境确认（2026-11-18 深夜）**：① 本地冲突实验通过——故意替换附件后体检正确红字「■ 内容变更」；② 企业网盘同步后在 **Mac** 上成功同步并查看附件（相对链接跨平台可读的实证；官方口径仍为仅支持 Windows 桌面版，Mac 记录为已观察行为）；③ 同事端体检与故意冲突实验留待次日实测
- **v1.0.0 发布（2026-11-19）**：独立仓库 `jasonisbrave/attachlink`（22 文件首提交 + LICENSE 合并 + About/topics）；GitHub Release `v1.0.0` 附 `AttachLink-1.0.0.zip`（10.9MB）及 zip/exe 双 SHA-256 校验值——正式分发走 Releases，网盘仅作同步载体
- **v1.1.0：附件版本更新（2026-11-19，同事建议）**：基线格式 v2（`路径|哈希|版本|登记时间|更新时间`，自动兼容读 v1）；新菜单「附件版本更新」（有意替换受控流程：旧文件删除 + 基线 v+1 留痕，内容相同拒绝更替，文件操作全部成功才写基线）；体检 F 列带版本号（一致 vN）+ 新增 G 列「更替记录」（v1→v2 + 时间，打印可见的审计留痕）；**未经流程的替换仍红字**——版本更替与篡改区分达成。用户决策记录：旧文件删除不留副本（留痕走清单备注）+ 独立菜单入口。测试：`scripts\test-version.ps1` 11 断言 PASS（交互流程本体待用户手工验证）+ test-hash/test-chain 回归 PASS（F 列断言随版本号显示更新）；三处部署同步 v1.1.0
