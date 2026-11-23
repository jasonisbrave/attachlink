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
| `scripts/test-version.ps1` / `scripts/test-folder.ps1` | ✅ 新增 | 版本更新生命周期（11 断言）/ 文件夹引用+菜单二级化+两轮复核修复回归（25 断言，含外部文件夹/锁定文件混合场景/G 列串行/路径防护），2026-11-20 全 PASS |
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

**维护与迭代触发条件（2026-11-19 用户定）**：
- **常态跟进**：收集同事使用反馈，小版本迭代（发布前须用户确认）
- **大版本迭代触发器：看门狗失效**——症状为鼎信诺打开底稿 2~4 秒后右键菜单不出现。最可能的诱因是 Office 更新改变了加载项启动逻辑（根因 #1 所述三条初始化路径的行为变化）或 EXCEL7/AccessibleObjectFromWindow 附着通道变化。诊断链：`%TEMP%\AttachLinkWatchdog.log`（看门狗存活？EXCEL7 枚举？附着失败原因）→ 鼎信诺 Excel 内查 `AttachLink_GetVersion` → install-elevated 哈希核对两级同版。失效即启动架构级复盘（重走 WORKLOG 第二节的实证方法），不就地打补丁

**流程约定（2026-11-19 用户定，2026-11-21 补充）**：版本发布（GitHub Release 创建/资产更新 + git commit/push）必须先经用户确认，不得自行发布；**双复核（oracle→修复→designer）通过后、发布确认前，用户先在本机实测**（新版已在发布前装机），交互流程等自动化盲区由用户亲验——避免发布后返工。

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
- **v1.1.0：附件版本更新（2026-11-19，同事建议）**：基线格式 v2（`路径|哈希|版本|登记时间|更新时间`，自动兼容读 v1）；新菜单「附件版本更新」（有意替换受控流程：旧文件删除 + 基线 v+1 留痕，内容相同拒绝更替，文件操作全部成功才写基线）；体检 F 列带版本号（一致 vN）+ 新增 G 列「更替记录」（v1→v2 + 时间，打印可见的审计留痕）；**未经流程的替换仍红字**——版本更替与篡改区分达成。用户决策记录：旧文件删除不留副本（留痕走清单备注）+ 独立菜单入口。测试：`scripts\test-version.ps1` 11 断言 PASS + test-hash/test-chain 回归 PASS；三处部署同步
- **v1.1.1：安装器 UX（2026-11-19）**：成功弹窗「安装完成，重启 Excel 生效」；失败弹窗带原因 + **日志快照另存到安装器同目录**（AttachLinkSetup.log，覆盖写，GUI/静默一致，方便同事转发分发人；不自动弹记事本——用户定稿）。发布流程约定同步记档：**版本发布（Release + push）必须先经用户确认**
- **用户真实环境终验（2026-11-19）**：附件版本更新（受控更替 → 一致 v2 绿字 + 更替记录留痕）与安装器成败反馈/日志另存两项功能实测**全绿**——v1.1.1 交互路径（FileDialog/MsgBox 等自动化盲区）闭环确认
- **v1.2.0：文件夹引用/重命名/对话框记忆/菜单二级化（2026-11-20，用户+同事反馈）**：① 文件夹引用=整体复制入附件夹+相对链接（自包含原则，用户拍板），夹内全部文件（含子目录）按相对子路径键登记基线，体检新增「文件夹」行（N 文件一致/■ 内容变更），断链修复支持文件夹整体重连；② 重命名链接显示名（仅改 TextToDisplay，不动文件/地址/基线）；③ 插入对话框会话级记忆上次源目录（不再落附件夹）；④ 右键菜单二级化（5 顶项：插入附件链接/插入文件夹引用/重命名/▸附件清单(体检+断链修复)/▸归档(迁移+版本更新)，Tag 打在 popup 与子项上，看门狗 Menu-Exists 兼容）；⑤ 预览功能**有意不做**（系统级 QuickLook 已覆盖，用户决策）。测试：`scripts\test-folder.ps1` 16 断言 PASS + 四套回归全绿（test-folder 顶项列表顺带留档了完整右键菜单生态快照）；小坑记录：fso.CreateFolder 不能一次建多级（Core 直调需先建父级）；PS5.1 `Where-Object 属性 -eq` 简化语法对 COM 对象绑定失效（测试断言须用脚本块显式 `$_.Caption`）
- **VBA 风险评估（2026-11-20，同事提问 + 外部 AI 会诊 + 本项目代码逐条核实）**：结论——「看门狗让 VBA 权限一直开启」**不成立**（watchdog.ps1 零注册表写入，不碰宏安全级别/AccessVBOM；AccessVBOM 仅构建时 try/finally 临时开启）；常驻两项（HKLM 信任位置指向管理员控制单目录、RunLevel=Highest 计划任务）残余风险低；「无差别装载未签名 VBA 到任意含工作簿的 Excel 会话」机制成立但载荷固定不可注入，属攻击面扩大而非可利用洞。补充量化：被附着会话同时获得全局事件/快捷键（O9 邻域）。改进项入 backlog：T10 合并强化（签名/安装端哈希核对）、新增 T13（看门狗父进程白名单，配置化默认开放，权衡=与鼎信诺耦合）。全文见仓库 Issue #1（已随 1.2.1 处置闭环）
- **oracle v1.2.0 复核与修复（2026-11-20）**：复核裁定「修复后推送」并列 6 项阻塞，全部当日修复并回归——① 断链收集排除文件夹（FileExists 对文件夹恒 False，健康文件夹曾被误列入修复）② CheckFolderBaseline/RegisterFolderFilesBaseline 增加哈希/遍历失败计数，失败显示「校验不可用」绝不显示「一致」，登记侧如实报告部分完成 ③ InsertFolderCore 失败（「失败：」前缀契约）不再插入链接 ④ 路径防护（拒根目录/目标在源内/附件夹内自引用，IsSubPathOf 边界判断）⑤ 外部绝对文件夹链接显示「外部，未校验」不再截断键 ⑥ **G 列串行实锤**——插入文件夹分支时把 `cellG = ""` 重置行顶掉，后行继承前行更替记录（修复=行处理入口统一重置；教训：**逐行状态变量必须在类型分支前重置**）。另修「可删除后重新插入」误导文案（重插不重建基线仍红字）。测试：test-folder 扩至 21 断言（G 串行/路径防护/外部文件夹回归）全 PASS + 四套回归 + **逐入口编译探针**（compile probe，见下）
- **按需编译陷阱再证（WORKLOG #15 变体）**：构建自检（SelfTest）只触发被调过程的按需编译，**过程中的「变量未定义」在别的入口首次调用时才爆**（本次 BuildIndexCore 实际踩中，用户观察到 VBA 编译错误弹窗挂起 COM）。已在测试流程加入逐入口编译探针（对全部公开入口各强制调用一次）；后续新增公开过程必须同步加入探针清单
- **oracle 第二轮复核与修复（2026-11-20）**：3 阻塞+2 同批建议全部当日闭环——① Core 运行时异常返回串统一「失败：」前缀（原「插入文件夹引用失败（未完成）：」被 wrapper 漏判仍插链）② 文件夹体检 fBad 优先于 fFail（篡改红字绝不被「校验不可用」遮蔽，并存时显示「■ 内容变更（含不可用）」双计数）③ SaveBaseline 改 Function 返回成功状态（落盘失败时插入侧返回失败不插链、体检侧汇总警告未持久化）④ 看门狗 seen 键改 hwnd+PID、父进程查询失败不缓存（CIM 临时故障可重试）⑤ 路径防护经 GetAbsolutePathName/GetFolder().Path 规范化（防 ../ 与重复分隔符绕过）⑥ 文档收尾（README 方式五白名单用法/重部署清配置警告/两级同步注意、文件夹基线语义措辞精确化、断言数 25、T13 落地/T14 性能留验登记）。测试：test-folder 扩至 25 断言（外部文件夹「外部，未校验」、锁定文件+篡改混合场景）全 PASS + 四套回归 + 编译探针 ALL OK
- **#19 墓碑进程与安装器误报（2026-11-21）**：用户被「请先关闭 Excel」误报阻塞——实为**墓碑进程**（已退出但被句柄/COM 通道持有的进程对象残影：24KB 内存、无路径、无窗口、状态 Unknown，中/高完整性 taskkill 均报拒绝访问，因为无活体可杀；鼎信诺管理员会话的驻留 Excel 是另一来源）。修复：setup.ps1 的 Excel 检查加 `WorkingSet64 -gt 1MB` 过滤（活体至少几十 MB）。验证：墓碑在场安装通过。修复随 v1.3.0 exe 分发；重启电脑可清墓碑
- **#20 WPS 兼容四层攻坚（2026-11-21，v1.3.0）**：用户 WPS 实测报「编译错误:用户定义类型未定义」（插入文件夹/清单不可用）。逐层定位与修复——①**晚绑定改造**：WPS VBA 缺 Office 共享类型（FileDialog/CommandBar*/CustomProperty 等），31 处宿主类型声明全量改 `As Object` + mso*/xl* 枚举换字面值；清单表归属标记加 sheet 级名称回退。②**哈希进程封锁**：WPS 宏安全拦截子进程创建（WScript.Shell Run/Exec 均 0x80070005，certutil 命令根本不执行）→ 宿主分派：Excel=certutil、WPS=内置纯 VBA SHA-256（Double 中转位运算防 Long 溢出/负数整除方向错）。③**WPS Hex$ 溢出**：纯 VBA 版在 WPS 报 E6——**WPS 的 Hex$ 对 ≥2^31 数值经 Long 转换溢出**（Excel 无此问题），ToHex8 改按字节分段格式化。④验证：纯 VBA 实现与 Get-FileHash 对拍 7/8（空文件为已知边界→「校验不可用」）；**test-wps.ps1（新增，ket.Application COM 自动化）11 断言全 PASS**；性能 64KB/263ms（WPS 端数 MB 附件体检数十秒，一次性成本）。附带：重命名多选批量、bat 入口移除（setup.ps1 命令行为备用通道）、墓碑修复（#19）随本版分发
- **#21 批量重命名排查实录（2026-11-21，v1.3.0）**：用户报「全选超链接区域仍只能单个修改」，且「重启后仍不行」。排查：COM 复现连续/非连续/整表三种选区全部通过（collected=3），代码层无可复现缺陷；装运行时打点请用户真机复现——日志证实 collected=3、三轮 InputBox 依次成功。**根因终版：假重启**——Excel 为单实例应用，关窗口≠退进程，再「打开」Excel 实为连回驻留旧进程（旧代码旧菜单）；彻底退出后一切正常。与 #19 墓碑问题同族（进程驻留），**升级操作要点强化：更新 xlam 后必须确认任务管理器无 EXCEL.EXE 残留再重开**。防御性优化保留：GetScanRange（选区 CountLarge>65536 时与 UsedRange 交叠，防全选亿格迭代冻结）。用户决策记录：WPS 哈希耗时为宿主软件层限制、可接受，README 已注明；**鼎信诺在 WPS 无法使用（基于 Excel COM 自动化），WPS 无需实现归档功能——归档链路仅 Excel，README 已注明范围边界**；鼎信诺直开底稿即见菜单（无需管理员右键），时长可接受

- **#22 第三方 Excel XLL 加载项兼容性排查（2026-11-23）**：确认相关故障与 #1/#12 所述 COM 自动化启动差异有关；补充实测 `RegisterXLL` 可在自动化实例中真实装载系统 XLL（返回 True，测试函数立即可用）。已形成兼容性反馈材料并经私有渠道提交；含第三方实现细节的材料仅本地保存，已按精确文件名加入 `.gitignore`，不进入公开仓库。
- **#23 「删除附件」双层根因修复（2026-11-23，oracle 会诊）**：入口久试不成，根因两层——① 收集阶段 `hl.Worksheet.Name` 调用 Hyperlink 对象不存在的 Worksheet 属性 → 一进入口即 438 进 Fail 弹错误框（此前重构把 links 集合从存对象改为存 "地址|表名" 字符串时引入）；② 删除循环却仍按对象调 `links(i).Range/.Delete`，被 Resume Next 吞掉 → 文件已删、链接未删、样式未清、消息假成功。修复：字符串引用 `Split(...,2)` 解析后经 `wb.Worksheets(表名).Range(地址)` 重定位（表名含 |/空格/单引号无需转义）；抽离 **AttachLink_DeleteAttachmentCore**（无对话框直调层，wrapper 只留守卫+收集+确认+汇总展示，对齐 InsertFolderCore/BuildIndexCore 惯例，可被自动化直调即编译探针）；越界防护（删除只允许当前底稿附件夹内目标，拒附件夹根——IsAttachmentLink 接受外部绝对/`..\` 形态，破坏性操作必须收紧）；基线文件夹后代键连带清理（旧版只删自身键，`文件夹\...` 子键残留）；文件/链接/清理/基线四类失败独立计数、如实汇总（旧版以 links.Count 冒充链接成功数）。**新坑（第九坑）**：Excel 16.0 实测 `Range.Hyperlinks.Delete` 与单个 `Hyperlink.Delete` 均会把单元格边框/填充重置为无格式（LS→-4142、Interior→白）——「保留边框/填充」语义必须用 `Range.ClearHyperlinks`（实测保留），老宿主无此方法时回退 Hyperlinks.Delete（牺牲格式保功能）。测试：新增 `scripts\test-delete.ps1` 24 断言（#25 复核后扩至 35；特殊表名重定位、断链、锁定文件 delFail、保护工作表 linkDelFail、边框/填充保留、基线后代键清理、汇总如实）全 PASS + 六套回归（hash/folder/version/migrate/chain/wps）全 PASS。wrapper 级交互（确认框取消、越界拦截提示）属自动化盲区，留手工冒烟。**版本 bump v1.4.1**（随本修复发布）；`src\AttachLink_noconfig.bas` 经用户确认移出仓库归档（零引用未跟踪的孤立变体，归档于 `$env:TEMP\opencode\attachlink-archive\`，TEMP 清理后即彻底消失）；构建/测试期间 Excel 僵尸实例会致 `0x800AC472`/SaveAs 假锁，批前批后清扫进程可规避（#19/#21 同族）
- **#24 一键卸载整合 + 卸载残留根因（2026-11-23，v1.4.2，同事反馈）**：同事报「卸载后右键菜单仍列示，似乎并未实际卸载」。实测复现（探针法：COM 会话先 Open 工作簿产生 EXCEL7 窗口再等注入——无窗口看门狗不附着）：旧版 `setup.ps1 -Uninstall` 只卸用户级、打印手动提示；机器级看门狗每 2s 扫描并向新会话注入 Program Files 副本重建菜单——用户级卸载后全新会话 `AttachLinkCtx` 菜单仍在，残留实锤（同日用户实测删除功能时被 Program Files 旧副本劫持菜单，同机制二次实证）。**修复=一键整合**：`setup.ps1 -Uninstall` 升级为——用户级卸载 → 检测机器级足迹（计划任务/Program Files/HKLM 信任位置）→ 已提权直跑 / 未提权弹一次 UAC 跑 `install-elevated.ps1 -Uninstall` → 事后按残留事实核验如实报告（不解析提权窗口输出）；新增「Excel 运行中→会话内菜单保留到关闭」提示（#21 同族）。GUI「卸载（用户级）/（机器级）」两按钮合并为「一键卸载（用户级 + 鼎信诺看门狗）」，`/uninstall` CLI 自动获得整合行为，`/uninstall-elevated` 兼容保留；README 增「一键卸载」小节。**验证**：残留复现（卸前菜单 True）→ 一键卸载（一次 UAC）→ 三足迹全清 + 新会话菜单 False → 重装 v1.4.2 全套（APPDATA SHA 3962C9D4793A907F + Program Files + 计划任务）→ 新会话菜单 True + test-chain PASS + `AttachLinkSetup.exe /selfcheck` 全 OK（release\AttachLink-1.4.2 已重建）。小坑记录：① COM 探针无工作簿=无 EXCEL7 窗口，看门狗不附着（探针必须先开簿）；② install.ps1 注册后 `Installed` 回读偶发假 False（Excel 僵尸竞态），清进程重试即过——后续可在脚本内加重试
- **#25 oracle 一级复核修复批（2026-11-23，v1.4.2，推送前双复核第一轮）**：oracle 裁定「修复后复验」4 阻塞 + 范围裁定，全部闭环——① **DeleteAttachmentCore 安全契约**：附件夹不再信任调用方传参（从 wb 自推导 + GetAbsolutePathName 规范化），全量预检（附件夹内/非根/无换行 NUL/引用含 "|"）任一项不合法即**整体拒绝零删除**（首合法后越界同拒），missCnt 改 Core 自算；控制字符检查不用 Asc 逐字符扫（中文 ANSI 转换假阳性，实测附件夹中文名全被误拒）改 InStr 确定性包含。② **LoadBaselineStrict**（status 0/1/2）：删除路径区分「基线不存在」与「读取失败」，读取失败如实 baseFail 且**绝不回写部分字典**（防覆盖丢键）。③ setup.ps1 卸载退出码：残留/UAC 取消 → exit 8（GUI 有对应提示），核验干净才 exit 0。④ test-wps 版本断言 →1.4.2、test-folder 菜单断言补底稿引用/附件删除（总数断言不可靠：CommandBars('Cell') 含原生项，仅存在性断言）。⑤ **底稿引用两缺陷**（用户未提交代码）：同名副本静默复用 → SHA-256 比对一致才复用、不一致弹窗确认；SubAddress 转义不全（只判空格/横线）→ 统一单引号包裹 + 内部单引号双写。⑥ 死代码归档移出（导出交付包按用户 2026-11-23 决策已下线 + GetRelativePath/fso_GetParentFolder 零调用，归档于 %TEMP%\opencode\attachlink-archive\）；版本注释 1.4.0 行实化（INI 外置未做、导出已下线）。⑦ install-elevated 步骤编号 4/5 笔误、install.ps1「旧版本」措辞。⑧ **OPEN/OPENn 残留清理（README E7，源自 PR #2 复核后仓库侧实现）**：install.ps1 卸载分支精确清理指向 AttachLink 的条目（/A 禁用形、库目录相对名、`*\AttachLink\AttachLink.xlam` 后缀形）+ **删除后重排剩余序列连续编号**（OPEN 系列读到缺口即停，断档连坐其后第三方加载项；第三方值含 /R 前缀原样保留）——沙箱 6 断言（含诱饵键 OPENROAM 不被 ^OPEN(\d+)?$ 误伤）+ 真机 E2E（ExcelEnhancer/AuditCat 两第三方保全、AttachLink PF 条目清除、重装后 T6 检测块对 PF 旧版报警如常）全过。文档：README 功能一览加 1d 底稿引用与 8 附件删除（含破坏性三重约束）、菜单表同步、一键卸载小节措辞、E7 ✅；开发交接 T15 置为已实现（含与原提案偏差说明）。测试：test-delete 扩至 **35 断言**全 PASS（新增安全契约 5 + 基线严格性 3）+ 六套回归全 PASS。发布范围决策（用户确认）：**累积发布 1.4.2**（Git 基线 1.3.1，含底稿引用/批量隔离/..\识别等 1.4.0 时代未提交功能一并发布）
- **#26 oracle 二轮复核修复批（2026-11-23，v1.4.2）**：二轮裁定 3 新阻塞全部闭环——① **LoadBaselineStrict 逐行严格化**：非空行必须完整解析为受支持格式（v1=2 字段 / v2=5 字段、非空键/哈希/版本、归一化键不重复），任一不满足判 status=2 读取失败——杜绝「畸形行静默跳过 → 部分字典回写覆盖」；基线失败消息分流（读取失败=「文件未改动，未执行基线清理」/ 写回失败=「请人工核对」，写回失败不保证文件未改动）。② **底稿引用状态防护**：Fail/CleanupTarget/Open 失败三分支 stateCaptured 门控还原（防「捕获前出错」把未初始化 prev* 默认 False 写回全局——关闭事件/提示/刷屏）、Err.Description 先存后清；**AutomationSecurity=3（ForceDisable）真禁宏**（EnableEvents=False 不等于禁宏，README「禁宏」承诺此前不实）+ 宿主不支持则跳过，全部退出路径还原。③ **OPEN 清理原子化**：删除/写入全 -ErrorAction Stop、仅真实删除计数、重排前快照（原名原值）失败按快照回滚、终态核验（数量/连续性/第三方值逐项一致），失败 exit 9（GUI 退出码文档补 8/9）；setup.ps1 失败路径不再先报「卸载完成」再报失败。测试：test-delete 扩至 **39 断言**（新增畸形基线不回写场景 4 断言）全 PASS + 六套回归全 PASS + OPEN 沙箱（原子化版）终态核验 + 真机 E2E（卸载→第三方保全→重装往返）+ BOM/GBK/py_compile 闸门 + exe 重建 selfcheck；三处部署同步（117100 字节同构）。非阻塞留档：%TEMP% 归档为非持久位置（导出交付包/noconfig/辅助函数的唯一副本），用户知情——如需长期保留应移至持久私有目录。**三轮 2 项闭环**：① 底稿引用 AutomationSecurity 读取失败的旧错误在 Open 前无条件 Err.Clear（防泄入打开成功判定造成误报）+ Open 失败分支兜底关闭半开 targetWb；② OPEN 清理单项删除失败置失败标记、新增终态「无 AttachLink 残留」独立核验（覆盖全部删除失败且 removed=0 仍 exit 0 的漏洞）、exit 9 后移至其余清理全部执行之后（承诺属实）、setup.ps1 原样传播子脚本退出码 9（不再被 throw 吞成 1，GUI 退出码承诺成立）。三轮后重建：test-delete 39 断言 + 七套回归全 PASS，真机 E2E 卸载/重装均 exit 0、第三方条目保全，三处部署 117458 字节同构，exe 重建 selfcheck OK
