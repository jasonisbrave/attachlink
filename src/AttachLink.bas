Attribute VB_Name = "AttachLink"
'=====================================================================
' AttachLink —— 审计底稿附件链接加载项（Excel 加载项 / .xlam）
'---------------------------------------------------------------------
' 功能：
'   1. 插入附件链接：前置检查附件文件夹（无则自动创建；已有则提示核对）
'      -> 多选文件 -> 自动归档复制到底稿同级「附件」文件夹
'      -> 以相对路径在选中单元格插入超链接（底稿文件夹整体搬移不断链）
'      -> 对话框自动定位到上次选取附件的源目录（会话级记忆）
'   1b. 插入文件夹引用：选定文件夹整体复制入附件夹 + 相对链接（点击即
'      在资源管理器打开）；夹内全部文件自动纳入内容基线（可检测篡改）
'   1c. 重命名链接显示名：只改超链接显示文本（应对收到的文件命名不规
'      范），不改动文件本身与链接地址
'   2. 附件清单/体检：按底稿 sheet 分组编号（sheet名-01），
'      检查每个附件是否存在（断链红字）+ SHA-256 内容基线校验
'      （基线存附件夹内 _attachlink_baseline.csv，随附件夹迁移/同步；
'        内容被替换/篡改/同步同名覆盖时红字「■ 内容变更」持续警示）
'   3. 断链修复：指定附件文件夹当前位置，按文件名自动重连
'   4. 清除链接：清除选区内的附件链接并恢复单元格字体格式
'   5. 保存保护：保存时自动把被 Excel 改写为绝对路径的附件链接
'      归一化回相对路径
'   6. 归档迁移（批量）：鼎信诺归档是把底稿另存到新目录、附件夹不在其
'      处理范围。在归档目录打开任一底稿后运行：把账套附件文件夹整体
'      【移动】到归档目录同级（不留副本；目标同名以账套原件覆盖），
'      并把归档目录全部底稿的附件链接改写为相对地址（处理绝对路径
'      与 ../ 指回两种形态），使归档件自包含
'   7. 右键菜单（单元格右键）：插入附件链接（可多选）/ 插入文件夹引用 /
'      插入底稿引用 / 归档迁移（批量）+ 「附件管理」子菜单（附件清单 / 体检、
'      附件断链修复、附件版本更新、重命名链接显示名、附件删除）
' 设计约定：
'   - 与鼎信诺 COM 加载项共存（走「Excel 加载项」通道，不注册 COM）
'   - 底稿文件保持无宏（宏全部在本加载项内，不污染底稿）
'   - 使用普通超链接而非 HYPERLINK 公式（避免鼎信诺归档公式值化失效）
'   - WPS 兼容（1.3.0）：宿主对象一律晚绑定（As Object）+ 枚举用字面值——
'     WPS VBA 缺 Office 共享类型（FileDialog/CommandBar* 等）与 mso*/xl* 常量，
'     早绑定会触发「用户定义类型未定义」的模块编译错误；清单表归属标记在
'     CustomProperties 之外增加 sheet 级名称回退
' 构建与安装：见 build\build.ps1 与 build\install.ps1
'=====================================================================
Option Explicit

'============================ 配置区 =================================
' 可配置键：ATTACH_FOLDER_NAME / SHOW_EXTENSION / AUTO_EXTEND_CELLS /
'           NOTICE_ONCE_PER_FOLDER / HOTKEY_INSERT / HOTKEY_INDEX / INDEX_SHEET_NAME

' 附件文件夹名（位于底稿同级目录，可改名）
Public Const ATTACH_FOLDER_NAME As String = "附件"
' True = 链接显示文本含扩展名；False = 不含扩展名
Public Const SHOW_EXTENSION As Boolean = False
' True = 选中单元格不足时自动向下扩展（跳过非空格与合并格）
Public Const AUTO_EXTEND_CELLS As Boolean = True
' True = 「附件文件夹已存在」提示对每个「插入入口 × 文件夹」每次会话仅弹一次
' （三个入口相互独立）；False = 每次插入都提示
Public Const NOTICE_ONCE_PER_FOLDER As Boolean = True
' 快捷键：插入附件链接 / 附件清单（改为 "" 可禁用）
Public Const HOTKEY_INSERT As String = "^+a"    ' Ctrl+Shift+A
Public Const HOTKEY_INDEX As String = "^+k"     ' Ctrl+Shift+K
' 附件清单工作表名
Public Const INDEX_SHEET_NAME As String = "附件清单"

' 版本号（单一事实源：安装器/看门狗日志/支持排障都以此为准；语义化递增）
' 1.6.0：稳定运行优先版（不扩功能面）——① SHA-256 改 advapi32 CAPI 原生实现优先（条件编译声明：
'       Excel 64 位走 PtrSafe+LongPtr，WPS 走 #Else 分支；分块流式读取；首次使用做标准向量自检，失败则
'       本会话回退），修复 WPS 下使用变慢（纯 VBA 实测 0.21 MB/s：插入 5MB 约 23 秒、体检 50MB 约
'       4 分钟；CAPI 约 300 MB/s），certutil / 纯 VBA 仅作回退；② 底稿引用入口先做附件文件夹前置检查
'       （选目标底稿之前），「文件夹已存在」提示改为每个插入入口对同一文件夹每次会话各提示一次；
'       ③ 右键菜单重排：归档迁移提级到一级，「归档」子菜单更名「附件管理」，重命名链接显示名移入其中。
'       看门狗（watchdog.ps1）同版改为只装载不登记 + 在位判定加固（菜单 Tag 与 Workbooks 三态联合确认，
'       死菜单按未装载处理）。安装器修复一键更新后日志快照写错目录
' 1.5.1：安装器修复——GitHub 直连取更新清单被「最终地址白名单」拦下（GitHub 把 Release 资产 302 到
'       release-assets.githubusercontent.com）；现对本仓库的 Release 资产地址与内置加速镜像放行该跳转。
'       本版 VBA 功能无变化，仅版本号
' 1.5.0：附件身份与可信重连（插入即登记 SHA-256；同名同内容复用、异内容自动另存「名(n).ext」；
'       断链修复按内容哈希分层重连，候选不唯一不猜）/ 基线安全读写（先写临时文件回读校验、备份后替换，
'       损坏或未来格式不覆盖）/ 体检增强（同次体检同一文件只算一次哈希、状态栏进度、Esc 取消）/
'       保存归一化失败可见（状态栏提醒 + 体检红字「正常（■ 绝对路径）」）。安装器另含健康检查与
'       检查更新 / 一键更新（见 README）。看门狗脚本本版未改动
' 1.4.4：修复「卸载后重新安装」失败（提示「加载项注册后未处于启用状态」）——Excel 的加载项列表
'       按文件名识别条目，机器级看门狗登记过 Program Files 副本的机器上，卸载留下的悬空登记会让
'       重新安装绑定到它；现安装前自动清理并核对绑定的是库目录副本，卸载时一并清理 Add-in Manager
'       登记。仅安装脚本变更，本版 VBA 功能无变化，仅版本号
' 1.4.3：与 WPS 共存修复——看门狗只附着 EXCEL.EXE（WPS 表格复用 XLMAIN/EXCEL7 类名，
'       旧版会把 xlam 注入 WPS，致其卡在「正在打开文档」或反复弹「同名文件已打开」）；
'       安装器宿主识别（Excel.Application 被 WPS 接管时改走 CLSID）；看门狗日志改心跳、
'       xlam 缺失时暂停附着。本版 VBA 功能无变化，仅版本号
' 1.4.2：一键卸载整合（用户级 + 机器级看门狗）+ OPEN/OPENn 残留清理与序列重排 + 复核修复批
' 1.4.1：删除附件入口双层根因修复（AttachLink_DeleteAttachmentCore）
' 1.4.0：附件删除（删链接+删文件+基线清理）/ 底稿引用（跨底稿交叉索引）/
'       IsAttachmentLink 识别 ..\ 形态 / 批量操作逐文件故障隔离
'       （导出交付包曾实现后于 2026-11-23 按用户决策下线（代码归档移出，回退任务 T1）；
'        配置外置 INI 未实施，配置仍为源码常量）
' 1.3.1：右键菜单再浓缩（4 顶项：归档子菜单收编体检/断链修复，归档前后步骤统一入口）
' 1.3.0：WPS 表格支持（全量晚绑定改造，需宿主启用 VBA 环境）+ 重命名多选批量 + 安装器墓碑进程误报修复
' 1.2.1：风险处置（Issue #1）——看门狗父进程白名单（配置化默认开放）+ 装载版本日志
' 1.2.0：文件夹引用（复制入附件夹+夹内基线）/ 重命名链接显示名 / 插入对话框记忆上次源目录 / 右键菜单二级化
' ■ 发布流程约定：版本发布（GitHub Release + git push）必须先经用户确认
Public Const ATTACHLINK_VERSION As String = "1.6.0"
'=====================================================================

Private Const MENU_TAG As String = "AttachLinkCtx"
' 内容基线文件名（位于附件文件夹内，随附件夹整体迁移/同步；UTF-8，格式：相对子路径|SHA-256）
Private Const BASELINE_FILE As String = "_attachlink_baseline.csv"

' ■ VBA 位置规则：模块级 Declare 必须放在【第一条过程之前】（本声明区）。
' 曾放在模块中段导致全量编译失败——SaveAs(xlam) 触发全项目编译时报
' 0x800A03EC 或产物宏不可用，而按需编译的 Run 却正常，极具迷惑性
' （见 WORKLOG #15）。新增任何 Declare/模块级变量一律放这里。
' SHA-256 经 WinCrypt CAPI（advapi32）：系统原生组件，无 COM/.NET 依赖
' （.NET COM 路线已实测排除：SHA256Managed 在 .NET 4.x 无 ComputeHash_2、
'  SHA256CryptoServiceProvider 未注册）。
'（CAPI Declare 块已于 1.3.0 移除：WPS VBA 缺 LongPtr/PtrSafe 生态，哈希改经 certutil，见 Sha256File 注释）
' CAPI Declare 块于 v1.6.0 以条件编译形态恢复（R4）：#If Win64 / #Else 双分支，
' Excel 64 位（PtrSafe+LongPtr）与 WPS（VBA6 风格 #Else 分支）均编译运行通过
' （1.3.0 当年失败根因是无条件使用 LongPtr，WPS 报「用户定义类型未定义」）
#If Win64 Then
Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextW" (ByRef phProv As LongPtr, ByVal pszContainer As LongPtr, ByVal pszProvider As LongPtr, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" (ByVal hHash As LongPtr, ByVal pbData As LongPtr, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
#Else
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextW" (ByRef phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As Long, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#End If
' CAPI 会话句柄（模块级保存，避免在多个函数签名里重复条件类型；0 = 未占用）
#If Win64 Then
Private capiHProv As LongPtr
Private capiHHash As LongPtr
#Else
Private capiHProv As Long
Private capiHHash As Long
#End If
' CAPI 可用性会话状态：0 未自检 / 1 可用 / -1 不可用（自检失败后本会话不再重试）
Private capiState As Long
' 已提示过「附件文件夹已存在」的（入口 × 文件夹）记录（会话级去重，避免重复
' 打扰：键 = 入口标识 & "|" & 小写附件夹路径；NOTICE_ONCE_PER_FOLDER=False 时
' 记录不参与判断，每次必提示）
Private notifiedAttachDirs As Object
' 上次选取附件/文件夹的源目录（会话级记忆：插入对话框优先定位到这里，不再每次落在附件夹）
Private lastSourceDir As String
' 基线最近一次解析/写入失败原因（SaveBaseline/LoadBaselineEx 写入，调用方带进提示）
Private lastBaselineError As String
' 单次 Core 运行内的哈希缓存（按完整路径小写去重；各 Core 入口开始时重置，T14 去重基础）
Private hashCache As Object
' Sha256File 真实计算次数（缓存命中不计；诊断用，经 AttachLink_GetHashCount 读取）
Private hashComputeCount As Long
' 进度助手真实刷新次数（诊断用，经 AttachLink_GetProgressCount 读取）
Private progressCount As Long
' 保存归一化失败记录（T8）：键=工作簿 FullName（无路径时 Name），值="N|明细"
Private normFail As Object
' 测试钩子（仅测试脚本使用；生产路径恒为 0）：>0 时接下来 N 次「改指链接地址」
' 不真正赋值、按失败处理（每次消耗 1，经 AttachLink_TestRepointFail 设置）。
' 本机 Excel 实测：受保护工作表上 Hyperlink.Address 赋值并不抛错（与
' Hyperlinks.Add 不同），失败可见性链路只能经此钩子在测试中端到端演练。
' 安全性说明（N-03）：仅测试使用；同进程任何宏本来就能直接调用各 Core，
' 本钩子不构成额外攻击面。
Private testRepointFail As Long
' 重入防护（O-06）：任一 Core 运行期间再次进入任何 Core → 直接返回
' 「失败：已有操作正在进行」；所有出口（含取消/失败）经公开包装复位。
' Excel 宏宿主单线程执行，布尔标志无需原子操作。
Private coreBusy As Boolean
' 测试钩子（仅测试使用；同进程宏可直接调用各 Core，本钩子不构成额外攻击面）：
' InsertCore 进入后若发现为 True，先直调一次 InsertCrossRefCore（必被重入防护
' 拦截）并把结果记入 testReentrySeen，随即清标志。生产路径恒为 False。
Private testReentryProbe As Boolean
Private testReentrySeen As String
' 测试钩子（仅测试使用；同进程宏可直接调用各 Core，本钩子不构成额外攻击面）：
' >0 时 CheckFolderBaseline 第 N 次进入时按用户取消处理（Err.Raise 18，
' 经 O-06 的递归透传链路回到 BuildIndexCore 的取消分支）。生产路径恒为 0。
Private testCancelInFolder As Long
' 测试钩子（仅测试使用）：>0 时 InsertCoreInner 开头抛未处理错误，验证
' O-06/V2 包装器出错兜底（忙碌标志复位 + 「失败：<过程名> - <原因>」）。生产恒 0。
Private testThrowAtInsert As Long
' 测试钩子（仅测试使用）：>0 时 ArchiveFileEx 首次复制完成后向源文件追加字节
' （源大小变化），验证 V11b 复制后校验以目标实际哈希登记。生产恒 0。
Private testGrowAfterCopy As Long

' 归档判定类别（ArchiveFileEx 的 category 输出）
Private Const CAT_NEW As Long = 1       ' 目标不存在，已复制并登记
Private Const CAT_INSITU As Long = 2    ' 源文件已在附件夹内（不复制）
Private Const CAT_REUSED As Long = 3    ' 同名同内容，静默复用
Private Const CAT_RENAMED As Long = 4   ' 同名异内容，自动另存 主名(n).ext
Private Const CAT_POPUP As Long = 5     ' 哈希不可用，退回弹窗确认路径
Private Const CAT_ANOMALY As Long = 6   ' 副本与既有基线不一致（疑似被篡改），未复用、已另存（O-05）

' O-05：三方哈希一致性闸——夹内已有副本要被「复用/仅改指」扩散新链接前，必须
' copy==baseline==source。基线已登记且副本哈希与基线不符（副本疑似被篡改，体检会
' 标红）时返回 True：调用方不得复用，必须另存新名（新副本自带一致基线）。
Private Function BaselineMismatch(ByVal base As Object, ByVal key As String, ByVal hCopy As String) As Boolean
    If Len(hCopy) = 0 Then Exit Function
    If base.Exists(key) Then
        If StrComp(hCopy, Split(base(key), "|")(0), vbTextCompare) <> 0 Then BaselineMismatch = True
    End If
End Function

' O-05：副本与基线不一致时的安全另存——沿用「主名(n).ext」不覆盖任何既有文件，
' 新副本登记后自身基线三方一致。knownHash 传副本预期哈希（ByRef 回传目标实际
' 哈希，V11b 复制后校验）。返回另存路径；失败返回 ""。
Private Function SaveAsAnomalyCopy(ByVal fso As Object, ByVal srcPath As String, _
        ByVal attachDir As String, ByVal fName As String, ByRef category As Long, _
        ByRef detail As String, ByRef knownHash As String) As String
    Dim baseName As String, extName As String, n As Long, cand As String
    Dim szA As Variant, mtA As Variant
    baseName = fso.GetBaseName(fName)
    extName = fso.GetExtensionName(fName)
    n = 2
    Do
        cand = attachDir & "\" & baseName & "(" & n & ")" & IIf(Len(extName) > 0, "." & extName, "")
        If Not fso.FileExists(cand) Then
            SnapshotSrcFile fso, srcPath, szA, mtA
            fso.CopyFile srcPath, cand, False
            knownHash = VerifyCopiedHash(fso, srcPath, cand, szA, mtA, knownHash)
            If Len(knownHash) = 0 Then knownHash = Sha256File(cand)
            category = CAT_ANOMALY
            detail = fName & "（附件夹内副本与其内容基线不一致，体检会标红），本次未复用，已另存为 " & fso.GetFileName(cand)
            SaveAsAnomalyCopy = cand
            Exit Function
        End If
        n = n + 1
    Loop While n < 1000
    category = 0
    SaveAsAnomalyCopy = ""
End Function

'============================ 注册与注销 ==============================
Public Sub Auto_Open()
    ' 注意：仅交互式启动 Excel 时才会执行；鼎信诺等软件通过 COM 自动化拉起 Excel
    ' 时会跳过 Auto_Open，初始化由 ThisWorkbook.Workbook_Open 兜底执行
    AttachLink_Initialize
End Sub

' 初始化入口：注册右键菜单与快捷键（幂等，可重复调用）
Public Sub AttachLink_Initialize()
    RegisterMenu True
    If Len(HOTKEY_INSERT) > 0 Then Application.OnKey HOTKEY_INSERT, "AttachLink_Insert"
    If Len(HOTKEY_INDEX) > 0 Then Application.OnKey HOTKEY_INDEX, "AttachLink_BuildIndex"
End Sub

' 右键菜单自愈：菜单被其他加载项（如鼎信诺）重置/清除时，下次右键前自动补回
Public Sub AttachLink_EnsureMenu()
    On Error Resume Next
    Dim cb As Object, c As Object, found As Boolean
    Set cb = Application.CommandBars("Cell")
    If cb Is Nothing Then Exit Sub
    For Each c In cb.Controls
        If c.Tag = MENU_TAG Then found = True: Exit For
    Next c
    If Not found Then RegisterMenu True
End Sub

Public Sub Auto_Close()
    RegisterMenu False
    If Len(HOTKEY_INSERT) > 0 Then Application.OnKey HOTKEY_INSERT
    If Len(HOTKEY_INDEX) > 0 Then Application.OnKey HOTKEY_INDEX
End Sub

Private Sub RegisterMenu(ByVal adding As Boolean)
    Dim cb As Object, ctl As Object, btn As Object
    On Error Resume Next
    Set cb = Application.CommandBars("Cell")
    If cb Is Nothing Then Exit Sub
    For Each ctl In cb.Controls
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next ctl
    If adding Then
        Set btn = cb.Controls.Add(1, , , 1, True)
        btn.Caption = "插入附件链接（可多选）"
        btn.OnAction = "AttachLink_Insert"
        btn.Tag = MENU_TAG
        btn.BeginGroup = True

        Set btn = cb.Controls.Add(1, , , 2, True)
        btn.Caption = "插入文件夹引用"
        btn.OnAction = "AttachLink_InsertFolder"
        btn.Tag = MENU_TAG

        Set btn = cb.Controls.Add(1, , , 3, True)
        btn.Caption = "插入底稿引用"
        btn.OnAction = "AttachLink_InsertCrossRef"
        btn.Tag = MENU_TAG

        Set btn = cb.Controls.Add(1, , , 4, True)
        btn.Caption = "归档迁移（批量）"
        btn.OnAction = "AttachLink_ArchiveMigrate"
        btn.Tag = MENU_TAG

        ' 二级菜单：附件管理（清单/体检 / 断链修复 / 版本更新 / 重命名 / 删除——
        ' v1.6.0 R6 重排：归档迁移提级一级，重命名链接显示名移入子菜单）
        Dim pop As Object, subBtn As Object
        Set pop = cb.Controls.Add(10, , , 5, True)
        pop.Caption = "附件管理"
        pop.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 1, True)
        subBtn.Caption = "附件清单 / 体检"
        subBtn.OnAction = "AttachLink_BuildIndex"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 2, True)
        subBtn.Caption = "附件断链修复"
        subBtn.OnAction = "AttachLink_Repair"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 3, True)
        subBtn.Caption = "附件版本更新"
        subBtn.OnAction = "AttachLink_UpdateVersion"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 4, True)
        subBtn.Caption = "重命名链接显示名"
        subBtn.OnAction = "AttachLink_RenameLink"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 5, True)
        subBtn.Caption = "附件删除（删链接+文件）"
        subBtn.OnAction = "AttachLink_DeleteAttachment"
        subBtn.Tag = MENU_TAG
        ' 导出交付包已下线（2026-11-23 用户决策：功能鸡肋，回退任务清单 T1 待改进）
    End If
End Sub

'============================ 1. 插入附件链接 =========================
Public Sub AttachLink_Insert()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then
        MsgBox "当前底稿尚未保存，无法确定「" & ATTACH_FOLDER_NAME & "」文件夹位置。" & vbCrLf & _
               "请先保存底稿（Ctrl+S）后再插入附件链接。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要放置附件链接的单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If
    If Selection.Cells.CountLarge > 10000 Then
        MsgBox "选区单元格过多，请选中具体的单元格区域后再插入。", vbInformation, "AttachLink"
        Exit Sub
    End If
    If InStr(1, wb.Path, "://", vbTextCompare) > 0 Then
        MsgBox "底稿保存在网页路径（SharePoint/OneDrive 网页版），相对路径附件链接不可用。" & vbCrLf & _
               "请先把底稿另存到本地磁盘或局域网共享盘（\\服务器\共享文件夹）后再插入。", vbExclamation, "AttachLink"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = EnsureAttachDirPrompted(fso, wb, "insert", "本次插入的附件将归档复制到该文件夹")
    If Len(attachDir) = 0 Then Exit Sub

    Dim fd As Object
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "选择要插入的附件（Ctrl / Shift 可多选）"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "所有文件", "*.*"
        ' 优先定位到上次选取附件的源目录（会话级记忆），避免每次翻找
        If Len(lastSourceDir) > 0 Then
            .InitialFileName = lastSourceDir & "\"
        Else
            .InitialFileName = attachDir & "\"
        End If
        If .Show <> -1 Then Exit Sub
    End With
    lastSourceDir = fso.GetParentFolderName(fd.SelectedItems(1))

    ' 拼接源路径列表（Windows 文件名不含 |，可作分隔符），交给 Core 执行
    Dim srcPaths As String, i As Long
    For i = 1 To fd.SelectedItems.Count
        If i > 1 Then srcPaths = srcPaths & "|"
        srcPaths = srcPaths & fd.SelectedItems(i)
    Next i

    Dim summary As String
    summary = AttachLink_InsertCore(srcPaths, Selection)
    If Left$(summary, 3) = "失败：" Then
        MsgBox summary, vbCritical, "AttachLink"
    Else
        MsgBox summary, IIf(InStr(summary, "失败 ") > 0, vbExclamation, vbInformation), "AttachLink"
    End If
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "插入附件链接时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 插入核心（无对话框，供自动化直调；UI 包装负责前置提示与文件选择）。
' srcPaths = 以 | 连接的绝对路径列表；targetRange = 起始选区（一格一链，
' 不足自动向下扩展）。归档判定执行 T2 附件身份规则（ArchiveFileEx）：
' 同名同内容静默复用、同名异内容自动另存 主名(n).ext（不覆盖不弹窗）、
' 插入即登记内容基线（整批一次落盘）。返回汇总文本；「失败：」前缀=未执行。
' O-06/V2 重入防护包装：忙碌时直接失败返回；出错兜底——任何出口（含 Inner
' 抛未处理错误）都复位忙碌标志，错误如实转「失败：<过程名> - <原因>」
Public Function AttachLink_InsertCore(ByVal srcPaths As String, ByVal targetRange As Object) As String
    If coreBusy Then
        AttachLink_InsertCore = "失败：已有操作正在进行。"
        Exit Function
    End If
    coreBusy = True
    Err.Clear
    On Error Resume Next
    AttachLink_InsertCore = InsertCoreInner(srcPaths, targetRange)
    If Err.Number <> 0 Then
        AttachLink_InsertCore = "失败：AttachLink_InsertCore - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    coreBusy = False
End Function

Private Function InsertCoreInner(ByVal srcPaths As String, ByVal targetRange As Object) As String
    ' V2 测试钩子：注入未处理错误（在挂 On Error 之前 Raise，错误才能穿出 Inner、
    ' 验证包装器兜底复位忙碌标志）。生产恒 0。
    If testThrowAtInsert > 0 Then
        testThrowAtInsert = 0
        Err.Raise vbObjectError + 590, , "测试注入的未处理错误"
    End If
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        InsertCoreInner = "失败：没有活动工作簿。"
        Exit Function
    End If
    If Len(wb.Path) = 0 Then
        InsertCoreInner = "失败：活动工作簿尚未保存。"
        Exit Function
    End If
    If targetRange Is Nothing Then
        InsertCoreInner = "失败：未提供目标单元格。"
        Exit Function
    End If
    ' N-05：目标单元格必须属于当前底稿（防自动化直调时把链接插到别的工作簿）
    Dim ownerName As String
    On Error Resume Next
    ownerName = targetRange.Worksheet.Parent.FullName
    On Error GoTo Fail
    If Len(ownerName) = 0 Or StrComp(ownerName, wb.FullName, vbTextCompare) <> 0 Then
        InsertCoreInner = "失败：目标单元格不属于当前底稿。"
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME
    If Not EnsureFolder(fso, attachDir) Then
        InsertCoreInner = "失败：无法创建附件文件夹：" & attachDir
        Exit Function
    End If
    ResetHashCache
    Dim baseStatus As Long, baseRO As Boolean
    Dim base As Object
    Set base = LoadBaselineEx(fso, attachDir, baseStatus, baseRO)
    ' O-04：基线损坏 = 失败关闭——在任何文件写入之前返回
    If baseStatus = 2 Then
        InsertCoreInner = "失败：基线文件损坏或无法读取（" & lastBaselineError & _
                          "），为避免绕过附件身份校验，本次未插入任何附件。请先人工核对或重建基线。"
        Exit Function
    End If
    ' O-06 重入演练（仅测试钩子，生产恒 False）：Core 运行期间再调另一 Core，
    ' 必须被上面的重入防护拦截；结果附进汇总供测试断言
    Dim reentryNote As String
    If testReentryProbe Then
        testReentryProbe = False
        testReentrySeen = AttachLink_InsertCrossRefCore("", "Sheet1", Nothing)
        ' V3：Core 运行期间版本更新必须同样被忙碌互斥拦截（UpdateVersionCore 已纳入
        ' coreBusy 协议）——借本探测端到端验证，结果附进汇总供测试断言
        testReentrySeen = testReentrySeen & vbCrLf & "[版本更新探测] " & AttachLink_UpdateVersionCore("", "")
        reentryNote = vbCrLf & "[重入探测] " & testReentrySeen
    End If
    Dim baseDirty As Boolean

    Dim cellList As Collection
    Set cellList = FlattenSelection(targetRange)

    Dim items() As String, n As Long
    items = Split(srcPaths, "|")
    n = UBound(items) + 1

    ' V7：文件名含 # 时 Excel 会把链接地址截断成「附件\hash」加子地址，链接必断且修不了——
    ' 复制之前整批拒绝并列出清单（附件夹内不产生任何复制）
    Dim hashCnt As Long, hashList As String, i As Long
    For i = 0 To UBound(items)
        If Len(items(i)) > 0 Then
            If InStr(fso.GetFileName(items(i)), "#") > 0 Then
                hashCnt = hashCnt + 1
                If hashCnt <= 10 Then hashList = hashList & "、" & fso.GetFileName(items(i))
            End If
        End If
    Next i
    If hashCnt > 0 Then
        InsertCoreInner = "失败：以下文件名含有 #，Excel 超链接无法正确保存，请重命名后再插入：" & Mid$(hashList, 2)
        Exit Function
    End If

    Dim inserted As Long, skipped As Long, failed As Long
    Dim reusedCnt As Long, renamedCnt As Long, idx As Long
    Dim destPath As String, targetCell As Object, failList As String, renamedList As String
    Dim anomList As String
    Dim cat As Long, det As String

    For i = 0 To UBound(items)
        If Len(items(i)) > 0 Then
            ' 进度提示（T14 进度助手；WPS 不支持状态栏时静默跳过）
            Progress "正在登记附件指纹 " & (i + 1) & "/" & n & "：" & fso.GetFileName(items(i))
            ' T4：逐文件故障隔离——单个文件处理失败不中断整批
            On Error Resume Next
            Err.Clear
            destPath = ""
            cat = 0: det = ""
            destPath = ArchiveFileEx(fso, items(i), attachDir, base, baseDirty, cat, det)
            If Err.Number <> 0 Then
                failed = failed + 1
                failList = failList & "  " & fso.GetFileName(items(i)) & "：" & Err.Description & vbCrLf
                Err.Clear
                On Error GoTo Fail
            Else
                On Error GoTo Fail
                If Len(destPath) = 0 Then
                    skipped = skipped + 1
                Else
                    Do While idx >= cellList.Count
                        If Not AUTO_EXTEND_CELLS Then Exit Do
                        If Not ExtendCellList(cellList) Then Exit Do
                    Loop
                    If idx >= cellList.Count Then
                        skipped = skipped + 1
                    Else
                        Set targetCell = cellList(idx + 1)
                        ' V8a：受保护工作表上 Hyperlinks.Add 抛错——逐条故障隔离：
                        ' 该条计失败并写明原因（副本已复制并登记基线，可在未保护处
                        ' 重试、直接复用），整批与其余文件照常继续，不再穿出 Fail
                        On Error Resume Next
                        Err.Clear
                        InsertLink targetCell, _
                                   ATTACH_FOLDER_NAME & "\" & fso.GetFileName(destPath), _
                                   DisplayName(fso.GetFileName(destPath), fso)
                        Dim insErr As Long
                        insErr = Err.Number
                        If insErr <> 0 Then
                            Err.Clear
                            On Error GoTo Fail
                            failed = failed + 1
                            failList = failList & "  " & fso.GetFileName(destPath) & _
                                "：目标单元格所在工作表受保护，链接未能插入；文件已复制到附件夹并登记基线，" & _
                                "可在未保护的单元格重试插入（会直接复用）" & vbCrLf
                        Else
                            On Error GoTo Fail
                            idx = idx + 1
                            inserted = inserted + 1
                            If cat = CAT_REUSED Then reusedCnt = reusedCnt + 1
                            If cat = CAT_RENAMED Then
                                renamedCnt = renamedCnt + 1
                                If Len(renamedList) > 0 Then renamedList = renamedList & "、"
                                renamedList = renamedList & det
                            End If
                            ' O-05：副本与基线不一致未复用、已另存——单列说明（体检会标红的副本不得扩散新链接）
                            If cat = CAT_ANOMALY Then anomList = anomList & "  " & det & vbCrLf
                        End If
                    End If
                End If
            End If
        End If
    Next i
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo Fail

    ' 整批一次落盘；失败原因经 BaselinePersistNote 按 N-01 区分措辞
    Dim persistNote As String
    If baseDirty Then
        If Not SaveBaseline(fso, attachDir, base) Then
            persistNote = vbCrLf & "■ 基线登记未持久化（" & lastBaselineError & "），" & BaselinePersistNote(lastBaselineError) & "。"
        End If
    End If

    Dim msg As String
    msg = "已插入 " & inserted & " 个附件链接，文件已归档到同级「" & ATTACH_FOLDER_NAME & "」文件夹。"
    If reusedCnt > 0 Then
        msg = msg & vbCrLf & "其中 " & reusedCnt & " 个与附件夹内已有文件内容相同（已复用，未重复复制）。"
    End If
    If renamedCnt > 0 Then
        msg = msg & vbCrLf & renamedCnt & " 个与附件夹内同名文件内容不同，已另存为：" & renamedList & _
              "（如确需替换原附件，请对原链接使用「附件版本更新」）。"
    End If
    If Len(anomList) > 0 Then
        msg = msg & vbCrLf & "■ 以下附件夹内副本与其内容基线不一致（体检会标红），本次未复用、已另存：" & vbCrLf & anomList
    End If
    If skipped > 0 Then msg = msg & vbCrLf & "跳过 " & skipped & " 个文件（已取消或可用单元格不足）。"
    If failed > 0 Then msg = msg & vbCrLf & "失败 " & failed & " 个文件：" & vbCrLf & failList
    msg = msg & persistNote & reentryNote
    InsertCoreInner = msg
    Exit Function
Fail:
    On Error Resume Next
    Application.StatusBar = False
    InsertCoreInner = "失败：插入附件链接时出错 - " & Err.Description
End Function

'======================== 2. 附件清单 / 体检 ==========================
' 状态列（E）：正常 / 断链（文件不存在）
' 内容基线列（F）：一致 / ■ 内容变更 / 已登记基线 / 校验不可用
'   内容基线 = 附件夹内 _attachlink_baseline.csv（文件名|SHA-256，UTF-8），
'   随附件夹一起归档迁移/网盘同步。哈希不一致（附件被替换、篡改或同步端
'   同名覆盖）持续红字警示且【不自动更新基线】；确为有意替换时，删除该
'   基线文件后重新生成清单即可重建。
Public Sub AttachLink_BuildIndex()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then
        MsgBox "底稿尚未保存，请先保存后再生成附件清单（需检查附件文件是否存在）。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    ' V10：网页路径底稿无法生成本地附件清单
    If IsWebPathWorkbook(wb) Then
        MsgBox "底稿保存在网页路径（SharePoint/OneDrive 网页版），无法生成本地附件清单。" & vbCrLf & _
               "请先把底稿另存到本地磁盘或局域网共享盘后再体检。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    MsgBox AttachLink_BuildIndexCore(wb), vbInformation, "AttachLink 附件清单"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "生成附件清单时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 清单核心（独立公开函数：便于自动化测试），返回汇总文本。
' T14：处理进度经状态栏逐链接刷新（可 Esc 取消）；cancelAfter 为测试钩子——
' 处理完第 cancelAfter 个链接后进入与真实取消完全相同的代码路径。
' O-06/V2 重入防护包装：忙碌时直接失败返回；出错兜底——任何出口（含取消/
' 失败/Inner 抛未处理错误）都复位忙碌标志，错误如实转「失败：<过程名> - <原因>」
Public Function AttachLink_BuildIndexCore(ByVal wb As Object, Optional ByVal cancelAfter As Long = 0) As String
    If coreBusy Then
        AttachLink_BuildIndexCore = "失败：已有操作正在进行。"
        Exit Function
    End If
    coreBusy = True
    Err.Clear
    On Error Resume Next
    AttachLink_BuildIndexCore = BuildIndexInner(wb, cancelAfter)
    If Err.Number <> 0 Then
        AttachLink_BuildIndexCore = "失败：AttachLink_BuildIndexCore - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    coreBusy = False
End Function

Private Function BuildIndexInner(ByVal wb As Object, Optional ByVal cancelAfter As Long = 0) As String
    On Error GoTo Fail
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If wb Is Nothing Then
        BuildIndexInner = "失败：未提供工作簿。"
        Exit Function
    End If
    If Len(wb.Path) = 0 Then
        BuildIndexInner = "失败：底稿尚未保存，无法生成附件清单。"
        Exit Function
    End If
    ' V10：网页路径底稿（SharePoint/OneDrive「在桌面应用中打开」，wb.Path 以 https://
    ' 开头）下相对链接不可用，体检会把全部链接误判断链并清空旧清单——任何改动之前拒绝
    If IsWebPathWorkbook(wb) Then
        BuildIndexInner = "失败：底稿保存在网页路径（SharePoint/OneDrive 网页版），无法生成本地附件清单。请先把底稿另存到本地磁盘或局域网共享盘后再体检。"
        Exit Function
    End If
    Dim attachDir As String
    ' O-03：附件夹路径规范化（与链接解析路径同一坐标系，前缀/边界判定才可靠）
    attachDir = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)

    ' 内容基线（SHA-256；单文件失败显示「校验不可用」，不影响清单）
    ' 统一解析器：状态 2（损坏）时只做存在性检查，内容列红字「■ 基线损坏」且绝不落盘；
    ' readOnly（含更新版本字段）时可正常比对，但新登记无法持久化（SaveBaseline 拒绝覆盖）
    ResetHashCache
    Dim baseStatus As Long, baseRO As Boolean
    Dim base As Object
    Set base = LoadBaselineEx(fso, attachDir, baseStatus, baseRO)
    Dim baseDirty As Boolean
    Dim baseBroken As Boolean, baseErrText As String
    baseBroken = (baseStatus = 2)
    baseErrText = lastBaselineError

    ' O-06：改动运行环境前保存进入时的原值，取消/失败出口按原值还原
    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating

    ' V2：两个常见受保护触发点在【任何改动之前】给明确文案——清单表已存在且受保护 →
    ' 无法更新；工作簿结构受保护且没有清单表 → 无法新建（错误不再穿出包装器把
    ' coreBusy 永久锁死，审计底稿里受保护的工作表/结构很常见）
    Dim wsExist As Object
    On Error Resume Next
    Set wsExist = wb.Worksheets(INDEX_SHEET_NAME)
    On Error GoTo Fail
    If Not wsExist Is Nothing Then
        If wsExist.ProtectContents Then
            Application.ScreenUpdating = prevScreenUpdating
            BuildIndexInner = "失败：附件清单表受保护，无法更新清单。请先撤销该表的保护后再体检。"
            Exit Function
        End If
    Else
        If wb.ProtectStructure Then
            Application.ScreenUpdating = prevScreenUpdating
            BuildIndexInner = "失败：工作簿结构受保护，无法新建清单表。请先撤销工作簿保护后再体检。"
            Exit Function
        End If
    End If

    Application.ScreenUpdating = False
    Dim wsIdx As Object
    Set wsIdx = GetOrCreateIndexSheet(wb)
    If wsIdx Is Nothing Then
        Application.ScreenUpdating = prevScreenUpdating
        BuildIndexInner = "已存在同名「" & INDEX_SHEET_NAME & "」工作表，但它不是本加载项生成的（缺少归属标记）。" & vbCrLf & _
            "为避免误清空已有数据，本次未覆盖。请先手动改名或删除该表，再重新生成清单。"
        Exit Function
    End If
    wsIdx.Cells.Clear

    ' 标题与表头
    With wsIdx.Range("A1:G1")
        .Merge
        .Cells(1, 1).Value = "附件清单 —— " & wb.Name
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
    End With
    wsIdx.Range("A2").Value = "生成时间：" & Format(Now, "yyyy-mm-dd hh:nn")
    If baseBroken Then
        ' 清单顶部醒目说明：基线损坏，本次未做内容比对与登记
        wsIdx.Range("A2").Value = "生成时间：" & Format(Now, "yyyy-mm-dd hh:nn") & _
            "　■ 基线损坏（" & baseErrText & "），本次仅检查文件存在性，未做内容比对与登记"
        wsIdx.Range("A2").Font.Color = RGB(192, 0, 0)
        wsIdx.Range("A2").Font.Bold = True
    End If
    wsIdx.Range("A3:G3").Value = Array("编号", "附件名", "相对路径", "所在位置", "状态", "内容基线", "更替记录")

    Dim ws As Object, hl As Object
    Dim r As Long, total As Long, brokenCnt As Long, changedCnt As Long, unavailCnt As Long
    Dim cnt As Long, g As Long, absCnt As Long
    Dim linkPath As String, ok As Boolean
    Dim h As String, key As String, cellF As String, cellG As String, meta() As String
    ' 注意：错误处理变量名不得与 VBA 关键字同形（不区分大小写）——eNum 即 Enum，
    ' 会报「语法错误」；故用 errNum/errDesc
    Dim errNum As Long, errDesc As String

    ' T14：允许用户经 Esc / Ctrl+Break 取消（xlErrorHandler → Err 18 → 取消路径）；
    ' 读不到 / 设不上（WPS）则跳过，不影响体检
    Dim prevCancelKey As Long, cancelKeyOk As Boolean
    On Error Resume Next
    Err.Clear
    prevCancelKey = Application.EnableCancelKey
    If Err.Number = 0 Then cancelKeyOk = True
    Application.EnableCancelKey = 2             ' xlErrorHandler
    On Error GoTo Fail

    ' 先数出附件链接总数 N（进度分母与取消提示用）
    Dim totalLinks As Long
    totalLinks = 0
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, INDEX_SHEET_NAME, vbTextCompare) <> 0 Then
            For Each hl In ws.Hyperlinks
                If IsAttachmentLink(hl, fso, wb.Path) Then totalLinks = totalLinks + 1
            Next hl
        End If
    Next ws

    r = 4
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, INDEX_SHEET_NAME, vbTextCompare) <> 0 Then
            cnt = 0
            For Each hl In ws.Hyperlinks
                If IsAttachmentLink(hl, fso, wb.Path) Then cnt = cnt + 1
            Next hl
            If cnt > 0 Then
                ' 分组标题行
                With wsIdx.Range(wsIdx.Cells(r, 1), wsIdx.Cells(r, 7))
                    .Merge
                    .Cells(1, 1).Value = ws.Name & "（" & cnt & " 个附件）"
                    .Cells(1, 1).Font.Bold = True
                    .Interior.Color = RGB(217, 217, 217)
                End With
                r = r + 1
                g = 1
                For Each hl In ws.Hyperlinks
                    If IsAttachmentLink(hl, fso, wb.Path) Then
                        Progress "附件体检 " & (total + 1) & "/" & totalLinks & "：" & hl.TextToDisplay
                        wsIdx.Cells(r, 1).Value = ws.Name & "-" & Format(g, "00")
                        wsIdx.Cells(r, 2).Value = hl.TextToDisplay
                        wsIdx.Cells(r, 3).Value = hl.Address
                        wsIdx.Cells(r, 4).Value = ws.Name & "!" & hl.Range.Address(False, False)
                        linkPath = fso.GetAbsolutePathName(ResolveLinkPath(fso, wb.Path, hl.Address))
                        ok = fso.FileExists(linkPath)
                        cellF = "—"
                        cellG = ""
                        Dim isFolder As Boolean
                        isFolder = False
                        If Not ok Then isFolder = fso.FolderExists(linkPath)
                        If isFolder Then
                            wsIdx.Cells(r, 5).Value = "文件夹"
                            wsIdx.Cells(r, 5).Font.Color = RGB(0, 102, 204)
                            ' O-03：边界判定（前缀匹配会把「附件2\…」误判为夹内）
                            If Not IsSubPathOf(linkPath, attachDir) Then
                                ' 外部文件夹（如归档前的指回链接）：仅查存在性，不校验基线（键无从落在本附件夹内）
                                cellF = "外部，未校验"
                            ElseIf baseBroken Then
                                cellF = "■ 基线损坏"
                            Else
                                ' 夹内全部文件的内容基线体检（缺失自动登记 v1；哈希/遍历失败计入不可用）
                                Dim fCnt As Long, fReg As Long, fBad As Long, fFail As Long
                                fCnt = 0: fReg = 0: fBad = 0: fFail = 0
                                CheckFolderBaseline fso, base, linkPath, attachDir, fReg, fCnt, fBad, fFail, baseDirty
                                If fBad > 0 Then
                                    ' oracle #2：内容变更优先于校验不可用——篡改红字绝不被遮蔽
                                    cellF = "■ 内容变更"
                                    changedCnt = changedCnt + 1
                                    If fFail > 0 Then
                                        cellF = "■ 内容变更（含不可用）"
                                        unavailCnt = unavailCnt + 1
                                    End If
                                ElseIf fFail > 0 Then
                                    cellF = "校验不可用"
                                    unavailCnt = unavailCnt + 1
                                ElseIf fCnt = 0 Then
                                    cellF = "空文件夹"
                                ElseIf fReg > 0 Then
                                    cellF = "已登记 " & fCnt & " 文件"
                                Else
                                    cellF = fCnt & " 文件一致"
                                End If
                            End If
                        ElseIf ok Then
                            wsIdx.Cells(r, 5).Value = "正常"
                            wsIdx.Cells(r, 5).Font.Color = RGB(0, 128, 0)
                            ' T8：绝对路径且父目录=本底稿附件夹（保存时未能归一化的残留）→ 红字提示；
                            ' 体检只检测不改写，解除保护后重新保存即可由保存归一化自动修正
                            If InStr(hl.Address, ":") > 0 Or Left$(hl.Address, 2) = "\\" Then
                                If StrComp(fso.GetParentFolderName(linkPath), attachDir, vbTextCompare) = 0 Then
                                    wsIdx.Cells(r, 5).Value = "正常（■ 绝对路径）"
                                    wsIdx.Cells(r, 5).Font.Color = RGB(192, 0, 0)
                                    wsIdx.Cells(r, 5).Font.Bold = True
                                    absCnt = absCnt + 1
                                End If
                            End If
                            ' 内容基线校验（SHA-256；值=哈希|版本|登记时间|更新时间）
                            ' 键=附件夹内相对子路径；指向外部附件夹的链接
                            ' （迁移前形态）按文件名兜底。O-03：边界判定防「附件2\…」误判
                            If IsSubPathOf(linkPath, attachDir) Then
                                key = LCase$(Mid$(linkPath, Len(attachDir) + 2))
                            Else
                                key = LCase$(fso.GetFileName(linkPath))
                            End If
                            cellG = ""
                            If baseBroken Then
                                ' 基线损坏：绝不显示「一致」或「已登记」，红字提示人工处置
                                cellF = "■ 基线损坏"
                            Else
                            h = HashFileCached(linkPath)
                            If Len(h) = 0 Then
                                cellF = "校验不可用"
                                unavailCnt = unavailCnt + 1
                            ElseIf Not base.Exists(key) Then
                                base.Add key, h & "|1|" & Format$(Now, "yyyy-mm-dd hh:nn") & "|" & Format$(Now, "yyyy-mm-dd hh:nn")
                                baseDirty = True
                                cellF = "已登记 v1"
                            Else
                                meta = Split(base(key), "|")
                                If StrComp(meta(0), h, vbTextCompare) = 0 Then
                                    cellF = "一致 v" & meta(1)
                                Else
                                    ' 不更新基线：持续警示，直到走「附件版本更新」或人工重建基线
                                    cellF = "■ 内容变更"
                                    changedCnt = changedCnt + 1
                                End If
                                ' 更替记录：版本 >=2 的附件显式备注最近一次替换动作（审计留痕）
                                If IsNumeric(meta(1)) Then
                                    If CLng(meta(1)) >= 2 Then
                                        cellG = "v" & (CLng(meta(1)) - 1) & "→v" & meta(1) & " " & meta(3)
                                    End If
                                End If
                            End If
                            End If
                        Else
                            wsIdx.Cells(r, 5).Value = "断链"
                            wsIdx.Cells(r, 5).Font.Color = RGB(192, 0, 0)
                            wsIdx.Cells(r, 5).Font.Bold = True
                            brokenCnt = brokenCnt + 1
                        End If
                        wsIdx.Cells(r, 6).Value = cellF
                        If Left$(cellF, 6) = "■ 内容变更" Or Left$(cellF, 6) = "■ 基线损坏" Then
                            wsIdx.Cells(r, 6).Font.Color = RGB(192, 0, 0)
                            wsIdx.Cells(r, 6).Font.Bold = True
                        ElseIf Left$(cellF, 2) = "一致" Or Left$(cellF, 3) = "已登记" Then
                            wsIdx.Cells(r, 6).Font.Color = RGB(0, 128, 0)
                        End If
                        If Len(cellG) > 0 Then
                            wsIdx.Cells(r, 7).Value = cellG
                            wsIdx.Cells(r, 7).Font.Color = RGB(0, 102, 204)
                        End If
                        total = total + 1
                        g = g + 1
                        r = r + 1
                        ' T14 测试钩子：处理完第 cancelAfter 个链接后进入与真实取消相同的代码路径
                        If cancelAfter > 0 Then
                            If total >= cancelAfter Then GoTo Cancelled
                        End If
                    End If
                Next hl
            End If
        End If
    Next ws

    ' 基线落盘（有新登记才写；损坏时不写）。失败时汇总警告（登记未持久化，下次体检将重新登记）
    Dim persistFail As Boolean, persistErr As String
    If baseDirty And Not baseBroken Then
        If Not SaveBaseline(fso, attachDir, base) Then
            persistFail = True
            persistErr = lastBaselineError
        End If
    End If

    ' 汇总行
    If total = 0 Then
        wsIdx.Cells(r + 1, 1).Value = "未发现任何附件链接。"
    Else
        wsIdx.Cells(r + 1, 1).Value = "合计 " & total & " 个附件：正常 " & (total - brokenCnt) & _
            "，断链 " & brokenCnt & IIf(brokenCnt > 0, "（可用右键菜单「附件断链修复」处理）", "") & _
            IIf(changedCnt > 0, "，内容变更 " & changedCnt, "")
        wsIdx.Cells(r + 1, 1).Font.Bold = True
    End If

    ' 版式
    With wsIdx
        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 42
        .Columns("D").ColumnWidth = 26
        .Columns("E").ColumnWidth = 8
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 24
        .Range("A3:G3").Font.Bold = True
        .Range("A3:G3").Interior.Color = RGB(237, 237, 237)
        If r > 4 Then .Range("A3:G" & (r - 1)).Borders.LineStyle = 1
    End With

    Application.ScreenUpdating = prevScreenUpdating
    On Error Resume Next
    Application.StatusBar = False
    If cancelKeyOk Then Application.EnableCancelKey = prevCancelKey
    On Error GoTo Fail
    wsIdx.Activate

    Dim msg As String
    If total = 0 Then
        msg = "当前底稿中未发现附件链接。"
    Else
        msg = "附件清单已更新：共 " & total & " 个附件，正常 " & (total - brokenCnt) & "，断链 " & brokenCnt
        If changedCnt > 0 Then
            msg = msg & "，内容变更 " & changedCnt & " 个（红字标注——未经「附件版本更新」流程的替换/篡改/同步同名覆盖）。" & vbCrLf & _
                "如确为有意的版本更替，请选中该附件链接单元格，使用右键「附件版本更新」登记（留痕）；" & _
                "全部重建基线才需删除附件夹内 " & BASELINE_FILE & " 后重新生成清单。"
        ElseIf brokenCnt > 0 Then
            msg = msg & "。可用右键菜单「附件断链修复」重连。"
        ElseIf unavailCnt > 0 Then
            msg = msg & "，内容校验不可用 " & unavailCnt & " 个（SHA-256 计算失败）。"
        ElseIf baseBroken Then
            ' V9：基线损坏时不得再以「内容基线全部一致」收尾（同一提示框前后矛盾）
            msg = msg & "，■ 基线文件损坏，本次未做内容比对、未写入任何登记。"
        Else
            msg = msg & "，内容基线全部一致。"
        End If
    End If
    If absCnt > 0 Then
        msg = msg & "，绝对路径未归一化 " & absCnt & " 个。" & vbCrLf & _
              "这些链接保存时未能归一化为相对路径，工作表受保护或只读时会发生；搬移底稿后会断链，解除保护后重新保存（Ctrl+S）即可自动修正。"
    End If
    If baseBroken Then
        msg = "■ 基线损坏（" & baseErrText & "）：本次仅检查文件存在性，未做内容比对与登记；" & _
              "请人工修复或删除附件夹内 " & BASELINE_FILE & " 后重新生成清单重建基线。" & vbCrLf & msg
    End If
    ' N-01：落盘失败按原因区分提示（损坏/未来格式只读不写「体检将补登记」、不提示检查写权限）
    If persistFail Then
        msg = msg & vbCrLf & "■ 基线文件写入失败：本次新登记未持久化（" & persistErr & "）。" & BaselinePersistNote(persistErr)
    End If
    BuildIndexInner = msg
    Exit Function
Cancelled:
    ' T14 取消路径：不保存基线（丢弃本次登记）；清单表清空后 A1 红字说明；还原运行环境
    On Error Resume Next
    Application.StatusBar = False
    If cancelKeyOk Then Application.EnableCancelKey = prevCancelKey
    Application.ScreenUpdating = prevScreenUpdating
    wsIdx.Cells.Clear
    wsIdx.Range("A1").Value = "体检已取消（已处理 " & total & "/" & totalLinks & " 个附件）。基线未改动。"
    wsIdx.Range("A1").Font.Color = RGB(192, 0, 0)
    wsIdx.Range("A1").Font.Bold = True
    BuildIndexInner = "已取消：已处理 " & total & "/" & totalLinks & " 个附件，基线未改动。"
    Exit Function
Fail:
    If Err.Number = 18 Then GoTo Cancelled          ' 用户按 Esc / Ctrl+Break
    ' V2：其它错误还原运行环境（按进入时保存的原值）后如实返回「失败：…」，
    ' 不再上抛——未处理错误穿出包装器会把 coreBusy 永久锁死、本会话全部功能不可用
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.StatusBar = False
    If cancelKeyOk Then Application.EnableCancelKey = prevCancelKey
    Err.Clear
    BuildIndexInner = "失败：生成附件清单时出错 - " & errDesc
End Function

'============================ 3. 断链修复 =============================
Public Sub AttachLink_Repair()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then
        MsgBox "底稿尚未保存，请先保存后再执行修复。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If InStr(1, wb.Path, "://", vbTextCompare) > 0 Then
        MsgBox "底稿保存在网页路径（SharePoint/OneDrive 网页版），无法按本地目录修复断链。" & vbCrLf & _
               "请先把底稿另存到本地磁盘或局域网共享盘后再执行。", vbExclamation, "AttachLink"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 统计断链数（oracle #1：文件夹的 FileExists 恒为 False，必须同时排除存在的文件夹，
    ' 否则健康文件夹会被误列入修复并用源目录合并覆盖）
    Dim brokenCnt As Long
    Dim ws As Object, hl As Object, rp As String
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, INDEX_SHEET_NAME, vbTextCompare) <> 0 Then
            For Each hl In ws.Hyperlinks
                If IsAttachmentLink(hl, fso, wb.Path) Then
                    rp = ResolveLinkPath(fso, wb.Path, hl.Address)
                    If Not fso.FileExists(rp) And Not fso.FolderExists(rp) Then brokenCnt = brokenCnt + 1
                End If
            Next hl
        End If
    Next ws
    If brokenCnt = 0 Then
        MsgBox "未发现断链附件，无需修复。", vbInformation, "AttachLink"
        Exit Sub
    End If

    ' 选择附件文件夹当前位置
    Dim fd As Object
    Set fd = Application.FileDialog(4)
    With fd
        .Title = "选择附件文件当前所在的文件夹（共 " & brokenCnt & " 个断链待修复）"
        .InitialFileName = wb.Path & "\"
        If .Show <> -1 Then Exit Sub
    End With
    Dim srcDir As String
    srcDir = fd.SelectedItems(1)
    If Right$(srcDir, 1) = "\" Then srcDir = Left$(srcDir, Len(srcDir) - 1)

    ' N-08：结果含未校验/冲突/重连失败/路径异常时用警示图标，不再一律 vbInformation
    Dim repairMsg As String, repairIcon As Long
    repairMsg = AttachLink_RepairCore(wb, srcDir)
    repairIcon = vbInformation
    If InStr(repairMsg, "未校验") > 0 Or InStr(repairMsg, "冲突") > 0 Or _
       InStr(repairMsg, "重连失败") > 0 Or InStr(repairMsg, "路径异常") > 0 Then repairIcon = vbExclamation
    MsgBox repairMsg, repairIcon, "AttachLink"
    Exit Sub
Fail:
    MsgBox "修复时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 断链修复核心（无对话框，供自动化直调）：遍历底稿断链，按 T3 可信重连规则
' 重连并返回多行汇总。判定顺序（文件链接）：
'   ① 本附件夹已有同名文件：内容与期望基线一致→只改指不复制；不一致→冲突不动；
'   ② 所选目录同名文件：哈希一致→重连；不一致→不按名连，转③；无基线→按名重连并标注未校验；
'   ③ 所选目录（递归）内容一致候选：恰好一个→以原文件名落回重连；多个→候选不唯一不重连；
'   ④ 都没有：有同名异内容文件→「同名但内容不同，未重连」；否则「未找到」。
' 文件夹链接：本附件夹已有同名文件夹→逐文件比对（O-01）；不存在→按名重连，
'   基线有「相对子路径\」前缀条目时先逐个核对内容，全部一致才合并复制。
' O-02：附件夹内链接只用完整相对子路径的基线键，不再回退借根目录同名键；
' O-03：链接统一经 ResolveAttachTarget 规范化+边界校验——自称「附件\」却解析到
'       附件夹外的（如 附件\..\x.pdf）计「链接路径异常」，不复制、不改指、不登记；
'       附件夹之外的归档前指回形态（绝对路径 / ..\ 形态）仍按文件名在夹内根处理。
' O-01：本附件夹已有同名文件夹时逐文件比对（修复源↔本地、本地↔基线），任一
'       同名文件缺失/不一致/哈希不可用即整项不动；全部一致仅改指不复制；
'       本地多出修复源没有的文件不算冲突，但在汇总里如实说明。
' O-04：基线损坏 = 失败关闭（任何复制/合并/改指之前返回）。
' O-06：重入防护包装——忙碌时直接失败返回，包装器保证任何出口都复位标志。
Public Function AttachLink_RepairCore(ByVal wb As Object, ByVal srcDir As String) As String
    If coreBusy Then
        AttachLink_RepairCore = "失败：已有操作正在进行。"
        Exit Function
    End If
    coreBusy = True
    Err.Clear
    On Error Resume Next
    AttachLink_RepairCore = RepairCoreInner(wb, srcDir)
    If Err.Number <> 0 Then
        AttachLink_RepairCore = "失败：AttachLink_RepairCore - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    coreBusy = False
End Function

Private Function RepairCoreInner(ByVal wb As Object, ByVal srcDir As String) As String
    On Error GoTo Fail
    If wb Is Nothing Then
        RepairCoreInner = "失败：未提供工作簿。"
        Exit Function
    End If
    If Len(wb.Path) = 0 Then
        RepairCoreInner = "失败：底稿尚未保存。"
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcDir) Then
        RepairCoreInner = "失败：所选文件夹不存在：" & srcDir
        Exit Function
    End If
    srcDir = fso.GetFolder(srcDir).Path
    ResetHashCache
    Dim attachDir As String
    attachDir = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)
    ' O-04：基线损坏 = 失败关闭——不得绕过附件身份校验按「无基线」乱连
    Dim baseStatus As Long, baseRO As Boolean
    Dim base As Object
    Set base = LoadBaselineEx(fso, attachDir, baseStatus, baseRO)
    If baseStatus = 2 Then
        RepairCoreInner = "失败：基线文件损坏或无法读取（" & lastBaselineError & _
                          "），为避免绕过附件身份校验，本次未执行任何修复。请先人工核对或重建基线。"
        Exit Function
    End If
    EnsureFolder fso, attachDir
    Dim baseDirty As Boolean

    ' 收集断链（规范化后判存在性：「附件\..\x」这类逃逸形态解析到夹外，按断链进入主循环，
    ' 再由 O-03 闸口归类为「链接路径异常」，绝不复制/改指）
    Dim broken As Collection
    Set broken = New Collection
    Dim ws As Object, hl As Object, rp As String
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, INDEX_SHEET_NAME, vbTextCompare) <> 0 Then
            For Each hl In ws.Hyperlinks
                If IsAttachmentLink(hl, fso, wb.Path) Then
                    rp = fso.GetAbsolutePathName(ResolveLinkPath(fso, wb.Path, hl.Address))
                    If Not fso.FileExists(rp) And Not fso.FolderExists(rp) Then broken.Add hl
                End If
            Next hl
        End If
    Next ws
    If broken.Count = 0 Then
        RepairCoreInner = "未发现断链附件，无需修复。"
        Exit Function
    End If

    ' 所选目录递归文件索引（按内容查找候选；哈希经缓存去重，每文件一次运行只算一次）
    Dim allFiles As Collection
    Set allFiles = New Collection
    CollectFilesRecursive fso, srcDir, allFiles

    Dim cSame As Long, cContent As Long, cLocal As Long, cUnver As Long
    Dim cDiff As Long, cMulti As Long, cMiss As Long, cConflict As Long, cFail As Long
    Dim cFoldOk As Long, cFoldUnver As Long, cFoldBad As Long
    Dim cFoldConflict As Long, cAnomaly As Long
    Dim lSame As String, lContent As String, lLocal As String, lUnver As String
    Dim lDiff As String, lMulti As String, lMiss As String, lConflict As String, lFail As String
    Dim lFoldBad As String, lFoldConflict As String, lAnomaly As String, lFoldOk As String
    Dim lFoldUnver As String

    Dim i As Long, fName As String, linkPath As String, key As String, expected As String
    Dim localPath As String, candPath As String, hC As String, hLocal As String, hF As String
    Dim handled As Boolean, cand As Variant, rel As String
    Dim hits As Collection, hitList As String
    Dim k As Variant, fAny As Boolean, fBad As Boolean
    Dim rg As Long, fc As Long, ff As Long
    Dim relSub As String, relAddr As String, repErr As String
    Dim wantExt As String, scanned As Long, candTotal As Long
    Dim norm As String, isRelClaim As Boolean, extraNote As String
    Dim candFiles As Collection, locFiles As Collection, extraCnt As Long
    Dim sz As Variant, mt As Variant, hV As String, foldKey As String

    For i = 1 To broken.Count
        Set hl = broken(i)
        fName = fso.GetFileName(hl.Address)
        ' O-03：统一规范化 + 附件夹边界判定
        norm = fso.GetAbsolutePathName(ResolveLinkPath(fso, wb.Path, hl.Address))
        isRelClaim = False
        If Not (InStr(hl.Address, ":") > 0 Or Left$(hl.Address, 2) = "\\") Then
            ' V6：附件/a.pdf 正斜杠形态同样算自称「附件\」的相对链接
            If StrComp(Left$(NormSlash(hl.Address), Len(ATTACH_FOLDER_NAME) + 1), ATTACH_FOLDER_NAME & "\", vbTextCompare) = 0 Then isRelClaim = True
        End If
        relSub = "": key = "": expected = ""
        If ResolveAttachTarget(fso, wb.Path, attachDir, hl.Address, relSub) <> "" Then
            ' 链接解析路径在本附件夹内：基线键 = 完整相对子路径（O-02，不再回退借根目录同名键）
            key = LCase$(relSub)
        ElseIf isRelClaim Then
            ' 自称「附件\」却解析到附件夹外（如 附件\..\outside.pdf）：路径异常，不复制/不改指/不登记
            cAnomaly = cAnomaly + 1
            If cAnomaly <= 10 Then lAnomaly = lAnomaly & "、" & hl.Address
        Else
            ' 归档前指回形态（绝对路径 / ..\ 形态指向别处附件夹）：按文件名在本附件夹根处理
            relSub = fName
            key = LCase$(fName)
        End If
        If base.Exists(key) Then expected = Split(base(key), "|")(0)
        If Len(relSub) = 0 Then GoTo NextLink
        localPath = attachDir & "\" & relSub
        relAddr = ATTACH_FOLDER_NAME & "\" & relSub
        ' A4：逐链接故障隔离——单个链接处理抛错只计一项「重连失败」，继续处理其余链接；
        ' 重连成功的登记在循环结束后照常整批落盘
        On Error Resume Next
        Err.Clear

        If fso.FolderExists(srcDir & "\" & fName) Then
            ' ---- 文件夹链接 ----
            candPath = srcDir & "\" & fName
            If fso.FolderExists(localPath) Then
                ' O-01：本附件夹已有同名文件夹——逐文件比对（修复源↔本地、本地↔基线），
                ' 任何不一致/缺失/哈希不可用 → 整项不动，绝不带覆盖合并（旧版 CopyFolder True
                ' 会把本地较新文件覆盖掉、还会让「内容变更」警示消失）
                fBad = False
                Set candFiles = New Collection
                CollectFilesRecursive fso, candPath, candFiles
                Err.Clear
                For Each cand In candFiles
                    rel = Mid$(CStr(cand), Len(candPath) + 2)
                    If Not fso.FileExists(localPath & "\" & rel) Then
                        fBad = True
                        Exit For
                    End If
                    hF = HashFileCached(CStr(cand))
                    hLocal = HashFileCached(localPath & "\" & rel)
                    If Len(hF) = 0 Or Len(hLocal) = 0 Then
                        fBad = True
                        Exit For
                    End If
                    If StrComp(hF, hLocal, vbTextCompare) <> 0 Then
                        fBad = True
                        Exit For
                    End If
                    ' 本地文件对不上基线也算冲突（已登记文件被改动）
                    foldKey = LCase$(relSub & "\" & rel)
                    If base.Exists(foldKey) Then
                        If StrComp(hLocal, Split(base(foldKey), "|")(0), vbTextCompare) <> 0 Then
                            fBad = True
                            Exit For
                        End If
                    End If
                Next cand
                Err.Clear
                If fBad Then
                    cFoldConflict = cFoldConflict + 1
                    If cFoldConflict <= 10 Then lFoldConflict = lFoldConflict & "、" & fName
                Else
                    ' 本地多出的文件不算冲突，但要在汇总里如实说明
                    extraCnt = 0
                    Set locFiles = New Collection
                    CollectFilesRecursive fso, localPath, locFiles
                    Err.Clear
                    For Each cand In locFiles
                        rel = Mid$(CStr(cand), Len(localPath) + 2)
                        If Not fso.FileExists(candPath & "\" & rel) Then extraCnt = extraCnt + 1
                    Next cand
                    Err.Clear
                    extraNote = ""
                    If extraCnt > 0 Then extraNote = "（本地另有 " & extraCnt & " 个修复源没有的文件，未改动）"
                    If TryRepoint(hl, relAddr, repErr) Then
                        cFoldOk = cFoldOk + 1
                        If cFoldOk <= 10 Then lFoldOk = lFoldOk & "、" & fName & extraNote
                    Else
                        cFail = cFail + 1
                        If cFail <= 10 Then lFail = lFail & "、" & fName & "（重连失败：" & repErr & "）"
                    End If
                End If
            Else
                ' 本附件夹尚无同名文件夹：先核对修复源与基线（O-01 前缀用完整相对子路径），
                ' 有基线条目且任一不符 → 不合并；全部一致或无基线 → 合并复制并补登记
                fAny = False: fBad = False
                For Each k In base.Keys
                    If Left$(k, Len(relSub) + 1) = LCase$(relSub) & "\" Then
                        fAny = True
                        rel = Mid$(k, Len(relSub) + 2)
                        If Not fso.FileExists(candPath & "\" & rel) Then
                            fBad = True
                            Exit For
                        End If
                        hF = HashFileCached(candPath & "\" & rel)
                        If Len(hF) = 0 Then
                            fBad = True
                            Exit For
                        End If
                        If StrComp(hF, Split(base(k), "|")(0), vbTextCompare) <> 0 Then
                            fBad = True
                            Exit For
                        End If
                    End If
                Next k
                If fAny And fBad Then
                    cFoldBad = cFoldBad + 1
                    If cFoldBad <= 10 Then lFoldBad = lFoldBad & "、" & fName
                Else
                    EnsureFolderDeep fso, fso.GetParentFolderName(localPath)
                    On Error Resume Next
                    Err.Clear
                    fso.CopyFolder candPath, localPath, True
                    If Err.Number <> 0 Then
                        cFail = cFail + 1
                        If cFail <= 10 Then lFail = lFail & "、" & fName & "（" & Err.Description & "）"
                        Err.Clear
                    Else
                        rg = 0: fc = 0: ff = 0
                        RegisterFolderFilesBaseline fso, base, localPath, attachDir, rg, fc, ff
                        If rg > 0 Then baseDirty = True
                        If TryRepoint(hl, relAddr, repErr) Then
                            If fAny Then
                                cFoldOk = cFoldOk + 1
                                If cFoldOk <= 10 Then lFoldOk = lFoldOk & "、" & fName
                            Else
                                cFoldUnver = cFoldUnver + 1
                                If cFoldUnver <= 10 Then lFoldUnver = lFoldUnver & "、" & fName
                            End If
                        Else
                            cFail = cFail + 1
                            If cFail <= 10 Then lFail = lFail & "、" & fName & "（文件夹已落回，但链接改指失败：" & repErr & "）"
                        End If
                    End If
                End If
            End If
        Else
            ' ---- 文件链接 ----
            handled = False
            If fso.FileExists(localPath) Then
                ' ① 本附件夹已有同名文件（含子路径原位）
                If Len(expected) > 0 Then
                    hLocal = HashFileCached(localPath)
                    If Len(hLocal) > 0 Then
                        If StrComp(hLocal, expected, vbTextCompare) = 0 Then
                            If TryRepoint(hl, relAddr, repErr) Then
                                cLocal = cLocal + 1
                                If cLocal <= 10 Then lLocal = lLocal & "、" & fName
                            Else
                                cFail = cFail + 1
                                If cFail <= 10 Then lFail = lFail & "、" & fName & "（重连失败：" & repErr & "）"
                            End If
                        Else
                            cConflict = cConflict + 1
                            If cConflict <= 10 Then lConflict = lConflict & "、" & fName
                        End If
                    Else
                        cFail = cFail + 1
                        If cFail <= 10 Then lFail = lFail & "、" & fName & "（哈希不可用）"
                    End If
                Else
                    If TryRepoint(hl, relAddr, repErr) Then
                        EnsureBaselineKey base, LCase$(relSub), HashFileCached(localPath), baseDirty
                        cUnver = cUnver + 1
                        If cUnver <= 10 Then lUnver = lUnver & "、" & fName
                    Else
                        cFail = cFail + 1
                        If cFail <= 10 Then lFail = lFail & "、" & fName & "（重连失败：" & repErr & "）"
                    End If
                End If
                handled = True
            Else
                candPath = srcDir & "\" & fName
                If fso.FileExists(candPath) Then
                    If Len(expected) = 0 Then
                        ' O-02 护栏：子目录断链无基线时，若候选内容与本附件夹根目录同名基线一致，
                        ' 说明唯一能找回的就是根目录文件——不得张冠李戴复制进子目录（旧版借根目录
                        ' 键找回的缺陷形态），按「同名但内容不同」如实归类、不动
                        If InStr(relSub, "\") > 0 And base.Exists(LCase$(fName)) Then
                            hC = HashFileCached(candPath)
                            If Len(hC) > 0 Then
                                If StrComp(hC, Split(base(LCase$(fName)), "|")(0), vbTextCompare) = 0 Then
                                    cDiff = cDiff + 1
                                    If cDiff <= 10 Then lDiff = lDiff & "、" & fName
                                    handled = True
                                End If
                            End If
                        End If
                        If Not handled Then
                            ' ② 无基线：按名重连并标注未校验，登记 v1 建立身份
                            hC = HashFileCached(candPath)
                            sz = fso.GetFile(candPath).Size
                            mt = fso.GetFile(candPath).DateLastModified
                            EnsureFolderDeep fso, fso.GetParentFolderName(localPath)
                            On Error Resume Next
                            Err.Clear
                            fso.CopyFile candPath, localPath, False
                            If Err.Number <> 0 Then
                                cFail = cFail + 1
                                If cFail <= 10 Then lFail = lFail & "、" & fName & "（" & Err.Description & "）"
                                Err.Clear
                            Else
                                ' N-06：复制后校验（源被改动/目标大小不符 → 对目标重新哈希）
                                hV = VerifyCopiedHash(fso, candPath, localPath, sz, mt, hC)
                                If Len(hV) = 0 Then
                                    cFail = cFail + 1
                                    If cFail <= 10 Then lFail = lFail & "、" & fName & "（复制后校验失败）"
                                Else
                                    CacheHash localPath, hV
                                    EnsureBaselineKey base, LCase$(relSub), hV, baseDirty
                                    If TryRepoint(hl, relAddr, repErr) Then
                                        cUnver = cUnver + 1
                                        If cUnver <= 10 Then lUnver = lUnver & "、" & fName
                                    Else
                                        cFail = cFail + 1
                                        If cFail <= 10 Then lFail = lFail & "、" & fName & "（文件已落回附件夹，但链接改指失败：" & repErr & "）"
                                    End If
                                End If
                            End If
                        End If
                        handled = True
                    Else
                        hC = HashFileCached(candPath)
                        If Len(hC) > 0 Then
                            If StrComp(hC, expected, vbTextCompare) = 0 Then
                                ' ② 同名且内容一致：复制落回并重连
                                sz = fso.GetFile(candPath).Size
                                mt = fso.GetFile(candPath).DateLastModified
                                EnsureFolderDeep fso, fso.GetParentFolderName(localPath)
                                On Error Resume Next
                                Err.Clear
                                fso.CopyFile candPath, localPath, False
                                If Err.Number <> 0 Then
                                    cFail = cFail + 1
                                    If cFail <= 10 Then lFail = lFail & "、" & fName & "（" & Err.Description & "）"
                                    Err.Clear
                                Else
                                    ' N-06：复制后校验，不一致时对目标重新哈希
                                    hV = VerifyCopiedHash(fso, candPath, localPath, sz, mt, hC)
                                    If Len(hV) = 0 Then
                                        cFail = cFail + 1
                                        If cFail <= 10 Then lFail = lFail & "、" & fName & "（复制后校验失败）"
                                    Else
                                        CacheHash localPath, hV
                                        EnsureBaselineKey base, LCase$(relSub), hV, baseDirty
                                        If TryRepoint(hl, relAddr, repErr) Then
                                            cSame = cSame + 1
                                            If cSame <= 10 Then lSame = lSame & "、" & fName
                                        Else
                                            cFail = cFail + 1
                                            If cFail <= 10 Then lFail = lFail & "、" & fName & "（文件已落回附件夹，但链接改指失败：" & repErr & "）"
                                        End If
                                    End If
                                End If
                                handled = True
                            End If
                        End If
                        ' 同名异内容 / 哈希不可用：不按名连，转入按内容查找
                    End If
                End If
            End If
            If Not handled Then
                If Len(expected) > 0 Then
                    ' ③ 按内容查找（A5：仅在扩展名相同的文件里找；查找期间刷新进度防假死）
                    wantExt = LCase$(fso.GetExtensionName(fName))
                    candTotal = 0
                    For Each cand In allFiles
                        If StrComp(LCase$(fso.GetExtensionName(CStr(cand))), wantExt, vbTextCompare) = 0 Then candTotal = candTotal + 1
                    Next cand
                    Set hits = New Collection
                    scanned = 0
                    For Each cand In allFiles
                        If StrComp(LCase$(fso.GetExtensionName(CStr(cand))), wantExt, vbTextCompare) = 0 Then
                            scanned = scanned + 1
                            Progress "断链修复：按内容查找 " & fName & "（" & scanned & "/" & candTotal & "）"
                            hC = HashFileCached(CStr(cand))
                            If Len(hC) > 0 Then
                                If StrComp(hC, expected, vbTextCompare) = 0 Then hits.Add CStr(cand)
                            End If
                        End If
                    Next cand
                    If hits.Count = 1 Then
                        ' 恰好一个：以链接原相对子路径落回附件夹（源不动），链接地址不变
                        sz = fso.GetFile(hits(1)).Size
                        mt = fso.GetFile(hits(1)).DateLastModified
                        EnsureFolderDeep fso, fso.GetParentFolderName(localPath)
                        On Error Resume Next
                        Err.Clear
                        fso.CopyFile hits(1), localPath, False
                        If Err.Number <> 0 Then
                            cFail = cFail + 1
                            If cFail <= 10 Then lFail = lFail & "、" & fName & "（" & Err.Description & "）"
                            Err.Clear
                        Else
                            ' N-06：复制后校验，不一致时对目标重新哈希（不得沿用期望哈希登记）
                            hV = VerifyCopiedHash(fso, hits(1), localPath, sz, mt, expected)
                            If Len(hV) = 0 Then
                                cFail = cFail + 1
                                If cFail <= 10 Then lFail = lFail & "、" & fName & "（复制后校验失败）"
                            Else
                                CacheHash localPath, hV
                                EnsureBaselineKey base, LCase$(relSub), hV, baseDirty
                                If TryRepoint(hl, relAddr, repErr) Then
                                    cContent = cContent + 1
                                    If cContent <= 10 Then lContent = lContent & "、" & fName & "←" & fso.GetFileName(hits(1))
                                Else
                                    cFail = cFail + 1
                                    If cFail <= 10 Then lFail = lFail & "、" & fName & "（文件已落回附件夹，但链接改指失败：" & repErr & "）"
                                End If
                            End If
                        End If
                    ElseIf hits.Count > 1 Then
                        hitList = ""
                        Dim hi As Long
                        For hi = 1 To hits.Count
                            If hi > 10 Then
                                hitList = hitList & "；…"
                                Exit For
                            End If
                            If hi > 1 Then hitList = hitList & "；"
                            hitList = hitList & fso.GetFileName(hits(hi))
                        Next hi
                        cMulti = cMulti + 1
                        If cMulti <= 10 Then lMulti = lMulti & "、" & fName & "（候选：" & hitList & "）"
                    Else
                        If fso.FileExists(srcDir & "\" & fName) Then
                            cDiff = cDiff + 1
                            If cDiff <= 10 Then lDiff = lDiff & "、" & fName
                        Else
                            cMiss = cMiss + 1
                            If cMiss <= 10 Then lMiss = lMiss & "、" & fName
                        End If
                    End If
                Else
                    cMiss = cMiss + 1
                    If cMiss <= 10 Then lMiss = lMiss & "、" & fName
                End If
            End If
        End If
NextLink:
        ' A4：逐链接故障隔离的检查点——本链接处理期间的任何未捕获错误只计一项失败
        If Err.Number <> 0 Then
            cFail = cFail + 1
            If cFail <= 10 Then lFail = lFail & "、" & fName & "（处理出错：" & Err.Description & "）"
            Err.Clear
        End If
    Next i
    On Error GoTo Fail

    ' 重连成功的登记整批一次落盘（失败原因经 BaselinePersistNote 按 N-01 区分措辞）
    Dim persistNote As String
    If baseDirty Then
        If Not SaveBaseline(fso, attachDir, base) Then
            persistNote = vbCrLf & "■ 基线登记未持久化（" & lastBaselineError & "），" & BaselinePersistNote(lastBaselineError) & "。"
        End If
    End If

    Dim sb As String
    sb = "断链修复完成：共 " & broken.Count & " 个断链。"
    ' N-08：未校验类置顶，便于用户先看到身份未确认的附件
    If cUnver > 0 Then sb = sb & vbCrLf & "· 重连·未校验（无基线）：" & cUnver & " 个" & ListSuffix(lUnver)
    If cFoldUnver > 0 Then sb = sb & vbCrLf & "· 重连·文件夹未校验（无基线）：" & cFoldUnver & " 个" & ListSuffix(lFoldUnver)
    If cSame > 0 Then sb = sb & vbCrLf & "· 重连·内容一致（同名）：" & cSame & " 个" & ListSuffix(lSame)
    If cContent > 0 Then sb = sb & vbCrLf & "· 重连·内容一致（按内容找回）：" & cContent & " 个" & ListSuffix(lContent)
    If cLocal > 0 Then sb = sb & vbCrLf & "· 重连·本附件夹已有同名副本（仅改指，未复制）：" & cLocal & " 个" & ListSuffix(lLocal)
    If cFoldOk > 0 Then sb = sb & vbCrLf & "· 重连·文件夹内容一致（已合并或仅改指）：" & cFoldOk & " 个" & ListSuffix(lFoldOk)
    If cAnomaly > 0 Then sb = sb & vbCrLf & "· 链接路径异常（指向附件夹之外），未处理：" & cAnomaly & " 个" & ListSuffix(lAnomaly)
    If cConflict > 0 Then sb = sb & vbCrLf & "· 冲突·未处理（附件夹内同名文件内容不同）：" & cConflict & " 个" & ListSuffix(lConflict)
    If cFoldConflict > 0 Then sb = sb & vbCrLf & "· 冲突·未处理（附件夹内同名文件夹内容不同）：" & cFoldConflict & " 个" & ListSuffix(lFoldConflict)
    If cFoldBad > 0 Then sb = sb & vbCrLf & "· 文件夹与基线不符未合并 " & cFoldBad & " 个" & ListSuffix(lFoldBad)
    If cDiff > 0 Then sb = sb & vbCrLf & "· 同名但内容不同·未重连：" & cDiff & " 个" & ListSuffix(lDiff)
    If cMulti > 0 Then sb = sb & vbCrLf & "· 候选不唯一·未重连：" & cMulti & " 个" & ListSuffix(lMulti)
    If cMiss > 0 Then sb = sb & vbCrLf & "· 未找到：" & cMiss & " 个" & ListSuffix(lMiss) & "（按内容查找仅限相同扩展名；扩展名不同但内容一致的文件请人工核对）"
    If cFail > 0 Then sb = sb & vbCrLf & "· 重连失败：" & cFail & " 个" & ListSuffix(lFail)
    sb = sb & persistNote
    sb = sb & vbCrLf & "建议运行「附件清单 / 体检」核对。"
    RepairCoreInner = sb
    Exit Function
Fail:
    RepairCoreInner = "失败：断链修复时出错 - " & Err.Description
End Function

' 列表后缀："、" 分隔的清单转 "：a、b"；空串转空
Private Function ListSuffix(ByVal s As String) As String
    If Len(s) > 0 Then ListSuffix = "：" & Mid$(s, 2) Else ListSuffix = ""
End Function

' 递归收集目录内全部文件完整路径
Private Sub CollectFilesRecursive(ByVal fso As Object, ByVal folderPath As String, ByVal out As Collection)
    Dim f As Object, sf As Object
    On Error Resume Next
    For Each f In fso.GetFolder(folderPath).Files
        out.Add f.Path
    Next f
    For Each sf In fso.GetFolder(folderPath).SubFolders
        CollectFilesRecursive fso, sf.Path, out
    Next sf
End Sub

'============================ 4. 清除链接 =============================
Public Sub AttachLink_ClearSelection()
    On Error GoTo Fail
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要清除附件链接的单元格区域。", vbInformation, "AttachLink"
        Exit Sub
    End If
    If Selection.Cells.CountLarge > 10000 Then
        MsgBox "选区单元格过多，请选中具体的单元格区域后再清除。", vbInformation, "AttachLink"
        Exit Sub
    End If

    Dim cnt As Long, area As Object, c As Object
    For Each area In Selection.Areas
        For Each c In area.Cells
            cnt = cnt + ClearCellLinks(c)
        Next c
    Next area

    If cnt = 0 Then
        MsgBox "选区内未发现附件链接。", vbInformation, "AttachLink"
    Else
        MsgBox "已清除 " & cnt & " 个附件链接。", vbInformation, "AttachLink"
    End If
    Exit Sub
Fail:
    MsgBox "清除链接时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

'============================ 5. 保存保护 =============================
' 保存前把被 Excel 改写为绝对路径、且确实位于本底稿附件文件夹内的
' 链接归一化回相对路径。由 ThisWorkbook 的 app_WorkbookBeforeSave 调用。
' T8：返回未能归一化的链接数（0=全部成功）；逐链接故障隔离（受保护/只读
' 工作表不中断整批），失败明细记入模块级 normFail（键=工作簿 FullName，
' 无路径时用 Name）；某次保存全部成功时清除该工作簿的记录。
Public Function NormalizeAttachmentLinks(ByVal wb As Object) As Long
    On Error Resume Next
    If wb Is Nothing Then Exit Function
    If Len(wb.Path) = 0 Then Exit Function
    ' N-02：先清理已关闭工作簿的归一化失败记录（否则死记录让 OnTime 被反复调度）
    PruneNormFailRecords
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)
    Dim ws As Object, hl As Object, parentDir As String
    Dim failCnt As Long, det As String, rpErr As String
    For Each ws In wb.Worksheets
        For Each hl In ws.Hyperlinks
            If IsAttachmentLink(hl, fso, wb.Path) Then
                If InStr(hl.Address, ":") > 0 Or Left$(hl.Address, 2) = "\\" Then
                    ' O-03：绝对地址先规范化再比对（大小写/重复分隔符/.. 一律解析）
                    parentDir = fso.GetAbsolutePathName(fso.GetParentFolderName(hl.Address))
                    If StrComp(parentDir, attachDir, vbTextCompare) = 0 Then
                        ' 改指统一走 TryRepoint（与 A4 同款逐链接故障隔离 + 测试钩子单点注入）
                        If Not TryRepoint(hl, ATTACH_FOLDER_NAME & "\" & fso.GetFileName(hl.Address), rpErr) Then
                            failCnt = failCnt + 1
                            If failCnt <= 20 Then
                                det = det & ws.Name & "!" & hl.Range.Address(False, False) & "：" & rpErr & "；"
                            End If
                        End If
                    End If
                End If
            End If
        Next hl
    Next ws
    ' 记录 / 清除该工作簿的失败明细（供状态栏提醒与下次打开时提醒）
    If normFail Is Nothing Then Set normFail = CreateObject("Scripting.Dictionary")
    Dim wbKey As String
    wbKey = NormWbKey(wb)
    If failCnt > 0 Then
        normFail(wbKey) = failCnt & "|" & det
    Else
        If normFail.Exists(wbKey) Then normFail.Remove wbKey
    End If
    NormalizeAttachmentLinks = failCnt
End Function

' 归一化记录的工作簿键：有路径用 FullName，无路径用 Name
Private Function NormWbKey(ByVal wb As Object) As String
    If Len(wb.Path) > 0 Then NormWbKey = wb.FullName Else NormWbKey = wb.Name
End Function

' N-02：清理 normFail 中已关闭工作簿的记录——否则底稿关闭后状态栏提醒仍引用死记录，
' OnTime 被无意义反复调度。键为完整路径时按路径查打开中的工作簿；纯名称键
' （未保存工作簿，关闭后无法重开）同样无法核实，一并清理。
Private Sub PruneNormFailRecords()
    On Error Resume Next
    If normFail Is Nothing Then Exit Sub
    If normFail.Count = 0 Then Exit Sub
    Dim k As Variant
    For Each k In normFail.Keys
        If WorkbookByPath(CStr(k)) Is Nothing Then normFail.Remove k
    Next k
End Sub

' T8：归一化失败的状态栏提醒（全程不弹模态框）。
' 保存完成后由 ThisWorkbook 经 Application.OnTime 调用（无参=汇总全部记录）；
' 打开工作簿时带 wb 调用（只提醒该工作簿）。WPS 不支持状态栏时静默跳过。
Public Sub AttachLink_NotifyNormalizeFailure(Optional ByVal wb As Object)
    On Error Resume Next
    If normFail Is Nothing Then Exit Sub
    PruneNormFailRecords
    If normFail.Count = 0 Then Exit Sub
    Dim failCnt As Long, k As Variant, parts() As String
    If Not wb Is Nothing Then
        Dim kk As String
        kk = NormWbKey(wb)
        If normFail.Exists(kk) Then
            parts = Split(normFail(kk), "|")
            failCnt = CLng(parts(0))
        End If
    Else
        For Each k In normFail.Keys
            parts = Split(normFail(k), "|")
            failCnt = failCnt + CLng(parts(0))
        Next k
    End If
    If failCnt > 0 Then
        Application.StatusBar = "■ AttachLink：上次保存有 " & failCnt & " 个附件链接未能归一化为相对路径（搬移底稿后会断链）——详见「附件清单 / 体检」"
    End If
End Sub

'============================ 1b. 插入文件夹引用 ======================
' 把选定的文件夹整体复制到底稿同级「附件」夹内，并插入指向该文件夹副本的
' 相对链接（点击即在资源管理器打开）。复制语义与文件附件一致（归档自包含）；
' 夹内全部文件自动纳入内容基线（体检可检测夹内篡改）。
Public Sub AttachLink_InsertFolder()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then
        MsgBox "当前底稿尚未保存，无法确定「" & ATTACH_FOLDER_NAME & "」文件夹位置。" & vbCrLf & _
               "请先保存底稿（Ctrl+S）后再插入文件夹引用。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要放置文件夹链接的单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If
    If InStr(1, wb.Path, "://", vbTextCompare) > 0 Then
        MsgBox "底稿保存在网页路径（SharePoint/OneDrive 网页版），相对路径附件链接不可用。", vbExclamation, "AttachLink"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = EnsureAttachDirPrompted(fso, wb, "folder", "本次插入的附件将归档复制到该文件夹")
    If Len(attachDir) = 0 Then Exit Sub

    Dim fd As Object
    Set fd = Application.FileDialog(4)
    With fd
        .Title = "选择要引用的文件夹（将整体复制到底稿附件夹）"
        If Len(lastSourceDir) > 0 Then
            .InitialFileName = lastSourceDir & "\"
        Else
            .InitialFileName = wb.Path & "\"
        End If
        If .Show <> -1 Then Exit Sub
    End With
    Dim srcDir As String
    srcDir = fd.SelectedItems(1)
    ' 规范化（oracle #5 补）：GetFolder().Path 保留磁盘根、去除非根尾部反斜杠；
    ' 不要预先裁剪尾部 \（"C:\" 裁成 "C:" 会被解析为盘当前目录，绕过根目录防护）
    srcDir = fso.GetFolder(srcDir).Path
    lastSourceDir = fso.GetParentFolderName(srcDir)

    ' 路径防护（oracle #4/#5）：根目录防护前置（先于 fName/destDir 计算与覆盖确认弹窗）
    If Len(fso.GetParentFolderName(srcDir)) = 0 Then
        MsgBox "不能引用磁盘根目录。", vbExclamation, "AttachLink"
        Exit Sub
    End If

    Dim fName As String, destDir As String
    fName = fso.GetFileName(srcDir)
    destDir = attachDir & "\" & fName
    ' V1：同名目标不再询问「合并覆盖」——Core 逐文件比对后自动复用或整夹另存 名(2)，
    ' 绝不覆盖附件夹里的既有副本

    If IsSubPathOf(attachDir & "\" & fName, srcDir) Or StrComp(attachDir & "\" & fName, srcDir, vbTextCompare) = 0 Then
        MsgBox "目标位置不能位于源文件夹内部（或与源相同）。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If IsSubPathOf(srcDir, attachDir) Then
        MsgBox "源文件夹已位于本底稿附件夹内，无需引用（可直接建立链接或检查断链）。", vbExclamation, "AttachLink"
        Exit Sub
    End If

    ' 确认覆盖弹窗在 Core 之前；Core 失败（返回串以「失败：」开头）则不插入链接（oracle #3）
    Dim summary As String
    summary = AttachLink_InsertFolderCore(srcDir)
    If Left$(summary, 3) = "失败：" Then
        MsgBox summary, vbCritical, "AttachLink 插入文件夹引用"
        Exit Sub
    End If
    MsgBox summary, vbInformation, "AttachLink 插入文件夹引用"

    ' 链接插入当前选区的第一个可用格（沿用插入的扩展约定：一格一链接）。
    ' V1：Core 决定复用原副本或整夹另存 名(2)，链接地址以汇总里「· 链接指向：」为准
    Dim cellList As Collection
    Set cellList = FlattenSelection(Selection)
    If cellList.Count = 0 Then
        MsgBox "选区内没有可放置链接的单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If
    Dim linkAddr As String, pAt As Long
    linkAddr = ""
    pAt = InStr(summary, "· 链接指向：")
    If pAt > 0 Then
        linkAddr = Mid$(summary, pAt + Len("· 链接指向："))
        pAt = InStr(linkAddr, vbCrLf)
        If pAt > 0 Then linkAddr = Left$(linkAddr, pAt - 1)
        linkAddr = Trim$(linkAddr)
    End If
    If Len(linkAddr) = 0 Then linkAddr = ATTACH_FOLDER_NAME & "\" & fName
    InsertLink cellList(1), linkAddr, fso.GetFileName(linkAddr)
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "插入文件夹引用时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 文件夹引用核心（独立公开函数：便于自动化测试）。以当前活动工作簿所在
' 目录为基准，复制源文件夹到附件夹并把夹内全部文件登记入内容基线（仅登记
' 缺失键，不改既有版本）。V1：同名目标存在时绝不覆盖合并——逐文件三方比对，
' 全一致复用、否则整夹另存 名(2)。返回汇总文本；含「· 链接指向：」行供 UI 解析。
' O-06/V2 重入防护包装：忙碌时直接失败返回；出错兜底——任何出口都复位标志，
' Inner 抛未处理错误如实转「失败：<过程名> - <原因>」
Public Function AttachLink_InsertFolderCore(ByVal srcDir As String) As String
    If coreBusy Then
        AttachLink_InsertFolderCore = "失败：已有操作正在进行。"
        Exit Function
    End If
    coreBusy = True
    Err.Clear
    On Error Resume Next
    AttachLink_InsertFolderCore = InsertFolderCoreInner(srcDir)
    If Err.Number <> 0 Then
        AttachLink_InsertFolderCore = "失败：AttachLink_InsertFolderCore - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    coreBusy = False
End Function

Private Function InsertFolderCoreInner(ByVal srcDir As String) As String
    On Error GoTo Fail
    Dim wbHost As Object
    Set wbHost = ActiveWorkbook
    If wbHost Is Nothing Then
        InsertFolderCoreInner = "失败：没有活动工作簿。"
        Exit Function
    End If
    If Len(wbHost.Path) = 0 Then
        InsertFolderCoreInner = "失败：活动工作簿尚未保存。"
        Exit Function
    End If
    ' V10：网页路径底稿（SharePoint/OneDrive）下相对链接不可用
    If IsWebPathWorkbook(wbHost) Then
        InsertFolderCoreInner = "失败：底稿保存在网页路径（SharePoint/OneDrive 网页版），无法插入文件夹引用。请先把底稿另存到本地磁盘或局域网共享盘。"
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcDir) Then
        InsertFolderCoreInner = "失败：源文件夹不存在：" & srcDir
        Exit Function
    End If
    srcDir = fso.GetFolder(srcDir).Path                ' oracle #5：Core 直调同样规范化（含大小写归一）

    Dim attachDir As String, fName As String, destDir As String
    attachDir = wbHost.Path & "\" & ATTACH_FOLDER_NAME
    fName = fso.GetFileName(srcDir)
    destDir = attachDir & "\" & fName
    ' 路径防护（oracle #4，Core 直调同样校验）
    If Len(fso.GetParentFolderName(srcDir)) = 0 Then
        InsertFolderCoreInner = "失败：不能引用磁盘根目录。"
        Exit Function
    End If
    If IsSubPathOf(destDir, srcDir) Or StrComp(destDir, srcDir, vbTextCompare) = 0 Then
        InsertFolderCoreInner = "失败：目标位置不能位于源文件夹内部（或与源相同）。"
        Exit Function
    End If
    If IsSubPathOf(srcDir, attachDir) Then
        InsertFolderCoreInner = "失败：源文件夹已位于本底稿附件夹内，无需引用。"
        Exit Function
    End If
    If Not EnsureFolder(fso, attachDir) Then
        InsertFolderCoreInner = "失败：无法创建附件文件夹：" & attachDir
        Exit Function
    End If
    ' V7：夹内任一文件名含 # → Excel 超链接无法正确保存，整批拒绝（复制之前）
    Dim srcAll As Collection
    Set srcAll = New Collection
    CollectFilesRecursive fso, srcDir, srcAll
    Dim hashCnt As Long, hashList As String, fChk As Variant
    For Each fChk In srcAll
        If InStr(fso.GetFileName(CStr(fChk)), "#") > 0 Then
            hashCnt = hashCnt + 1
            If hashCnt <= 10 Then hashList = hashList & "、" & fso.GetFileName(CStr(fChk))
        End If
    Next fChk
    If hashCnt > 0 Then
        InsertFolderCoreInner = "失败：以下文件名含有 #，Excel 超链接无法正确保存，请重命名后再插入：" & Mid$(hashList, 2)
        Exit Function
    End If
    ' O-04：基线损坏 = 失败关闭——在任何复制之前返回（统一解析器区分损坏/只读/缺失）
    ResetHashCache
    Dim base As Object, baseStatus As Long, baseRO As Boolean
    Set base = LoadBaselineEx(fso, attachDir, baseStatus, baseRO)
    If baseStatus = 2 Then
        InsertFolderCoreInner = "失败：基线文件损坏或无法读取（" & lastBaselineError & _
                                "），为避免绕过附件身份校验，本次未复制文件夹。请先人工核对或重建基线。"
        Exit Function
    End If

    ' V1：同名目标存在时绝不覆盖合并（与单文件「同名同内容复用、异内容另存」同规则）——
    ' 逐文件三方比对（源↔副本、副本↔基线）：全部一致 → 复用现有副本；任一不符 /
    ' 副本缺失 / 哈希不可用 / 同名路径是文件 → 整夹另存 名(2)（现有副本任何文件不得改动）
    Dim targetDir As String, foldNote As String, extraCnt As Long
    Dim regCnt As Long, fileCnt As Long, failCnt As Long
    targetDir = destDir
    If fso.FileExists(destDir) Or fso.FolderExists(destDir) Then
        If fso.FolderExists(destDir) And FolderCopiesMatch(fso, base, srcAll, attachDir, srcDir, destDir, extraCnt) Then
            foldNote = "附件夹内已有同名文件夹且内容一致，已直接复用（未复制）。"
            If extraCnt > 0 Then
                foldNote = foldNote & "副本里另有 " & extraCnt & " 个非本次来源的文件，未改动。"
            End If
            ' 复用：不复制、不重复登记，链接指向现有副本
            InsertFolderCoreInner = "文件夹引用完成：" & vbCrLf & _
                "· " & foldNote & vbCrLf & _
                "· 链接指向：" & ATTACH_FOLDER_NAME & "\" & fName & vbCrLf & _
                "· 点击即在资源管理器打开"
            Exit Function
        End If
        targetDir = NextFolderCopyName(fso, attachDir, fName)
        If Len(targetDir) = 0 Then
            InsertFolderCoreInner = "失败：同名文件夹已存在且内容不同，另存名已用尽（1000 个候选均占用），本次未复制任何文件。"
            Exit Function
        End If
        foldNote = "附件夹内已有同名文件夹但内容不同（或副本与其内容基线不一致），本次未覆盖，已另存为 " & _
                   fso.GetFileName(targetDir)
    End If
    fso.CopyFolder srcDir, targetDir

    ' 夹内全部文件登记基线（仅补缺失键；已有键的版本与哈希不动）。
    ' 登记失败（哈希/遍历错误）时如实报告部分完成（oracle #2）
    RegisterFolderFilesBaseline fso, base, targetDir, attachDir, regCnt, fileCnt, failCnt
    If Not SaveBaseline(fso, attachDir, base) Then
        InsertFolderCoreInner = "失败：基线文件写入不成功（文件夹已复制，但登记未持久化，" & _
            BaselinePersistNote(lastBaselineError) & "）：" & _
            lastBaselineError & vbCrLf & attachDir & "\" & BASELINE_FILE
        Exit Function
    End If

    Dim tail As String
    If failCnt > 0 Then tail = vbCrLf & "■ 有 " & failCnt & " 个文件登记失败（不可读/被占用），体检将显示「校验不可用」"
    Dim noteLine As String
    If Len(foldNote) > 0 Then noteLine = "· " & foldNote & vbCrLf
    InsertFolderCoreInner = "文件夹引用完成：" & vbCrLf & noteLine & _
        "· 已复制到 " & targetDir & "（共 " & fileCnt & " 个文件）" & vbCrLf & _
        "· 新登记内容基线 " & regCnt & " 个文件（夹内篡改可被体检检测）" & tail & vbCrLf & _
        "· 链接指向：" & ATTACH_FOLDER_NAME & "\" & fso.GetFileName(targetDir) & vbCrLf & _
        "· 点击即在资源管理器打开"
    Exit Function
Fail:
    InsertFolderCoreInner = "失败：插入文件夹引用未完成 - " & Err.Description
End Function

' V1：文件夹引用复用判定——源与既有副本逐文件三方比对：每个源文件在副本同相对
' 路径存在、哈希相同，且（已有基线键时）副本哈希与基线一致；任一缺失/不一致/
' 哈希不可用即 False（保守不复用）。副本多出文件不算冲突，数量经 extraCnt 输出
' 供汇总如实说明。srcFiles 为源目录递归文件集合（调用方已收集）。
Private Function FolderCopiesMatch(ByVal fso As Object, ByVal base As Object, _
        ByVal srcFiles As Collection, ByVal attachDir As String, _
        ByVal srcRoot As String, ByVal locRoot As String, _
        ByRef extraCnt As Long) As Boolean
    Dim locAll As Collection, f As Variant, rel As String, locPath As String
    Dim hS As String, hL As String, key As String
    Set locAll = New Collection
    CollectFilesRecursive fso, locRoot, locAll
    extraCnt = locAll.Count - srcFiles.Count
    If srcFiles.Count = 0 Then Exit Function          ' 空源文件夹不复用，走另存
    For Each f In srcFiles
        rel = Mid$(CStr(f), Len(srcRoot) + 2)
        locPath = locRoot & "\" & rel
        If Not fso.FileExists(locPath) Then Exit Function
        hS = HashFileCached(CStr(f))
        hL = HashFileCached(locPath)
        If Len(hS) = 0 Or Len(hL) = 0 Then Exit Function
        If StrComp(hS, hL, vbTextCompare) <> 0 Then Exit Function
        key = LCase$(Mid$(locPath, Len(attachDir) + 2))
        If BaselineMismatch(base, key, hL) Then Exit Function
    Next f
    FolderCopiesMatch = True
End Function

' V1：文件夹整夹另存的下一个可用名（名(2)、名(3)…，与单文件另存同规则）；
' 同名文件或文件夹均视为占用。用尽返回 ""。
Private Function NextFolderCopyName(ByVal fso As Object, ByVal attachDir As String, ByVal fName As String) As String
    Dim n As Long, cand As String
    n = 2
    Do
        cand = attachDir & "\" & fName & "(" & n & ")"
        If Not fso.FileExists(cand) And Not fso.FolderExists(cand) Then
            NextFolderCopyName = cand
            Exit Function
        End If
        n = n + 1
    Loop While n < 1000
    NextFolderCopyName = ""
End Function

' 递归登记文件夹内全部文件到基线（仅缺失键）；同时统计文件总数与失败数（哈希/遍历错误）
Private Sub RegisterFolderFilesBaseline(ByVal fso As Object, ByVal base As Object, _
                                        ByVal folderPath As String, ByVal attachDir As String, _
                                        ByRef regCnt As Long, ByRef fileCnt As Long, ByRef failCnt As Long)
    Dim f As Object, sf As Object, key As String, h As String
    On Error Resume Next
    For Each f In fso.GetFolder(folderPath).Files
        fileCnt = fileCnt + 1
        key = LCase$(Mid$(f.Path, Len(attachDir) + 2))
        If Not base.Exists(key) Then
            h = HashFileCached(f.Path)
            If Len(h) > 0 Then
                base.Add key, h & "|1|" & Format$(Now, "yyyy-mm-dd hh:nn") & "|" & Format$(Now, "yyyy-mm-dd hh:nn")
                regCnt = regCnt + 1
            Else
                failCnt = failCnt + 1
            End If
        End If
    Next f
    If Err.Number <> 0 Then failCnt = failCnt + 1: Err.Clear
    For Each sf In fso.GetFolder(folderPath).SubFolders
        RegisterFolderFilesBaseline fso, base, sf.Path, attachDir, regCnt, fileCnt, failCnt
    Next sf
    If Err.Number <> 0 Then failCnt = failCnt + 1: Err.Clear
End Sub

' 子路径边界判断（大小写不敏感）：child 等于 parent 或位于 parent 之内
Private Function IsSubPathOf(ByVal child As String, ByVal parent As String) As Boolean
    Dim c As String, p As String
    c = LCase$(child)
    p = LCase$(parent)
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    IsSubPathOf = (c = p) Or (Len(c) > Len(p) And Left$(c, Len(p) + 1) = p & "\")
End Function

'============================ 1c. 重命名链接显示名 ====================
' 只改超链接的显示文本（应对收到的文件命名不规范/不清晰），不改动
' 底层文件名与链接地址——体检、哈希基线、归档迁移均不受影响。
' 支持多选批量：选区内全部附件链接按行序逐个弹窗（标题含进度 i/N），
' 留空或取消即停止后续（已改的保留），结束后汇总。
Public Sub AttachLink_RenameLink()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要重命名的附件链接单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 收集选区内全部附件链接（行序）。
    ' 超大选区（Ctrl+A 全选/整行整列）自动与已用区域交叠限域——否则逐格 COM 迭代
    ' 亿级单元格形同冻结（2026-11-21 用户实测「全选后只能单个修改」的根因）
    Dim links As Collection
    Set links = New Collection
    Dim c As Object, hl As Object, scanRange As Object
    Set scanRange = GetScanRange()
    If scanRange Is Nothing Then
        MsgBox "选区内未发现附件链接。", vbInformation, "AttachLink"
        Exit Sub
    End If
    For Each c In scanRange.Cells
        For Each hl In c.Hyperlinks
            If IsAttachmentLink(hl, fso, wb.Path) Then links.Add hl
        Next hl
    Next c
    If links.Count = 0 Then
        MsgBox "选区内未发现附件链接。", vbInformation, "AttachLink"
        Exit Sub
    End If

    Dim i As Long, newName As String, doneCnt As Long, stopped As Boolean
    For i = 1 To links.Count
        newName = InputBox("重命名附件（" & i & "/" & links.Count & "）：留空或点「取消」停止批量。" & vbCrLf & _
                           "（仅改变链接显示文本，不改动文件本身）", _
                           "AttachLink 重命名链接显示名", links(i).TextToDisplay)
        newName = Trim$(newName)
        If Len(newName) = 0 Then
            stopped = True
            Exit For
        End If
        If StrComp(newName, links(i).TextToDisplay, vbBinaryCompare) <> 0 Then
            links(i).TextToDisplay = newName
            doneCnt = doneCnt + 1
        End If
    Next i

    If links.Count = 1 And doneCnt = 0 And Not stopped Then
        ' 单选未改动：无需提示
    ElseIf stopped And i = 1 Then
        ' 第一个就取消：静默退出
    Else
        MsgBox "批量重命名结束：已改名 " & doneCnt & " 个（共 " & links.Count & " 个" & _
               IIf(stopped, "，在第 " & i & " 个停止", "") & "）。", vbInformation, "AttachLink"
    End If
    Exit Sub
Fail:
    MsgBox "重命名时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 收集/扫描范围解析：Selection 过大（全选/整行/整列，CountLarge>65536）时与
' 工作表已用区域交叠，避免逐格迭代亿级单元格；交叠为空返回 Nothing
Private Function GetScanRange() As Object
    On Error Resume Next
    Dim sel As Object, ws As Object
    Set sel = Selection
    Set ws = sel.Worksheet
    If sel.Cells.CountLarge > 65536 Then
        Set GetScanRange = Application.Intersect(sel, ws.UsedRange)
    Else
        Set GetScanRange = sel
    End If
End Function

' [诊断保留 2026-11-21] 批量重命名核心机制复现（含全选限域回归；下版随测试入仓可移除）
Public Function AttachLink_RenameDiag(ByVal newNamePrefix As String) As String
    Dim out As String
    On Error GoTo Fail
    Dim wb As Object: Set wb = ActiveWorkbook
    If TypeName(Selection) <> "Range" Then AttachLink_RenameDiag = "sel=" & TypeName(Selection): Exit Function
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim links As Collection: Set links = New Collection
    Dim c As Object, hl As Object
    For Each c In GetScanRange().Cells
        For Each hl In c.Hyperlinks
            If IsAttachmentLink(hl, fso, wb.Path) Then links.Add hl
        Next hl
    Next c
    out = "collected=" & links.Count & ";"
    Dim i As Long
    For i = 1 To links.Count
        links(i).TextToDisplay = newNamePrefix & i
        out = out & "r" & i & "=ok;"
    Next i
    AttachLink_RenameDiag = out
    Exit Function
Fail:
    AttachLink_RenameDiag = out & "ERR=" & Err.Number & ":" & Err.Description
End Function

'============================ 1d. 附件删除 ============================
' 删除选区内全部附件链接，并弹窗确认后连带删除附件夹内对应的文件/文件夹
' 与基线条目。与「清除选区链接」（仅删链接不删文件）互为补充。
Public Sub AttachLink_DeleteAttachment()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then
        MsgBox "当前底稿尚未保存，无法定位附件文件夹。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    ' V10：网页路径底稿无法定位本地附件夹，直接拒绝
    If IsWebPathWorkbook(wb) Then
        MsgBox "底稿保存在网页路径（SharePoint/OneDrive 网页版），无法执行附件删除。" & vbCrLf & _
               "请先把底稿另存到本地磁盘或局域网共享盘后再执行。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要删除的附件链接单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)

    ' 收集字符串引用，避免持有 Hyperlink COM 对象跨越文件删除操作。
    ' links 与 paths 始终保持相同索引。
    Dim links As Collection, paths As Collection
    Set links = New Collection
    Set paths = New Collection
    Dim c As Object, hl As Object, scanR As Object, fullPath As String
    Set scanR = GetScanRange()
    If Not scanR Is Nothing Then
        For Each c In scanR.Cells
            For Each hl In c.Hyperlinks
                If IsAttachmentLink(hl, fso, wb.Path) Then
                    fullPath = fso.GetAbsolutePathName(ResolveLinkPath(fso, wb.Path, hl.Address))
                    ' 删除是破坏性操作：只允许处理当前底稿附件夹内的目标。
                    ' 同时拒绝附件夹根本身（IsSubPathOf 含等值，需另行排除）。
                    If Not IsSubPathOf(fullPath, attachDir) Or StrComp(fullPath, attachDir, vbTextCompare) = 0 Then
                        MsgBox "选区包含指向当前附件文件夹之外的链接，已停止操作：" & vbCrLf & fullPath & vbCrLf & vbCrLf & _
                               "请先执行归档迁移或断链修复；仅想移除链接时请用「清除选区链接」。", _
                               vbExclamation, "AttachLink"
                        Exit Sub
                    End If
                    links.Add hl.Range.Address(False, False) & "|" & hl.Range.Worksheet.Name
                    paths.Add fullPath
                End If
            Next hl
        Next c
    End If
    If links.Count = 0 Then
        MsgBox "选区内未发现附件链接。", vbInformation, "AttachLink"
        Exit Sub
    End If

    ' 构建确认清单
    Dim fileList As String, i As Long, p As Variant
    Dim fileCnt As Long, folderCnt As Long, missCnt As Long
    For Each p In paths
        If fso.FileExists(p) Then
            fileList = fileList & "  [文件] " & fso.GetFileName(p) & vbCrLf
            fileCnt = fileCnt + 1
        ElseIf fso.FolderExists(p) Then
            fileList = fileList & "  [文件夹] " & fso.GetFileName(p) & "\（含全部内容）" & vbCrLf
            folderCnt = folderCnt + 1
        Else
            fileList = fileList & "  [断链] " & fso.GetFileName(p) & "（文件不存在，仅删链接）" & vbCrLf
            missCnt = missCnt + 1
        End If
    Next p

    If MsgBox("附件删除确认（共 " & links.Count & " 个链接）：" & vbCrLf & vbCrLf & _
              fileList & vbCrLf & _
              "将执行：" & vbCrLf & _
              "  · 删除选区内 " & links.Count & " 个附件链接" & vbCrLf & _
              "  · 删除附件夹内 " & fileCnt & " 个文件、" & folderCnt & " 个文件夹" & vbCrLf & _
              "  · 清除对应基线条目" & vbCrLf & vbCrLf & _
              "此操作不可撤销。", vbOKCancel + vbExclamation, "AttachLink 附件删除") <> vbOK Then Exit Sub

    ' 执行交给 Core（无对话框，供自动化测试直调）：先删链接并确认成功，再删
    ' 文件/文件夹（链接删不掉的条目文件保留），最后清基线
    Dim linksJoined As String, pathsJoined As String
    For i = 1 To links.Count
        If i > 1 Then
            linksJoined = linksJoined & vbLf
            pathsJoined = pathsJoined & vbLf
        End If
        linksJoined = linksJoined & CStr(links(i))
        pathsJoined = pathsJoined & CStr(paths(i))
    Next i

    Dim summary As String, msgIcon As Long
    summary = AttachLink_DeleteAttachmentCore(wb, linksJoined, pathsJoined, fso)
    ' V4：Core 发现同一文件还有范围外链接在引用时会返回「需确认：…」——
    ' 列进确认框，用户确认后以 confirmShared=True 再执行
    If Left$(summary, 4) = "需确认：" Then
        If MsgBox(summary & vbCrLf & vbCrLf & _
                  "仍要删除该文件与选中的链接吗？", _
                  vbOKCancel + vbExclamation, "AttachLink 附件删除") <> vbOK Then Exit Sub
        summary = AttachLink_DeleteAttachmentCore(wb, linksJoined, pathsJoined, fso, True)
    End If
    If InStr(summary, "部分失败") > 0 Or InStr(summary, "失败：") > 0 Then msgIcon = vbExclamation Else msgIcon = vbInformation
    MsgBox summary, msgIcon, "AttachLink"
    Exit Sub
Fail:
    MsgBox "附件删除时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 删除执行核心（无对话框，供自动化测试直调）：删链接清样式（确认成功）→ 删
' 文件/文件夹（链接未删成的条目文件保留，V8b）→ 清基线，
' 返回汇总文本（含「部分失败」/「失败：」字样时 wrapper 以警示图标显示）。
' linksJoined / pathsJoined 为按行拼接的 "地址|表名" / 完整路径串（vbLf 分隔，两侧一一对应）。
' 安全契约（oracle 复核 #1）：不信任调用方传入的路径——附件夹由 wb 自身推导，
' 全部路径先经一次性预检（附件夹内、非根、无控制字符、数据一致），任一项
' 不合法即整体拒绝、零删除；首项合法后项越界同样不得删除首项。
' O-06/V2 重入防护包装：忙碌时直接失败返回；出错兜底——任何出口都复位标志。
' V4：新增可选参数 confirmShared——同一文件还有范围外链接引用时必须经界面
' 二次确认（True）才允许删除（防共享文件被静默删成多处断链）。
' linksJoined / pathsJoined 为按行拼接的 "地址|表名" / 完整路径串（vbLf 分隔，两侧一一对应）。
' 安全契约（oracle 复核 #1）：不信任调用方传入的路径——附件夹由 wb 自身推导，
' 全部路径先经一次性预检（附件夹内、非根、无控制字符、数据一致），任一项
' 不合法即整体拒绝、零删除；首项合法后项越界同样不得删除首项。
Public Function AttachLink_DeleteAttachmentCore(ByVal wb As Object, ByVal linksJoined As String, _
    ByVal pathsJoined As String, Optional ByVal fsoIn As Object, _
    Optional ByVal confirmShared As Boolean = False) As String
    If coreBusy Then
        AttachLink_DeleteAttachmentCore = "失败：已有操作正在进行。"
        Exit Function
    End If
    coreBusy = True
    Err.Clear
    On Error Resume Next
    AttachLink_DeleteAttachmentCore = DeleteAttachmentCoreInner(wb, linksJoined, pathsJoined, fsoIn, confirmShared)
    If Err.Number <> 0 Then
        AttachLink_DeleteAttachmentCore = "失败：AttachLink_DeleteAttachmentCore - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    coreBusy = False
End Function

Private Function DeleteAttachmentCoreInner(ByVal wb As Object, ByVal linksJoined As String, _
    ByVal pathsJoined As String, Optional ByVal fsoIn As Object, _
    Optional ByVal confirmShared As Boolean = False) As String
    On Error GoTo Fail

    Dim fso As Object
    If fsoIn Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    Else
        Set fso = fsoIn
    End If

    If wb Is Nothing Then
        DeleteAttachmentCoreInner = "失败：未提供工作簿，未执行任何删除。"
        Exit Function
    End If
    If Len(wb.Path) = 0 Then
        DeleteAttachmentCoreInner = "失败：底稿未保存，无法定位附件文件夹，未执行任何删除。"
        Exit Function
    End If
    ' V10：网页路径底稿（SharePoint/OneDrive「在桌面应用中打开」，wb.Path 以 https://
    ' 开头）下无法定位本地附件夹，直接拒绝，避免全部链接被误判断链后清空清单
    If IsWebPathWorkbook(wb) Then
        DeleteAttachmentCoreInner = "失败：底稿保存在网页路径（SharePoint/OneDrive 网页版），无法执行附件删除。请先把底稿另存到本地磁盘或局域网共享盘。"
        Exit Function
    End If

    ' 附件夹从 wb 自推导（不信任调用方），GetAbsolutePathName 规范化防 ../ 与重复分隔符
    Dim attachDir As String
    attachDir = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)

    Dim linkArr() As String, pathArr() As String
    linkArr = Split(linksJoined, vbLf)
    pathArr = Split(pathsJoined, vbLf)
    If UBound(linkArr) < 0 Or UBound(linkArr) <> UBound(pathArr) Then
        DeleteAttachmentCoreInner = "失败：删除内部数据不一致，未执行任何删除。"
        Exit Function
    End If

    ' ---- 预检：全部通过才开始删除；任一项不合法 → 整体拒绝、零删除 ----
    Dim i As Long, p As Variant, norm As String
    For i = 0 To UBound(pathArr)
        p = pathArr(i)
        If Len(CStr(p)) = 0 Then
            DeleteAttachmentCoreInner = "失败：第 " & (i + 1) & " 项路径为空，未执行任何删除。"
            Exit Function
        End If
        ' 拒绝换行/NUL 注入（不用 Asc 逐字符扫描：中文字符的 ANSI 码转换会假阳性）
        If InStr(CStr(p), vbCr) > 0 Or InStr(CStr(p), vbLf) > 0 Or InStr(CStr(p), Chr$(0)) > 0 Then
            DeleteAttachmentCoreInner = "失败：第 " & (i + 1) & " 项路径含换行/NUL 字符，未执行任何删除。"
            Exit Function
        End If
        norm = fso.GetAbsolutePathName(CStr(p))
        If Not IsSubPathOf(norm, attachDir) Or StrComp(norm, attachDir, vbTextCompare) = 0 Then
            DeleteAttachmentCoreInner = "失败：第 " & (i + 1) & " 项越界（不在当前底稿附件文件夹内或为附件夹根），未执行任何删除：" & vbCrLf & norm
            Exit Function
        End If
        pathArr(i) = norm
    Next i
    For i = 0 To UBound(linkArr)
        If InStr(CStr(linkArr(i)), "|") = 0 Then
            DeleteAttachmentCoreInner = "失败：第 " & (i + 1) & " 项链接引用格式非法，未执行任何删除。"
            Exit Function
        End If
    Next i

    ' V4：共享文件二次确认——同内容复用让多处链接指向同一文件成为常态；删除之前
    ' 扫描全部工作表（清单表除外），找出「指向本次目标文件但不在删除范围内」的链接。
    ' 存在且未经 confirmShared 确认 → 不删除，如实返回需确认清单（界面列进确认框）
    If Not confirmShared Then
        Dim sharedCnt As Long, sharedList As String
        Dim wsX As Object, hlX As Object, fullX As String, refX As String
        For Each wsX In wb.Worksheets
            If StrComp(wsX.Name, INDEX_SHEET_NAME, vbTextCompare) <> 0 Then
                For Each hlX In wsX.Hyperlinks
                    If IsAttachmentLink(hlX, fso, wb.Path) Then
                        fullX = fso.GetAbsolutePathName(ResolveLinkPath(fso, wb.Path, hlX.Address))
                        If IsSubPathOf(fullX, attachDir) And StrComp(fullX, attachDir, vbTextCompare) <> 0 Then
                            If PathInArray(fullX, pathArr) Then
                                refX = hlX.Range.Address(False, False) & "|" & wsX.Name
                                If Not RefInArray(refX, linkArr) Then
                                    sharedCnt = sharedCnt + 1
                                    If sharedCnt <= 10 Then
                                        sharedList = sharedList & vbCrLf & "· " & wsX.Name & "!" & hlX.Range.Address(False, False)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next hlX
            End If
        Next wsX
        If sharedCnt > 0 Then
            Dim sharedMore As String
            If sharedCnt > 10 Then sharedMore = vbCrLf & "· ……另有 " & (sharedCnt - 10) & " 处（未逐一列出）"
            DeleteAttachmentCoreInner = "需确认：要删除的文件还被其它 " & sharedCnt & " 处链接引用：" & vbCrLf & _
                Mid$(sharedList, 3) & sharedMore & vbCrLf & _
                "删除文件后这些链接会断开（日后可用「附件断链修复」找回）。" & vbCrLf & _
                "如只想移除当前选中的链接、保留文件，请点「取消」。"
            Exit Function
        End If
    End If

    Dim delCnt As Long, delFail As Long
    Dim linkDelCnt As Long, linkDelFail As Long
    Dim cleanFail As Long
    Dim missCnt As Long
    Dim dupCnt As Long      ' 同一次删除里，后面的链接指向前面已删掉的同一文件（不记成「本就不存在」）

    ' V8b：先删链接并确认成功，再删文件——链接删不掉（表受保护）时该条文件必须
    ' 保留，绝不再现「文件已永久删除、链接还在」的半截状态
    Dim linkOkArr() As Boolean
    ReDim linkOkArr(0 To UBound(linkArr))
    Dim delPaths As Collection
    Set delPaths = New Collection
    Dim protCnt As Long, protList As String

    ' 1) 删链接 + 清内容（保留单元格边框/填充格式；清除超链接字体样式和文本）
    '    字符串引用 "地址|表名" 在此解析重定位（Split 限 2 段，表名含 | 不受影响）
    Dim cellRef As Object, wsRef As Object
    Dim refParts() As String
    Dim addr As String, sheetName As String
    Dim remaining As Long, itemErr As Long, cleanErr As Long
    For i = UBound(linkArr) To 0 Step -1
        Set cellRef = Nothing: Set wsRef = Nothing
        Erase refParts
        addr = "": sheetName = ""
        remaining = -1: itemErr = 0: cleanErr = 0
        linkOkArr(i) = False
        On Error Resume Next
        refParts = Split(CStr(linkArr(i)), "|", 2)
        If UBound(refParts) = 1 Then
            addr = refParts(0)
            sheetName = refParts(1)
            ' 直接按表名取 Worksheet 对象再 Range(addr)，不拼接含表名的字符串，
            ' 表名中的空格/单引号等特殊字符无需转义
            Set wsRef = wb.Worksheets(sheetName)
            If Not wsRef Is Nothing Then Set cellRef = wsRef.Range(addr)
        End If
        If cellRef Is Nothing Then
            linkDelFail = linkDelFail + 1
        Else
            ' 删链接：ClearHyperlinks 是唯一「删链接且保留边框/填充」的方法——
            ' 实测（2026-11-23，Excel 16.0）Range.Hyperlinks.Delete 与单个
            ' Hyperlink.Delete 均会把单元格边框/填充重置为无格式；ClearHyperlinks 保留。
            ' 宿主若无此方法（老版本）则回退 Hyperlinks.Delete（牺牲格式保功能）。
            Err.Clear
            cellRef.ClearHyperlinks
            itemErr = Err.Number
            If itemErr <> 0 Then
                Err.Clear
                cellRef.Hyperlinks.Delete
                itemErr = Err.Number
            End If
            ' 无异常不代表链接确实消失，复核集合
            If itemErr = 0 Then
                remaining = cellRef.Hyperlinks.Count
                itemErr = Err.Number
            End If
            If itemErr <> 0 Or remaining <> 0 Then
                linkDelFail = linkDelFail + 1
            Else
                linkDelCnt = linkDelCnt + 1
                linkOkArr(i) = True
                ' 用 Value="" 而非 ClearContents（为规避宿主兼容差异沿用 Value=""）；
                ' 每步单独记录错误，避免后续成功调用掩盖前一步失败
                Err.Clear
                cellRef.Value = ""
                If Err.Number <> 0 Then cleanErr = Err.Number
                Err.Clear
                cellRef.Font.Underline = -4142
                If Err.Number <> 0 And cleanErr = 0 Then cleanErr = Err.Number
                Err.Clear
                cellRef.Font.ColorIndex = -4105
                If Err.Number <> 0 And cleanErr = 0 Then cleanErr = Err.Number
                Err.Clear
                cellRef.Font.Bold = False
                If Err.Number <> 0 Then cleanErr = Err.Number
                If cleanErr <> 0 Then cleanFail = cleanFail + 1
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next i

    ' 2) 删除文件/文件夹——仅链接已确认删除的条目才动文件（missCnt 在此自行统计，
    '    不信任调用方预计算）；链接未删成的条目文件保留，汇总如实说明
    For i = 0 To UBound(pathArr)
        If Not linkOkArr(i) Then
            protCnt = protCnt + 1
            If protCnt <= 10 Then protList = protList & "、" & fso.GetFileName(CStr(pathArr(i)))
            GoTo NextDelPath
        End If
        p = pathArr(i)
        Err.Clear
        On Error Resume Next
        If fso.FileExists(p) Then
            Err.Clear
            fso.DeleteFile p, True
            If Err.Number = 0 Then
                delCnt = delCnt + 1
                delPaths.Add CStr(p)
            Else
                delFail = delFail + 1: Err.Clear
            End If
        ElseIf fso.FolderExists(p) Then
            Err.Clear
            fso.DeleteFolder p, True
            If Err.Number = 0 Then
                delCnt = delCnt + 1
                delPaths.Add CStr(p)
            Else
                delFail = delFail + 1: Err.Clear
            End If
        Else
            If PathInDeleted(p, delPaths) Then
                ' 多处链接指向同一文件：前面的条目已把它删掉，那是本次删除造成的，不能记成「本就不存在」
                dupCnt = dupCnt + 1
            Else
                ' 断链：目标本就不存在，不算文件删除失败
                missCnt = missCnt + 1
            End If
        End If
        Err.Clear
        On Error GoTo 0
NextDelPath:
    Next i

    ' 3) 清基线：文件删自身键；文件夹连带删除 "文件夹\..." 全部后代键
    '    严格加载（oracle 复核 #3）：区分「基线文件不存在」与「读取失败」——
    '    读取失败必须如实报 baseReadFail，且绝不把部分字典回写（防覆盖丢键）；
    '    写回失败单独报 baseWriteFail（写回失败不保证文件未改动，措辞区分）
    '    V8b：只清理确实已删除的路径（链接未删成的条目键保留）
    Dim base As Object, key As Variant
    Dim baseStatus As Long
    Dim baseReadFail As Boolean, baseWriteFail As Boolean
    Set base = LoadBaselineStrict(fso, attachDir, baseStatus)
    If baseStatus = 2 Then
        baseReadFail = True
    ElseIf base.Count > 0 Then
        Dim newBase As Object
        Set newBase = CreateObject("Scripting.Dictionary")
        Dim found As Boolean, rel As String
        For Each key In base.Keys
            found = False
            For Each p In delPaths
                rel = LCase$(Mid$(CStr(p), Len(attachDir) + 2))
                If StrComp(CStr(key), rel, vbTextCompare) = 0 Then
                    found = True
                ElseIf StrComp(Left$(CStr(key), Len(rel) + 1), rel & "\", vbTextCompare) = 0 Then
                    found = True
                End If
                If found Then Exit For
            Next p
            If Not found Then newBase.Add key, base(key)
        Next key
        If Not SaveBaseline(fso, attachDir, newBase) Then baseWriteFail = True
    End If

    ' 4) 如实汇总：不以链接总数冒充链接成功数
    Dim msg As String
    If delFail > 0 Or linkDelFail > 0 Or cleanFail > 0 Or baseReadFail Or baseWriteFail Or protCnt > 0 Then
        msg = "附件删除结束（部分失败）："
    Else
        msg = "附件删除完成："
    End If
    msg = msg & "已删除 " & delCnt & " 个文件/文件夹、" & linkDelCnt & " 个链接"
    If dupCnt > 0 Then msg = msg & "（其中 " & dupCnt & " 处链接指向同一文件，文件只删除一次）"
    If delFail > 0 Then msg = msg & "，" & delFail & " 个文件/文件夹删除失败"
    If linkDelFail > 0 Then msg = msg & "，" & linkDelFail & " 个链接删除失败"
    If cleanFail > 0 Then msg = msg & "，" & cleanFail & " 个单元格内容/样式清理失败"
    If baseReadFail Then msg = msg & "，基线读取失败（文件未改动，未执行基线清理，请人工核对）"
    If baseWriteFail Then msg = msg & "，基线清理结果保存失败（请人工核对）"
    If missCnt > 0 Then msg = msg & "，" & missCnt & " 个链接目标本就不存在"
    If protCnt > 0 Then
        msg = msg & "，" & protCnt & " 个工作表受保护，链接未能删除，文件已保留（" & Mid$(protList, 2) & "）"
    End If
    DeleteAttachmentCoreInner = msg
    Exit Function
Fail:
    DeleteAttachmentCoreInner = "失败：删除附件时出错 - " & Err.Description
End Function

Private Function PathInDeleted(ByVal p As String, ByVal col As Collection) As Boolean
    ' 本次删除已删掉的路径里是否有 p（不区分大小写）。多处链接指向同一文件时，后处理的链接
    ' 再看该路径已不存在——那是本次删除造成的，不能记成「链接目标本就不存在」
    Dim it As Variant
    If col Is Nothing Then Exit Function
    For Each it In col
        If StrComp(CStr(it), p, vbTextCompare) = 0 Then
            PathInDeleted = True
            Exit Function
        End If
    Next it
End Function

'============================ 1e. 底稿引用 =============================
' 在单元格中插入指向同项目其他底稿**指定工作表**的超链接。
' 目标底稿复制入附件夹（自包含原则），链接指向附件夹内副本。
' 用 Workbooks.Open(ReadOnly) 读取工作表列表供选择（防选错工作表名导致
' 点击链接时 Excel 卡死）。打开目标时加三重防护：
'   EnableEvents=False（防目标宏运行）+ UpdateLinks=0（防外部链接刷新）
'   + DisplayAlerts=False（防弹窗），读完立即关闭并还原。
Public Sub AttachLink_InsertCrossRef()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then
        MsgBox "当前底稿尚未保存。请先保存底稿（Ctrl+S）。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    ' V10：网页路径底稿无法创建本地附件文件夹
    If IsWebPathWorkbook(wb) Then
        MsgBox "底稿保存在网页路径（SharePoint/OneDrive 网页版），无法创建本地附件文件夹。" & vbCrLf & _
               "请先把底稿另存到本地磁盘或局域网共享盘后再插入。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要放置底稿引用的单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' R5：附件夹前置检查提到入口（三项校验之后、选目标底稿之前）——用户在选
    ' 文件之前就知道底稿副本将归档到哪里，可当场退出；取消即退出，不打开目标底稿
    Dim attachDir As String
    attachDir = EnsureAttachDirPrompted(fso, wb, "crossref", "本次插入的底稿副本将归档复制到该文件夹")
    If Len(attachDir) = 0 Then Exit Sub

    ' 第一步：选择底稿文件
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "第一步：选择要引用的底稿文件"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel 底稿", "*.xlsx;*.xls;*.xlsm"
        If Len(lastSourceDir) > 0 Then
            .InitialFileName = lastSourceDir & "\"
        Else
            .InitialFileName = wb.Path & "\"
        End If
        If .Show <> -1 Then Exit Sub
    End With
    Dim targetPath As String
    targetPath = fd.SelectedItems(1)
    lastSourceDir = fso.GetParentFolderName(targetPath)

    If StrComp(targetPath, wb.FullName, vbTextCompare) = 0 Then
        MsgBox "不能引用当前底稿自身。", vbInformation, "AttachLink"
        Exit Sub
    End If

    ' 第二步：安全打开目标底稿读取工作表列表（四重防护防卡死/防宏）
    ' oracle 二轮：EnableEvents=False 不等于禁宏——宏经 AutomationSecurity 才真正
    ' 关闭（msoAutomationSecurityForceDisable=3）；Fail 路径必须 stateCaptured 门控，
    ' 防止「捕获前出错」把未初始化的 prev* 默认 False 写回全局（关闭事件/提示/刷屏）
    Dim targetWb As Object, sheetList As String, i As Long
    Dim prevEvents As Boolean, prevAlerts As Boolean, prevScreen As Boolean
    Dim stateCaptured As Boolean, failDesc As String
    Dim prevAutoSec As Variant, autoSecSet As Boolean
    stateCaptured = False
    autoSecSet = False
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating
    stateCaptured = True
    On Error Resume Next
    Application.EnableEvents = False              ' 防事件连锁
    Application.DisplayAlerts = False             ' 防弹窗
    Application.ScreenUpdating = False
    Err.Clear
    prevAutoSec = Application.AutomationSecurity  ' 宿主不支持则不设（保持默认）
    If Err.Number = 0 Then
        Application.AutomationSecurity = 3        ' 真正禁宏：ForceDisable
        autoSecSet = (Err.Number = 0)
    End If
    Err.Clear                                     ' oracle 三轮：读取失败的旧错误不得泄入下方 Open 判定
    Set targetWb = Application.Workbooks.Open(targetPath, UpdateLinks:=0, ReadOnly:=True)
    If Err.Number <> 0 Or targetWb Is Nothing Then
        Err.Clear
        If Not targetWb Is Nothing Then targetWb.Close SaveChanges:=False   ' 半开状态兜底（Open 返回对象但带错）
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
        If autoSecSet Then Application.AutomationSecurity = prevAutoSec
        stateCaptured = False
        MsgBox "无法读取目标底稿（文件可能被锁定或格式不支持）。", vbExclamation, "AttachLink"
        On Error GoTo Fail
        Exit Sub
    End If
    On Error GoTo Fail

    Dim sheetCount As Long
    sheetCount = targetWb.Worksheets.Count
    For i = 1 To sheetCount
        sheetList = sheetList & "  " & i & ". " & targetWb.Worksheets(i).Name & vbCrLf
    Next i

    Dim pick As String, targetSheet As String
    If sheetCount = 1 Then
        targetSheet = targetWb.Worksheets(1).Name
    Else
        pick = InputBox("第二步：选择要引用的工作表（输入编号或名称）：" & vbCrLf & vbCrLf & _
                        sheetList, "AttachLink 底稿引用", "1")
        pick = Trim$(pick)
        If Len(pick) = 0 Then GoTo CleanupTarget
        If IsNumeric(pick) Then
            Dim idx As Long
            idx = CLng(pick)
            If idx < 1 Or idx > sheetCount Then
                MsgBox "编号超出范围（1-" & sheetCount & "）。", vbExclamation, "AttachLink"
                GoTo CleanupTarget
            End If
            targetSheet = targetWb.Worksheets(idx).Name
        Else
            Dim found As Boolean
            found = False
            For i = 1 To sheetCount
                If StrComp(targetWb.Worksheets(i).Name, pick, vbTextCompare) = 0 Then
                    targetSheet = targetWb.Worksheets(i).Name
                    found = True
                    Exit For
                End If
            Next i
            If Not found Then
                MsgBox "未找到工作表「" & pick & "」。", vbExclamation, "AttachLink"
                GoTo CleanupTarget
            End If
        End If
    End If

    ' 关闭目标底稿，还原应用状态
    targetWb.Close SaveChanges:=False
    Set targetWb = Nothing
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    If autoSecSet Then Application.AutomationSecurity = prevAutoSec
    stateCaptured = False

    ' 第三步：归档目标底稿并插入链接（附件夹已在入口前置检查创建/核对；
    ' Core 与「插入附件链接」同一身份规则：
    ' 同名同内容复用、同名异内容自动另存、插入即登记基线；自包含原则不变）
    Dim cellList As Collection
    Set cellList = FlattenSelection(Selection)
    If cellList.Count = 0 Then
        MsgBox "选区内没有可放置链接的单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If
    Dim summary As String
    summary = AttachLink_InsertCrossRefCore(targetPath, targetSheet, cellList(1))
    If Left$(summary, 3) = "失败：" Then
        MsgBox summary, vbCritical, "AttachLink 底稿引用"
    Else
        MsgBox summary, vbInformation, "AttachLink 底稿引用"
    End If
    Exit Sub
CleanupTarget:
    On Error Resume Next
    If Not targetWb Is Nothing Then targetWb.Close SaveChanges:=False
    If stateCaptured Then
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
    End If
    If autoSecSet Then Application.AutomationSecurity = prevAutoSec
    On Error GoTo Fail
    Exit Sub
Fail:
    On Error Resume Next
    failDesc = Err.Description                    ' 清理动作可能清掉 Err，先存
    If Not targetWb Is Nothing Then targetWb.Close SaveChanges:=False
    If stateCaptured Then
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
    End If
    If autoSecSet Then Application.AutomationSecurity = prevAutoSec
    MsgBox "插入底稿引用时出错：" & failDesc, vbCritical, "AttachLink"
End Sub

' 底稿引用核心（无对话框，供自动化直调）：把目标底稿归档入附件夹并插入带
' SubAddress 的链接。归档与「插入附件链接」共用同一身份判定（ArchiveFileEx）：
' 同名同内容静默复用、同名异内容自动另存 主名(n).ext（不覆盖原副本、不弹窗）、
' 插入即登记内容基线。返回汇总文本；「失败：」前缀=未插入。
' O-06/V2 重入防护包装：忙碌时直接失败返回；出错兜底——任何出口都复位标志，
' Inner 抛未处理错误如实转「失败：<过程名> - <原因>」
Public Function AttachLink_InsertCrossRefCore(ByVal targetPath As String, ByVal targetSheet As String, _
                                              ByVal targetCell As Object) As String
    If coreBusy Then
        AttachLink_InsertCrossRefCore = "失败：已有操作正在进行。"
        Exit Function
    End If
    coreBusy = True
    Err.Clear
    On Error Resume Next
    AttachLink_InsertCrossRefCore = InsertCrossRefCoreInner(targetPath, targetSheet, targetCell)
    If Err.Number <> 0 Then
        AttachLink_InsertCrossRefCore = "失败：AttachLink_InsertCrossRefCore - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    coreBusy = False
End Function

Private Function InsertCrossRefCoreInner(ByVal targetPath As String, ByVal targetSheet As String, _
                                         ByVal targetCell As Object) As String
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        InsertCrossRefCoreInner = "失败：没有活动工作簿。"
        Exit Function
    End If
    If Len(wb.Path) = 0 Then
        InsertCrossRefCoreInner = "失败：活动工作簿尚未保存。"
        Exit Function
    End If
    ' V10：网页路径底稿（SharePoint/OneDrive）下无法创建本地附件文件夹
    If IsWebPathWorkbook(wb) Then
        InsertCrossRefCoreInner = "失败：底稿保存在网页路径（SharePoint/OneDrive 网页版），无法创建本地附件文件夹。请先把底稿另存到本地磁盘或局域网共享盘。"
        Exit Function
    End If
    If targetCell Is Nothing Then
        InsertCrossRefCoreInner = "失败：未提供目标单元格。"
        Exit Function
    End If
    ' N-05：目标单元格必须属于当前底稿
    Dim ownerName As String
    On Error Resume Next
    ownerName = targetCell.Worksheet.Parent.FullName
    On Error GoTo Fail
    If Len(ownerName) = 0 Or StrComp(ownerName, wb.FullName, vbTextCompare) <> 0 Then
        InsertCrossRefCoreInner = "失败：目标单元格不属于当前底稿。"
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(targetPath) Then
        InsertCrossRefCoreInner = "失败：目标底稿不存在：" & targetPath
        Exit Function
    End If
    ' V7：底稿文件名含 # 时链接地址会被 Excel 截断，复制之前拒绝
    If InStr(fso.GetFileName(targetPath), "#") > 0 Then
        InsertCrossRefCoreInner = "失败：以下文件名含有 #，Excel 超链接无法正确保存，请重命名后再插入：" & fso.GetFileName(targetPath)
        Exit Function
    End If
    If StrComp(targetPath, wb.FullName, vbTextCompare) = 0 Then
        InsertCrossRefCoreInner = "失败：不能引用当前底稿自身。"
        Exit Function
    End If
    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME
    If Not EnsureFolder(fso, attachDir) Then
        InsertCrossRefCoreInner = "失败：无法创建附件文件夹：" & attachDir
        Exit Function
    End If
    ResetHashCache
    Dim baseStatus As Long, baseRO As Boolean
    Dim base As Object
    Set base = LoadBaselineEx(fso, attachDir, baseStatus, baseRO)
    ' O-04：基线损坏 = 失败关闭——在归档复制之前返回
    If baseStatus = 2 Then
        InsertCrossRefCoreInner = "失败：基线文件损坏或无法读取（" & lastBaselineError & _
                                  "），为避免绕过附件身份校验，本次未插入底稿引用。请先人工核对或重建基线。"
        Exit Function
    End If
    Dim baseDirty As Boolean, cat As Long, det As String
    Dim destFinal As String
    destFinal = ArchiveFileEx(fso, targetPath, attachDir, base, baseDirty, cat, det)
    If Len(destFinal) = 0 Then
        InsertCrossRefCoreInner = "失败：目标底稿未能归档（同名冲突确认中取消或无法写入）。"
        Exit Function
    End If

    ' 插入链接：Address=附件夹内相对路径，SubAddress=工作表!A1
    If targetCell.Hyperlinks.Count > 0 Then targetCell.Hyperlinks(1).Delete
    Dim relAddr As String, subAddr As String, dispName As String
    relAddr = ATTACH_FOLDER_NAME & "\" & fso.GetFileName(destFinal)
    dispName = fso.GetBaseName(targetPath) & "-" & targetSheet
    ' SubAddress 转义（oracle 复核）：统一单引号包裹（Excel 对简单表名同样合法），
    ' 表名内的单引号按 Excel 规则双写——空格/引号/括号/数字开头等全部安全
    subAddr = "'" & Replace(targetSheet, "'", "''") & "'!A1"
    targetCell.Hyperlinks.Add Anchor:=targetCell, Address:=relAddr, SubAddress:=subAddr, TextToDisplay:=dispName

    Dim persistNote As String
    If baseDirty Then
        If Not SaveBaseline(fso, attachDir, base) Then
            persistNote = vbCrLf & "■ 基线登记未持久化（" & lastBaselineError & "），" & BaselinePersistNote(lastBaselineError) & "。"
        End If
    End If

    Dim msg As String
    msg = "底稿引用已插入：" & vbCrLf & _
          "  显示名：" & dispName & vbCrLf & _
          "  底稿副本：" & relAddr & vbCrLf & _
          "  工作表：" & targetSheet
    If cat = CAT_RENAMED Then
        msg = msg & vbCrLf & "  注意：附件夹内已有同名但内容不同的副本，已另存为 " & det & "（原副本未被覆盖）。"
    ElseIf cat = CAT_REUSED Then
        msg = msg & vbCrLf & "  附件夹内已有内容相同的副本，已直接复用（未重复复制）。"
    ElseIf cat = CAT_ANOMALY Then
        ' O-05：副本与基线不一致未复用（体检会标红），det 为完整说明句
        msg = msg & vbCrLf & "  注意：" & det & "。"
    End If
    msg = msg & vbCrLf & "点击链接即打开副本并跳转到该工作表。" & persistNote
    InsertCrossRefCoreInner = msg
    Exit Function
Fail:
    InsertCrossRefCoreInner = "失败：插入底稿引用未完成 - " & Err.Description
End Function

'============================ 6. 归档迁移（批量）======================
' 场景：鼎信诺「归档」把处理后的底稿另存到新目录（如 桌面\客户名\财报底稿），
' 附件文件夹不在其处理范围。归档件中的链接仍指向账套原地址，且手工复制
' 附件夹后链接不会跟着变——归档件不自包含。
' 用法：在【归档目录】中打开任一底稿后运行，选择编制阶段的账套附件文件夹：
'   1) 把该附件文件夹整体【移动】到归档目录同级（账套不再保留副本，
'      避免双份存储；目标已有同名文件以账套原件覆盖）
'   2) 批量改写归档目录下所有 Excel 底稿的附件链接（绝对地址 -> 相对地址）
'   3) 汇总每个文件的处理结果；单个文件失败不影响其余文件
' 注意：移动后账套底稿的附件链接将失效（后续以归档目录为准；如需回溯
'       账套底稿，可用「附件断链修复」指向归档附件夹重连）
' 范围：仅处理归档目录第一层的 .xls/.xlsx/.xlsm/.xlsb（与鼎信诺归档输出一致）
Public Sub AttachLink_ArchiveMigrate()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then
        MsgBox "当前底稿尚未保存。请在【归档目录】中打开任一底稿后再运行本功能。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If InStr(1, wb.Path, "://", vbTextCompare) > 0 Then
        MsgBox "底稿位于网页路径（SharePoint/OneDrive 网页版），无法执行归档迁移。", vbExclamation, "AttachLink"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 源目录建议值：当前底稿中出现最多的绝对附件链接目录（即账套附件夹）
    Dim suggest As String
    suggest = SuggestAttachSource(wb, fso)

    Dim fd As Object
    Set fd = Application.FileDialog(4)
    With fd
        .Title = "选择编制阶段的「" & ATTACH_FOLDER_NAME & "」文件夹（源 = 账套中的附件文件夹）"
        If Len(suggest) > 0 Then
            .InitialFileName = suggest
        Else
            .InitialFileName = wb.Path & "\"
        End If
        If .Show <> -1 Then Exit Sub
    End With
    Dim srcDir As String
    srcDir = fd.SelectedItems(1)
    If Right$(srcDir, 1) = "\" Then srcDir = Left$(srcDir, Len(srcDir) - 1)
    ' 容错：选到账套目录（附件夹的上级）时自动进入其中的附件子文件夹
    If StrComp(fso.GetFileName(srcDir), ATTACH_FOLDER_NAME, vbTextCompare) <> 0 Then
        If fso.FolderExists(srcDir & "\" & ATTACH_FOLDER_NAME) Then
            srcDir = srcDir & "\" & ATTACH_FOLDER_NAME
        End If
    End If
    If Not fso.FolderExists(srcDir) Then
        MsgBox "源文件夹不存在：" & srcDir, vbCritical, "AttachLink"
        Exit Sub
    End If

    Dim dstDir As String
    dstDir = wb.Path & "\" & ATTACH_FOLDER_NAME

    If MsgBox("归档迁移计划（移动模式）：" & vbCrLf & vbCrLf & _
              "源（账套附件文件夹，整体移动）：" & vbCrLf & "  " & srcDir & vbCrLf & _
              "目标（归档目录同级「" & ATTACH_FOLDER_NAME & "」，同名以账套原件覆盖）：" & vbCrLf & "  " & dstDir & vbCrLf & vbCrLf & _
              "移动后：账套中的附件文件夹移除（不留副本，账套底稿链接将失效）；" & vbCrLf & _
              "归档底稿链接全部改写为相对地址，归档件自包含。" & vbCrLf & vbCrLf & _
              "「确定」开始　　「取消」退出", vbOKCancel + vbExclamation, "AttachLink") <> vbOK Then Exit Sub

    MsgBox AttachLink_ArchiveMigrateCore(srcDir), vbInformation, "AttachLink 归档迁移"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "归档迁移时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 归档迁移核心（独立公开函数：便于自动化测试与复用）。以当前活动工作簿
' 所在目录为归档目录。返回汇总文本；失败信息也包含在返回文本中。
' O-06 重入防护包装：忙碌时直接失败返回；包装器保证任何出口都复位标志
Public Function AttachLink_ArchiveMigrateCore(ByVal srcDir As String) As String
    If coreBusy Then
        AttachLink_ArchiveMigrateCore = "失败：已有操作正在进行。"
        Exit Function
    End If
    coreBusy = True
    Err.Clear
    On Error Resume Next
    AttachLink_ArchiveMigrateCore = ArchiveMigrateCoreInner(srcDir)
    If Err.Number <> 0 Then
        AttachLink_ArchiveMigrateCore = "失败：AttachLink_ArchiveMigrateCore - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    coreBusy = False
End Function

Private Function ArchiveMigrateCoreInner(ByVal srcDir As String) As String
    Dim wbHost As Object
    Set wbHost = ActiveWorkbook
    If wbHost Is Nothing Then
        ArchiveMigrateCoreInner = "失败：没有活动工作簿（请在归档目录中打开任一底稿后运行）。"
        Exit Function
    End If
    If Len(wbHost.Path) = 0 Then
        ArchiveMigrateCoreInner = "失败：活动工作簿尚未保存。"
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcDir) Then
        ArchiveMigrateCoreInner = "失败：源文件夹不存在：" & srcDir
        Exit Function
    End If

    Dim archiveRoot As String, dstDir As String
    archiveRoot = wbHost.Path
    dstDir = archiveRoot & "\" & ATTACH_FOLDER_NAME
    If Not EnsureFolder(fso, dstDir) Then
        ArchiveMigrateCoreInner = "失败：无法创建目标文件夹：" & dstDir
        Exit Function
    End If

    Dim prevEvents As Boolean, prevAlerts As Boolean, prevScreen As Boolean
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating

    Dim errs As Collection
    Set errs = New Collection
    Dim movedCnt As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' 1) 整体移动（目标已有同名文件以账套原件覆盖；单文件失败记入清单不中断；
    '    V5：移动失败的文件 rel 记入 moveFail——它们的链接不得改写，保持原样）
    Dim moveFail As Collection
    Set moveFail = New Collection
    movedCnt = MoveTreeInto(fso, srcDir, dstDir, errs, moveFail)

    ' 2) 批量改写归档目录第一层的 Excel 底稿
    Dim scanned As Long, touched As Long, rewritten As Long
    Dim f As Object, wb As Object, n As Long, opened As Boolean
    For Each f In fso.GetFolder(archiveRoot).Files
        Select Case LCase$(fso.GetExtensionName(f.Name))
            Case "xls", "xlsx", "xlsm", "xlsb"
            Case Else
                GoTo NextFile
        End Select
        If Left$(f.Name, 2) = "~$" Then GoTo NextFile
        scanned = scanned + 1
        Set wb = Nothing
        opened = False
        On Error Resume Next
        ' 已打开的（含活动工作簿）就地处理；按完整路径匹配，避免与
        ' 其他目录中的同名工作簿（如账套原件）混淆
        Set wb = WorkbookByPath(f.Path)
        If wb Is Nothing Then
            Set wb = Application.Workbooks.Open(f.Path)
            If wb Is Nothing Then
                errs.Add f.Name & "：打开失败 - " & Err.Description
                Err.Clear
                GoTo NextFile
            End If
            opened = True
        End If
        On Error GoTo FileFail
        n = RewriteAbsoluteAttachLinks(wb, fso)
        If n > 0 Then
            wb.Save
            touched = touched + 1
            rewritten = rewritten + n
        End If
        If opened Then wb.Close SaveChanges:=False
        GoTo NextFile
FileFail:
        errs.Add f.Name & "：" & Err.Description
        On Error Resume Next
        If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
        On Error GoTo 0
NextFile:
    Next f

    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen

    Dim s As String, i As Long
    s = "归档迁移完成：" & vbCrLf & _
        "· 附件已移动到 " & dstDir & "（成功 " & movedCnt & " 个文件）" & vbCrLf & _
        "· 扫描底稿 " & scanned & " 个，改写链接 " & rewritten & " 处（涉及 " & touched & " 个底稿）"
    If errs.Count > 0 Then
        s = s & vbCrLf & "· 失败/跳过 " & errs.Count & " 项："
        For i = 1 To errs.Count
            If i > 5 Then
                s = s & vbCrLf & "  ……（其余省略，共 " & errs.Count & " 项）"
                Exit For
            End If
            s = s & vbCrLf & "  " & errs(i)
        Next i
    End If
    ' V5：移动失败的文件如实单列——它们的链接保持原地址（未改写，不是断链漏报）
    If moveFail.Count > 0 Then
        s = s & vbCrLf & "· 以下文件移动失败，链接保持原样："
        For i = 1 To moveFail.Count
            If i > 10 Then
                s = s & vbCrLf & "  ……（其余省略，共 " & moveFail.Count & " 个）"
                Exit For
            End If
            s = s & vbCrLf & "  " & moveFail(i)
        Next i
    End If
    s = s & vbCrLf & vbCrLf & "建议对归档底稿运行「附件清单 / 体检」复核。"
    ArchiveMigrateCoreInner = s
    Exit Function
Fail:
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    ArchiveMigrateCoreInner = "归档迁移失败（未完成）：" & Err.Description & vbCrLf & _
        "已移动文件 " & movedCnt & " 个。"
End Function

'============================ 7. 附件版本更新 =========================
' 场景：附件存在差错需要有意替换。直接覆盖会触发体检「■ 内容变更」红字
' （无法与恶意替换区分）。本功能走受控更替流程：删除旧文件、写入新文件、
' 基线版本 +1 并记录更替时间——体检显示「一致 vN」绿字，且清单「更替记录」
' 列显式备注替换动作（审计留痕）。未经本流程的替换仍红字警示。
' 用法：选中带附件链接的单元格 → 右键「附件版本更新」→ 选择新文件 → 确认。
Public Sub AttachLink_UpdateVersion()
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then
        MsgBox "当前底稿尚未保存，无法定位附件文件夹。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    ' V10：网页路径底稿无法定位本地附件
    If IsWebPathWorkbook(wb) Then
        MsgBox "底稿保存在网页路径（SharePoint/OneDrive 网页版），无法执行附件版本更新。" & vbCrLf & _
               "请先把底稿另存到本地磁盘或局域网共享盘后再执行。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要更新版本的附件链接单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 取选区中第一个附件链接（超大选区经 GetScanRange 限域，防全选冻结）
    Dim c As Object, hl As Object, target As Object, scanR As Object
    Set target = Nothing
    Set scanR = GetScanRange()
    If Not scanR Is Nothing Then
        For Each c In scanR.Cells
            For Each hl In c.Hyperlinks
                If IsAttachmentLink(hl, fso, wb.Path) Then Set target = hl: Exit For
            Next hl
            If Not target Is Nothing Then Exit For
        Next c
    End If
    If target Is Nothing Then
        MsgBox "选区内未发现附件链接。请选中要更新版本的附件链接单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If

    Dim attachDir As String, fullPath As String, fName As String
    attachDir = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)
    ' O-03：规范化 + 附件夹边界判定（前缀匹配会把「附件\..\x」误判为夹内）
    fullPath = fso.GetAbsolutePathName(ResolveLinkPath(fso, wb.Path, target.Address))
    fName = fso.GetFileName(fullPath)
    If Not IsSubPathOf(fullPath, attachDir) Or StrComp(fullPath, attachDir, vbTextCompare) = 0 Then
        MsgBox "该链接指向本底稿附件文件夹之外（可能是归档前的指回链接），请先完成归档迁移或断链修复。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If Not fso.FileExists(fullPath) Then
        If fso.FolderExists(fullPath) Then
            MsgBox "「附件版本更新」暂不支持文件夹引用。" & vbCrLf & _
                   "如需更换文件夹内容：清除该链接后删除附件夹内的旧文件夹，重新「插入文件夹引用」，" & vbCrLf & _
                   "并删除附件夹内 " & BASELINE_FILE & " 后重新生成清单重建基线（注意：仅重插不重建基线，" & vbCrLf & _
                   "内容差异仍会显示「■ 内容变更」）。", vbInformation, "AttachLink"
        Else
            MsgBox "附件文件不存在（断链）：请先用「附件断链修复」恢复后再更新版本。", vbExclamation, "AttachLink"
        End If
        Exit Sub
    End If

    ' 基线必须已登记（版本号与历史来自基线）；损坏/只读基线如实拒绝，不得静默
    Dim base As Object, key As Variant, meta() As String
    Dim baseStatus As Long, baseRO As Boolean
    Set base = LoadBaselineEx(fso, attachDir, baseStatus, baseRO)
    If baseStatus = 2 Then
        MsgBox "基线文件损坏（" & lastBaselineError & "），无法执行版本更新。" & vbCrLf & _
               "请人工修复，或删除附件夹内 " & BASELINE_FILE & " 后重新生成清单重建基线。", vbCritical, "AttachLink"
        Exit Sub
    End If
    If baseRO Then
        MsgBox "基线含更新版本字段，本版不能执行版本更新（不会覆盖该基线）。请升级 AttachLink。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    key = LCase$(Mid$(fullPath, Len(attachDir) + 2))
    If Not base.Exists(key) Then
        MsgBox "该附件尚未登记内容基线。请先运行「附件清单 / 体检」登记基线，再进行版本更新。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    meta = Split(base(key), "|")
    ' meta: 0=哈希 1=版本 2=登记时间 3=更新时间

    ' 选择新文件
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "选择替换「" & fName & "」的新版本文件（v" & (CLng(meta(1)) + 1) & "）"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "所有文件", "*.*"
        .InitialFileName = attachDir & "\"
        If .Show <> -1 Then Exit Sub
    End With
    Dim newPath As String
    newPath = fd.SelectedItems(1)
    If StrComp(newPath, fullPath, vbTextCompare) = 0 Then
        MsgBox "选择的还是当前附件本身，无需更替。", vbInformation, "AttachLink"
        Exit Sub
    End If

    ' 内容相同无需更替
    Dim hNew As String
    hNew = Sha256File(newPath)
    If Len(hNew) = 0 Then
        MsgBox "新文件 SHA-256 校验不可用（读取失败），已取消。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    If StrComp(hNew, meta(0), vbTextCompare) = 0 Then
        MsgBox "新文件与当前版本（v" & meta(1) & "）内容完全相同，无需更替。", vbInformation, "AttachLink"
        Exit Sub
    End If

    If MsgBox("附件版本更替确认：" & vbCrLf & vbCrLf & _
              "附件：" & fName & vbCrLf & _
              "版本：v" & meta(1) & " → v" & (CLng(meta(1)) + 1) & vbCrLf & _
              "旧版本文件将被【替换】（更替动作将记入体检清单与基线，可留痕追溯）。" & vbCrLf & vbCrLf & _
              "「是」执行更替　　「否」取消", vbYesNo + vbExclamation, "AttachLink 附件版本更新") <> vbYes Then Exit Sub

    ' N-07：文件更替与基线更新收敛到 Core（可自动化直调、可失败注入测试）；
    ' Core 内部任何失败路径下原附件与基线均保持不变
    Dim coreMsg As String
    coreMsg = AttachLink_UpdateVersionCore(fullPath, newPath)
    If Left$(coreMsg, 2) = "失败" Then
        MsgBox coreMsg, vbCritical, "AttachLink"
    Else
        MsgBox coreMsg, vbInformation, "AttachLink"
    End If
    Exit Sub
Fail:
    MsgBox "附件版本更新时出错（文件可能被占用，请关闭相关程序后重试；基线未被修改）：" & Err.Description, vbCritical, "AttachLink"
End Sub

' N-07：附件版本更替核心（无对话框，供 UI 与自动化直调共用）。
' 更替走「newPath → 同目录临时文件 → ReplaceFileSafely 备份替换」（旧版 Delete+Copy
' 在中途失败时原文件已删：复制被占用/被截获即两失）；临时副本先与源哈希比对（N-06
' 同款防截获），全部成功后才写基线。返回结果消息；以「失败」开头表示未执行更替，
' 此时原附件字节与基线均未变动。
' V3：忙碌保护——体检进行中发起版本更新会把文件换成新内容、体检结束再用内存里
' 的旧字典回写基线（二次体检误报「■ 内容变更」）。与六 Core 同款包装（含 V2 出错
' 兜底）。已知限制：落盘前重读磁盘基线合并不做（跨进程基线无锁属已知限制，
' 同进程并发已由 coreBusy 排除；本批范围之外）。
Public Function AttachLink_UpdateVersionCore(ByVal fullPath As String, ByVal newPath As String) As String
    If coreBusy Then
        AttachLink_UpdateVersionCore = "失败：已有操作正在进行。"
        Exit Function
    End If
    coreBusy = True
    Err.Clear
    On Error Resume Next
    AttachLink_UpdateVersionCore = UpdateVersionCoreInner(fullPath, newPath)
    If Err.Number <> 0 Then
        AttachLink_UpdateVersionCore = "失败：AttachLink_UpdateVersionCore - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    coreBusy = False
End Function

Private Function UpdateVersionCoreInner(ByVal fullPath As String, ByVal newPath As String) As String
    On Error GoTo Fail
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        UpdateVersionCoreInner = "失败：没有活动工作簿。"
        Exit Function
    End If
    If Len(wb.Path) = 0 Then
        UpdateVersionCoreInner = "失败：底稿尚未保存。"
        Exit Function
    End If
    ' V10：网页路径底稿（SharePoint/OneDrive）下无法定位本地附件
    If IsWebPathWorkbook(wb) Then
        UpdateVersionCoreInner = "失败：底稿保存在网页路径（SharePoint/OneDrive 网页版），无法执行附件版本更新。请先把底稿另存到本地磁盘或局域网共享盘。"
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)
    fullPath = fso.GetAbsolutePathName(fullPath)
    ' O-03：入口统一规范化 + 附件夹边界判定
    If Not IsSubPathOf(fullPath, attachDir) Or StrComp(fullPath, attachDir, vbTextCompare) = 0 Then
        UpdateVersionCoreInner = "失败：附件路径不在本底稿附件文件夹内（可能是归档前的指回链接），请先完成归档迁移或断链修复。"
        Exit Function
    End If
    If Not fso.FileExists(fullPath) Then
        UpdateVersionCoreInner = "失败：附件文件不存在（断链），请先用「附件断链修复」恢复后再更新版本。"
        Exit Function
    End If
    If Not fso.FileExists(newPath) Then
        UpdateVersionCoreInner = "失败：新版本文件不存在：" & newPath
        Exit Function
    End If
    ' 基线必须已登记（版本号与历史来自基线）；损坏/只读基线如实拒绝，不得静默
    Dim base As Object, key As Variant, meta() As String
    Dim baseStatus As Long, baseRO As Boolean
    Set base = LoadBaselineEx(fso, attachDir, baseStatus, baseRO)
    If baseStatus = 2 Then
        UpdateVersionCoreInner = "失败：基线文件损坏（" & lastBaselineError & "），无法执行版本更新。" & vbCrLf & _
                   "请人工修复，或删除附件夹内 " & BASELINE_FILE & " 后重新生成清单重建基线。"
        Exit Function
    End If
    If baseRO Then
        UpdateVersionCoreInner = "失败：基线含更新版本字段，本版不能执行版本更新（不会覆盖该基线）。请升级 AttachLink。"
        Exit Function
    End If
    key = LCase$(Mid$(fullPath, Len(attachDir) + 2))
    If Not base.Exists(key) Then
        UpdateVersionCoreInner = "失败：该附件尚未登记内容基线。请先运行「附件清单 / 体检」登记基线，再进行版本更新。"
        Exit Function
    End If
    meta = Split(base(key), "|")
    ' meta: 0=哈希 1=版本 2=登记时间 3=更新时间
    If StrComp(fso.GetAbsolutePathName(newPath), fullPath, vbTextCompare) = 0 Then
        UpdateVersionCoreInner = "失败：选择的还是当前附件本身，无需更替。"
        Exit Function
    End If
    ' 内容相同无需更替
    Dim hNew As String
    hNew = Sha256File(newPath)
    If Len(hNew) = 0 Then
        UpdateVersionCoreInner = "失败：新文件 SHA-256 校验不可用（读取失败），已取消，原附件与基线均未改动。"
        Exit Function
    End If
    If StrComp(hNew, meta(0), vbTextCompare) = 0 Then
        UpdateVersionCoreInner = "新文件与当前版本（v" & meta(1) & "）内容完全相同，无需更替。"
        Exit Function
    End If

    ' 文件操作全部成功后才写基线（避免中途失败造成状态不一致）
    Dim tmpPath As String, hTmp As String
    tmpPath = attachDir & "\" & fso.GetFileName(newPath) & ".updtmp"
    On Error Resume Next
    If fso.FileExists(tmpPath) Then fso.DeleteFile tmpPath, True
    Err.Clear
    fso.CopyFile newPath, tmpPath, False
    If Err.Number <> 0 Then
        UpdateVersionCoreInner = "失败：新版本文件读取/复制失败（" & Err.Description & "），已取消，原附件与基线均未改动。"
        On Error Resume Next
        fso.DeleteFile tmpPath, True
        Err.Clear
        Exit Function
    End If
    ' N-06 同款：临时副本先与源哈希比对，防复制截断/被截获
    hTmp = Sha256File(tmpPath)
    If Len(hTmp) = 0 Or StrComp(hTmp, hNew, vbTextCompare) <> 0 Then
        UpdateVersionCoreInner = "失败：新版本文件读取校验不一致，已取消，原附件与基线均未改动。"
        On Error Resume Next
        fso.DeleteFile tmpPath, True
        Err.Clear
        Exit Function
    End If
    On Error GoTo Fail
    If Not ReplaceFileSafely(fso, tmpPath, fullPath) Then
        ' 清理临时副本，不留半截文件（与 SaveBaseline 的 tmp 契约一致）
        On Error Resume Next
        fso.DeleteFile tmpPath, True
        Err.Clear
        UpdateVersionCoreInner = "失败：替换旧版本失败（" & lastBaselineError & "），原附件与基线均未改动。"
        Exit Function
    End If

    Dim tCreate As String
    tCreate = meta(2)
    If Len(tCreate) = 0 Then tCreate = Format$(Now, "yyyy-mm-dd hh:nn")
    base(key) = hNew & "|" & (CLng(meta(1)) + 1) & "|" & tCreate & "|" & Format$(Now, "yyyy-mm-dd hh:nn")
    If Not SaveBaseline(fso, attachDir, base) Then
        ' 文件已替换但基线写失败：必须明说（体检将显示「■ 内容变更」），不得静默
        UpdateVersionCoreInner = "附件文件已替换为 v" & (CLng(meta(1)) + 1) & "，但基线写入失败：" & lastBaselineError & vbCrLf & _
                   "体检将显示「■ 内容变更」。请人工核对附件夹内 " & BASELINE_FILE & " 后重新运行「附件清单 / 体检」。"
        Exit Function
    End If

    UpdateVersionCoreInner = "附件版本更新完成：" & fso.GetFileName(fullPath) & " 已更替为 v" & (CLng(meta(1)) + 1) & "。" & vbCrLf & _
           "链接地址不变（同名替换）。建议重新运行「附件清单 / 体检」核对。"
    Exit Function
Fail:
    UpdateVersionCoreInner = "失败：附件版本更新时出错（文件可能被占用，请关闭相关程序后重试；基线未被修改）：" & Err.Description
End Function

'============================ 自检 ====================================
' 构建脚本用：确认模块可以编译
Public Function AttachLink_SelfTest() As Boolean
    AttachLink_SelfTest = (Len(ATTACH_FOLDER_NAME) > 0) And (Len(BASELINE_FILE) > 0) And (Len(ATTACHLINK_VERSION) > 0)
End Function

' 版本探针：供安装器/看门狗日志经 Application.Run 查询（与 ATTACHLINK_VERSION 常量同源）。
' 注意函数名不能叫 AttachLink_Version——VBA 标识符大小写不敏感，会与常量重名（实测踩坑）
Public Function AttachLink_GetVersion() As String
    AttachLink_GetVersion = ATTACHLINK_VERSION
End Function

'====================== 内容基线（SHA-256）辅助 ======================
' v1.6.0（R4）首选 advapi32 CAPI 原生 SHA-256：真实 WPS/Excel 实测约 300~350 MB/s
' （1.5.1 的 WPS 纯 VBA 路径实测仅 0.21 MB/s：插入 1MB 4.7s / 5MB 23.3s / 20MB 94s，
' 底稿引用 10MB 48.8s，体检 50MB 约 4 分钟，登记哈希成了主耗时）。CAPI 为系统原生
' 组件、无子进程（WPS 拦截 WScript.Shell/VBA.Shell 子进程创建，certutil 在 WPS 不可用），
' Excel 64 位与 WPS（#Else 分支）均编译运行通过，哈希与 Get-FileHash/标准向量一致。
' 回退路径：CAPI 返回空时按宿主分派——WPS 纯 VBA → certutil；Excel certutil → 纯 VBA
' （纯 VBA 含 50MB 护栏）。certutil 单文件约 150ms（Excel）进程开销，纯 VBA 仅兜底。
Private Function Sha256File(ByVal filePath As String) As String
    ' v1.6.0 分派：先 CAPI；空结果走原宿主回退链（宿主误判/环境异常时兜底，oracle 建议）
    hashComputeCount = hashComputeCount + 1   ' 真实计算计数（缓存命中不经过本函数）
    Dim h As String
    h = Sha256CAPI(filePath)
    If Len(h) > 0 Then
        Sha256File = h
        Exit Function
    End If
    If IsWpsHost() Then
        h = Sha256VBA(filePath)
        If Len(h) = 0 Then h = Sha256Certutil(filePath)
    Else
        h = Sha256Certutil(filePath)
        If Len(h) = 0 Then h = Sha256VBA(filePath)
    End If
    Sha256File = h
End Function

' ---- 单次运行哈希缓存（T14 去重基础）：同一 Core 运行内同一文件只真实计算一次 ----
Private Sub ResetHashCache()
    Set hashCache = CreateObject("Scripting.Dictionary")
End Sub

' 带缓存的哈希：命中直接返回；未命中真实计算并缓存（含失败结果空串，同次运行不重复尝试）
Private Function HashFileCached(ByVal filePath As String) As String
    If hashCache Is Nothing Then ResetHashCache
    Dim k As String
    k = LCase$(filePath)
    If hashCache.Exists(k) Then
        HashFileCached = hashCache(k)
    Else
        HashFileCached = Sha256File(filePath)
        hashCache(k) = HashFileCached
    End If
End Function

' 已知内容的文件直接写入缓存（复制落盘后立即登记，内容=源，无需重算）
Private Sub CacheHash(ByVal filePath As String, ByVal h As String)
    If Len(h) = 0 Then Exit Sub
    If hashCache Is Nothing Then ResetHashCache
    hashCache(LCase$(filePath)) = h
End Sub

' 诊断：本进程内 Sha256File 真实计算次数（缓存命中不计）
Public Function AttachLink_GetHashCount() As Long
    AttachLink_GetHashCount = hashComputeCount
End Function

' 诊断：清零真实计算计数（测试用）
Public Function AttachLink_ResetHashCount() As Long
    hashComputeCount = 0
    AttachLink_ResetHashCount = 0
End Function

' 测试/诊断入口（R4，不含对话框）：按指定引擎计算文件 SHA-256（小写十六进制或 ""）。
' engine：0 = 生产分派（等同 Sha256File，计入 hashComputeCount）；1 = 仅 CAPI；
' 2 = 仅 certutil；3 = 仅纯 VBA。1~3 直调对应引擎，不改计数、不走缓存。
Public Function AttachLink_HashEngine(ByVal filePath As String, ByVal engine As Long) As String
    Select Case engine
        Case 1
            AttachLink_HashEngine = Sha256CAPI(filePath)
        Case 2
            AttachLink_HashEngine = Sha256Certutil(filePath)
        Case 3
            AttachLink_HashEngine = Sha256VBA(filePath)
        Case Else
            AttachLink_HashEngine = Sha256File(filePath)
    End Select
End Function

' 测试/诊断入口（R4，不含对话框）：哈希引擎会话状态。
' 返回形如 "capi=1;win64=0"（capi：0 未自检 / 1 可用 / -1 不可用；win64：条件编译 0/1）
Public Function AttachLink_HashEngineState() As String
#If Win64 Then
    AttachLink_HashEngineState = "capi=" & capiState & ";win64=1"
#Else
    AttachLink_HashEngineState = "capi=" & capiState & ";win64=0"
#End If
End Function

' 测试钩子（仅测试使用，已登记编译探针）：设置本会话的 CAPI 引擎状态并返回设置前的值。
' newState：-1 = 视为「自检失败」（Sha256File 的生产分派改走回退链）；0 = 视为「尚未自检」（下次使用重新自检）；
' 其它值一律忽略——尤其不允许直接置成 1（可用），那会跳过自检、在 CAPI 真有问题的机器上产出错误哈希。
' 用于回归「CAPI 不可用时仍能经回退链得到正确哈希」；不读写任何文件，最坏后果是本会话哈希变慢。
Public Function AttachLink_TestCapiState(ByVal newState As Long) As Long
    AttachLink_TestCapiState = capiState
    ' 只接受 -1 与 0：不允许把状态直接置成 1——那会跳过自检，在 CAPI 真有问题的机器上产出错误哈希
    If newState = -1 Then
        capiState = -1
    ElseIf newState = 0 Then
        capiState = 0
    End If
End Function

' ---- T14 进度助手：状态栏 + DoEvents（WPS 不支持状态栏时静默跳过，不得出错）----
' 注意：本助手的 Resume Next 会吞掉 DoEvents 期间产生的 Err 18（用户 Esc），
' 必须检出并透传给调用方的错误分支，否则取消会失效
Private Sub Progress(ByVal msg As String)
    progressCount = progressCount + 1
    On Error Resume Next
    Application.StatusBar = msg
    DoEvents
    ' O-06 Err 卫生：DoEvents 期间产生的 Err 18（用户 Esc）必须检出并透传；
    ' 其它无关错误（宿主状态栏副作用等）清掉，不得冒泡成假失败
    If Err.Number = 18 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 18
    ElseIf Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

' 诊断：进度助手真实刷新次数（测试用）
Public Function AttachLink_GetProgressCount() As Long
    AttachLink_GetProgressCount = progressCount
End Function

' 诊断：清零进度计数（测试用）
Public Function AttachLink_ResetProgressCount() As Long
    progressCount = 0
    AttachLink_ResetProgressCount = 0
End Function

' 测试钩子（仅测试脚本使用）：n>0 时接下来 n 次「改指链接地址」不真正赋值、
' 按失败处理（TryRepoint 内逐次消耗）。本机 Excel 实测受保护工作表上
' Hyperlink.Address 赋值不抛错，T8/A4 的失败可见性链路借此端到端演练。
Public Sub AttachLink_TestRepointFail(ByVal n As Long)
    testRepointFail = n
End Sub

' 测试钩子（仅测试脚本使用）：n>0 时 CheckFolderBaseline 第 n 次进入时按用户取消处理
' （Err.Raise 18，演练 O-06 嵌套取消透传链路）。生产路径恒为 0。
Public Sub AttachLink_TestCancelInFolder(ByVal n As Long)
    testCancelInFolder = n
End Sub

' 测试钩子（仅测试脚本使用）：置位重入演练标志——InsertCore 运行期间再调
' InsertCrossRefCore，必须被重入防护拦截；拦截结果经 AttachLink_TestReentrySeen 读取。
Public Sub AttachLink_TestReentryProbe()
    testReentryProbe = True
End Sub

' 测试钩子（仅测试脚本使用）：一次性读取重入演练的拦截结果（读取后即清空，
' 避免跨测试组残留造成假断言；未演练时返回 ""）。
Public Function AttachLink_TestReentrySeen() As String
    AttachLink_TestReentrySeen = testReentrySeen
    testReentrySeen = ""
End Function

' 测试钩子（仅测试脚本使用）：n>0 时 InsertCoreInner 开头抛未处理错误，
' 验证包装器出错兜底（忙碌标志复位 + 「失败：<过程名> - <原因>」）。生产恒 0。
Public Sub AttachLink_TestThrowAtInsert(ByVal n As Long)
    testThrowAtInsert = n
End Sub

' 测试钩子（仅测试脚本使用）：n>0 时 ArchiveFileEx 接下来 n 次复制完成后向源文件
' 追加字节（源大小变化），验证 V11b 复制后校验以目标实际哈希登记。生产恒 0。
Public Sub AttachLink_TestGrowAfterCopy(ByVal n As Long)
    testGrowAfterCopy = n
End Sub

' 测试钩子（仅测试脚本使用）：V10 网页路径闸的可测入口——对给定路径文本返回
' 闸文案（"" = 可通过）。四个 Core 的闸都经 WebPathGateText/IsWebPathWorkbook。
Public Function AttachLink_TestWebPathGate(ByVal pathText As String) As String
    AttachLink_TestWebPathGate = WebPathGateText(pathText)
End Function

' 落盘失败附注（N-01：按失败原因区分措辞，不得张冠李戴）——
'   损坏        → 人工核对或删除后重新体检（体检并不会补登记）
'   未来格式只读 → 明确「本次未写入新登记」（不写「体检将补登记」，不提示检查写权限）
'   其它（权限/锁定/替换失败）→ 提示检查写权限或文件锁定
Private Function BaselinePersistNote(ByVal errText As String) As String
    If InStr(errText, "损坏") > 0 Then
        BaselinePersistNote = "基线文件损坏，请人工核对或删除后重新体检"
    ElseIf InStr(errText, "更新版本字段") > 0 Or InStr(errText, "只读") > 0 Then
        BaselinePersistNote = "基线为更新版本格式（本版只读），本次未写入新登记"
    Else
        BaselinePersistNote = "请检查附件夹写权限或文件锁定"
    End If
End Function

' 递归创建目录（EnsureFolder 的非递归版在父级缺失时会失败；子目录断链落回用）
Private Function EnsureFolderDeep(ByVal fso As Object, ByVal p As String) As Boolean
    On Error Resume Next
    If fso.FolderExists(p) Then
        EnsureFolderDeep = True
        Exit Function
    End If
    Dim parent As String
    parent = fso.GetParentFolderName(p)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then
            If Not EnsureFolderDeep(fso, parent) Then Exit Function
        End If
    End If
    fso.CreateFolder p
    EnsureFolderDeep = fso.FolderExists(p)
End Function

' 改指链接地址（A4 逐链接故障隔离：受保护/只读工作表上赋值可能抛错，不得中断整批；
' 失败时 errText 带原因）。测试钩子 testRepointFail>0 时不真正赋值、按失败处理。
Private Function TryRepoint(ByVal hl As Object, ByVal newAddr As String, ByRef errText As String) As Boolean
    If testRepointFail > 0 Then
        testRepointFail = testRepointFail - 1
        errText = "测试注入的改指失败"
        TryRepoint = False
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    hl.Address = newAddr
    If Err.Number <> 0 Then
        errText = Err.Description
        Err.Clear
        TryRepoint = False
    Else
        TryRepoint = True
    End If
End Function

' 确保基线已登记 key（仅缺失键登记 v1，不改既有版本/哈希）；哈希为空不登记
Private Sub EnsureBaselineKey(ByVal base As Object, ByVal key As String, ByVal h As String, ByRef dirty As Boolean)
    If Len(h) = 0 Then Exit Sub
    If Not base.Exists(key) Then
        base.Add key, h & "|1|" & Format$(Now, "yyyy-mm-dd hh:nn") & "|" & Format$(Now, "yyyy-mm-dd hh:nn")
        dirty = True
    End If
End Sub

' 宿主判定：WPS 安装路径含 kingsoft/wps（Application.Name 在 WPS 伪装为 Excel 不可靠）
Private Function IsWpsHost() As Boolean
    On Error Resume Next
    Dim p As String
    p = Application.Path
    IsWpsHost = (InStr(1, p, "kingsoft", vbTextCompare) > 0) Or (InStr(1, p, "wps", vbTextCompare) > 0)
End Function

' ---- CAPI 原生 SHA-256（v1.6.0 R4 主路径，advapi32，无子进程/无 COM/.NET 依赖）----
' 分块流式读取（VBA 原生 Open For Binary + Get，块 4MB），不设文件大小上限。
' 句柄存模块级 capiHProv/capiHHash（类型随条件编译），CapiFinish 无论成败都
' DestroyHash + ReleaseContext 并清零。
Private Function Sha256CAPI(ByVal filePath As String) As String
    On Error GoTo FailClose
    Sha256CAPI = ""
    If capiState < 0 Then Exit Function          ' 本会话自检失败过，不再重试（调用方回退）
    If capiState = 0 Then                        ' 首次使用自检：两个标准向量（空输入、"abc"）
        If Not CapiSelfTest() Then
            capiState = -1
            Exit Function
        End If
        capiState = 1
    End If
    Const CHUNK As Long = 4194304                ' 4MB
    Dim f As Integer, fsize As Long, remaining As Long, n As Long
    Dim buf() As Byte
    f = FreeFile
    Open filePath For Binary Access Read Shared As #f
    fsize = LOF(f)
    remaining = fsize
    If Not CapiBegin() Then GoTo FailClose
    Do While remaining > 0
        n = remaining
        If n > CHUNK Then n = CHUNK
        ReDim buf(0 To n - 1)                    ' 最后一块按剩余字节数
        Get #f, , buf()
        If Not CapiFeed(buf, n) Then GoTo FailClose
        remaining = remaining - n
    Loop
    Close #f
    Sha256CAPI = CapiFinish()                    ' 空串即失败（句柄已在 CapiFinish 释放）
    Exit Function
FailClose:
    On Error Resume Next
    Close #f
    CapiFinish
End Function

' CAPI 首次使用自检：两个标准 SHA-256 向量——空输入（0 字节文件走的就是这条路径）与 3 字节 "abc"，
' 任一不符即判失败（调用方把 capiState 置 -1，本会话回退 certutil / 纯 VBA）
Private Function CapiSelfTest() As Boolean
    CapiSelfTest = False
    If Not CapiBegin() Then Exit Function
    If CapiFinish() <> "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855" Then Exit Function
    Dim buf(0 To 2) As Byte
    buf(0) = 97: buf(1) = 98: buf(2) = 99
    If Not CapiBegin() Then Exit Function
    If Not CapiFeed(buf, 3) Then
        CapiFinish
        Exit Function
    End If
    CapiSelfTest = (CapiFinish() = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
End Function

' CAPI：获取上下文并创建 SHA-256 哈希对象（失败即释放已得句柄）
Private Function CapiBegin() As Boolean
    Const PROV_RSA_AES As Long = 24
    Const CALG_SHA_256 As Long = &H800C&
    Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
    CapiBegin = False
    If CryptAcquireContext(capiHProv, 0, 0, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then Exit Function
    If CryptCreateHash(capiHProv, CALG_SHA_256, 0, 0, capiHHash) = 0 Then
        CryptReleaseContext capiHProv, 0
        capiHProv = 0
        capiHHash = 0
        Exit Function
    End If
    CapiBegin = True
End Function

' CAPI：喂入一段数据（n=0 跳过；数据指针取数组首元素地址，与探针一致）
Private Function CapiFeed(buf() As Byte, ByVal n As Long) As Boolean
    CapiFeed = False
    If n = 0 Then CapiFeed = True: Exit Function
    If capiHHash = 0 Then Exit Function
    CapiFeed = (CryptHashData(capiHHash, VarPtr(buf(LBound(buf))), n, 0) <> 0)
End Function

' CAPI：取哈希值（小写十六进制）并无论成败都 DestroyHash + ReleaseContext、句柄清零
Private Function CapiFinish() As String
    CapiFinish = ""
    If capiHHash <> 0 Then
        Const HP_HASHVAL As Long = &H2&
        Dim outb(0 To 31) As Byte, outLen As Long, i As Long, s As String
        outLen = 32
        If CryptGetHashParam(capiHHash, HP_HASHVAL, outb(0), outLen, 0) <> 0 Then
            For i = 0 To outLen - 1
                s = s & Right$("0" & Hex$(outb(i)), 2)
            Next i
            CapiFinish = LCase$(s)
        End If
        CryptDestroyHash capiHHash
        capiHHash = 0
    End If
    If capiHProv <> 0 Then
        CryptReleaseContext capiHProv, 0
        capiHProv = 0
    End If
End Function

Private Function Sha256Certutil(ByVal filePath As String) As String
    On Error GoTo Fail
    Sha256Certutil = ""
    Dim fso As Object, tmp As String, sh As Object, cmd As String, rc As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 空文件特判（R4）：certutil 对 0 字节文件报 0x800703ee（ERROR_FILE_INVALID），
    ' 而空输入的 SHA-256 是固定标准向量，直接返回（与 Sha256VBA 空文件特判同理）
    If fso.GetFile(filePath).Size = 0 Then
        Sha256Certutil = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
        Exit Function
    End If
    ' 唯一临时文件（oracle：固定名在多 Excel 并发体检时会互相覆盖/读到旧结果，串哈希属证据正确性问题）
    tmp = Environ$("TEMP") & "\" & fso.GetTempName
    ' chcp 65001：中文路径下 certutil 标题行按 UTF-8 输出避免管道乱码；哈希行为纯 ASCII
    cmd = "cmd /c chcp 65001>nul & certutil -hashfile """ & filePath & """ SHA256>""" & tmp & """ 2>nul"
    Set sh = CreateObject("WScript.Shell")
    rc = sh.Run(cmd, 0, True)                             ' 隐藏窗口 + 等待完成；返回退出码
    If rc <> 0 Then GoTo Cleanup                          ' 命令失败不得解析历史文件
    If Not fso.FileExists(tmp) Then GoTo Cleanup
    ' 解析：找首个 64 位十六进制行
    Dim stm As Object, text As String, lines() As String, i As Long, line As String
    Dim j As Long, hexOk As Boolean, ch As String
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile tmp
        .Position = 0
        text = .ReadText
        .Close
    End With
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        line = UCase$(Replace(Replace(lines(i), vbCr, ""), " ", ""))
        If Len(line) = 64 Then
            hexOk = True
            For j = 1 To 64
                ch = Mid$(line, j, 1)
                If InStr("0123456789ABCDEF", ch) = 0 Then hexOk = False: Exit For
            Next j
    If hexOk Then Sha256Certutil = LCase$(line): Exit For
        End If
    Next i
Cleanup:
    On Error Resume Next
    fso.DeleteFile tmp, True                              ' 成功/失败均清理
    On Error GoTo 0
    Exit Function
Fail:
    Sha256Certutil = ""
    Resume Cleanup                                     ' 异常路径同样清理临时文件（oracle 复验）
End Function

' ---- 纯 VBA SHA-256（WPS 路径；无任何外部依赖/Declare）----
' 性能预期约 0.3~1 MB/s（VBA 解释循环）：小附件秒级，数 MB 附件十秒级——
' WPS 场景登记为一次性成本，可接受；Excel 不走此路径。
Private Function Sha256VBA(ByVal filePath As String) As String
    On Error GoTo Fail
    Sha256VBA = ""
    ' 读入全文件
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 空文件特判：直接返回标准 SHA-256（e3b0c442...），避免 ADODB 空 Read 的 UBound 失败
    If fso.GetFile(filePath).Size = 0 Then
        Sha256VBA = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
        Exit Function
    End If
    ' 大文件护栏（oracle 建议）：全文件读入 + 约 0.3~1MB/s，超 50MB 直接降级「校验不可用」
    If fso.GetFile(filePath).Size > 52428800 Then Exit Function
    Dim stm As Object, bytes As Variant
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.LoadFromFile filePath
    stm.Position = 0
    bytes = stm.Read
    stm.Close
    Dim n As Long
    n = UBound(bytes) - LBound(bytes) + 1

    ' 消息填充：msg = data + &H80 + 零 + 8 字节大端位长（512 对齐）
    Dim totalLen As Double
    totalLen = CDbl(n) * 8#
    Dim padded As Long
    padded = ((n + 8) \ 64 + 1) * 64
    Dim msg() As Byte
    ReDim msg(0 To padded - 1)
    Dim i As Long
    For i = 0 To n - 1
        msg(i) = bytes(i)
    Next i
    msg(n) = 128
    For i = 0 To 7
        msg(padded - 1 - i) = CByte(Int(totalLen / (256# ^ i)) - Int(Int(totalLen / (256# ^ i)) / 256#) * 256#)
    Next i

    ' 初始哈希
    Dim h0 As Long, h1 As Long, h2 As Long, h3 As Long, h4 As Long, h5 As Long, h6 As Long, h7 As Long
    h0 = &H6A09E667: h1 = &HBB67AE85: h2 = &H3C6EF372: h3 = &HA54FF53A
    h4 = &H510E527F: h5 = &H9B05688C: h6 = &H1F83D9AB: h7 = &H5BE0CD19

    ' 常量 K
    Dim K(0 To 63) As Long
    K(0) = &H428A2F98: K(1) = &H71374491: K(2) = &HB5C0FBCF: K(3) = &HE9B5DBA5
    K(4) = &H3956C25B: K(5) = &H59F111F1: K(6) = &H923F82A4: K(7) = &HAB1C5ED5
    K(8) = &HD807AA98: K(9) = &H12835B01: K(10) = &H243185BE: K(11) = &H550C7DC3
    K(12) = &H72BE5D74: K(13) = &H80DEB1FE: K(14) = &H9BDC06A7: K(15) = &HC19BF174
    K(16) = &HE49B69C1: K(17) = &HEFBE4786: K(18) = &HFC19DC6: K(19) = &H240CA1CC
    K(20) = &H2DE92C6F: K(21) = &H4A7484AA: K(22) = &H5CB0A9DC: K(23) = &H76F988DA
    K(24) = &H983E5152: K(25) = &HA831C66D: K(26) = &HB00327C8: K(27) = &HBF597FC7
    K(28) = &HC6E00BF3: K(29) = &HD5A79147: K(30) = &H6CA6351: K(31) = &H14292967
    K(32) = &H27B70A85: K(33) = &H2E1B2138: K(34) = &H4D2C6DFC: K(35) = &H53380D13
    K(36) = &H650A7354: K(37) = &H766A0ABB: K(38) = &H81C2C92E: K(39) = &H92722C85
    K(40) = &HA2BFE8A1: K(41) = &HA81A664B: K(42) = &HC24B8B70: K(43) = &HC76C51A3
    K(44) = &HD192E819: K(45) = &HD6990624: K(46) = &HF40E3585: K(47) = &H106AA070
    K(48) = &H19A4C116: K(49) = &H1E376C08: K(50) = &H2748774C: K(51) = &H34B0BCB5
    K(52) = &H391C0CB3: K(53) = &H4ED8AA4A: K(54) = &H5B9CCA4F: K(55) = &H682E6FF3
    K(56) = &H748F82EE: K(57) = &H78A5636F: K(58) = &H84C87814: K(59) = &H8CC70208
    K(60) = &H90BEFFFA: K(61) = &HA4506CEB: K(62) = &HBEF9A3F7: K(63) = &HC67178F2

    ' 分块处理
    Dim blk As Long, w(0 To 63) As Long
    Dim a As Long, b As Long, cc As Long, d As Long, e As Long, f As Long, g As Long, hh As Long
    Dim t1 As Long, t2 As Long, j As Long, off As Long
    For blk = 0 To padded - 1 Step 64
        off = blk
        For j = 0 To 15
            w(j) = DblToLong(CDbl(msg(off)) * 16777216# + CDbl(msg(off + 1)) * 65536# + CDbl(msg(off + 2)) * 256# + CDbl(msg(off + 3)))
            off = off + 4
        Next j
        For j = 16 To 63
            w(j) = Add32(Add32(Sig1(w(j - 2)), w(j - 7)), Add32(Sig0(w(j - 15)), w(j - 16)))
        Next j
        a = h0: b = h1: cc = h2: d = h3: e = h4: f = h5: g = h6: hh = h7
        For j = 0 To 63
            t1 = Add32(Add32(Add32(Add32(Ep1(e), Ch(e, f, g)), hh), K(j)), w(j))
            t2 = Add32(Ep0(a), Maj(a, b, cc))
            hh = g: g = f: f = e
            e = Add32(d, t1)
            d = cc: cc = b: b = a
            a = Add32(t1, t2)
        Next j
        h0 = Add32(h0, a): h1 = Add32(h1, b): h2 = Add32(h2, cc): h3 = Add32(h3, d)
        h4 = Add32(h4, e): h5 = Add32(h5, f): h6 = Add32(h6, g): h7 = Add32(h7, hh)
    Next blk

    Sha256VBA = LCase$(ToHex8(h0) & ToHex8(h1) & ToHex8(h2) & ToHex8(h3) & _
                       ToHex8(h4) & ToHex8(h5) & ToHex8(h6) & ToHex8(h7))
    Exit Function
Fail:
    Sha256VBA = ""
End Function

' ---- SHA-256 位运算辅助（Long 按补码，模 2^32；移位经 Double 中转避免负数整除方向错）----
Private Function Add32(ByVal x As Long, ByVal y As Long) As Long
    Dim dx As Double, dy As Double, s As Double
    dx = x: If dx < 0 Then dx = dx + 4294967296#
    dy = y: If dy < 0 Then dy = dy + 4294967296#
    s = dx + dy
    Do While s >= 4294967296#
        s = s - 4294967296#
    Loop
    Add32 = DblToLong(s)
End Function

Private Function DblToLong(ByVal d As Double) As Long
    If d >= 2147483648# Then d = d - 4294967296#
    DblToLong = CLng(d)
End Function

Private Function Ch(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Ch = (x And y) Or ((Not x) And z)
End Function

Private Function Maj(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Maj = (x And y) Or (x And z) Or (y And z)
End Function

' 循环右移 n 位（无符号语义）
Private Function Rotr(ByVal x As Long, ByVal n As Long) As Long
    Dim d As Double, hi As Double, lo As Double, v As Double
    d = x: If d < 0 Then d = d + 4294967296#
    hi = Int(d / (2# ^ n))                       ' 高 32-n 位下移
    lo = d - hi * (2# ^ n)                       ' 低 n 位
    v = hi + lo * (2# ^ (32 - n))
    Do While v >= 4294967296#
        v = v - 4294967296#
    Loop
    Rotr = DblToLong(v)
End Function

' 逻辑右移 n 位（无符号语义）
Private Function Shr(ByVal x As Long, ByVal n As Long) As Long
    Dim d As Double
    d = x: If d < 0 Then d = d + 4294967296#
    Shr = DblToLong(Int(d / (2# ^ n)))
End Function

Private Function Ep0(ByVal x As Long) As Long
    Ep0 = Rotr(x, 2) Xor Rotr(x, 13) Xor Rotr(x, 22)
End Function

Private Function Ep1(ByVal x As Long) As Long
    Ep1 = Rotr(x, 6) Xor Rotr(x, 11) Xor Rotr(x, 25)
End Function

Private Function Sig0(ByVal x As Long) As Long
    Sig0 = Rotr(x, 7) Xor Rotr(x, 18) Xor Shr(x, 3)
End Function

Private Function Sig1(ByVal x As Long) As Long
    Sig1 = Rotr(x, 17) Xor Rotr(x, 19) Xor Shr(x, 10)
End Function

Private Function ToHex8(ByVal x As Long) As String
    ' WPS 兼容：Hex$ 对 >=2^31 的数值会经 Long 转换而溢出（实测 E6）——
    ' 按字节分段格式化，每段 0~255 永不溢出
    Dim d As Double, s As String, i As Long, b As Double
    d = x: If d < 0 Then d = d + 4294967296#
    s = ""
    For i = 3 To 0 Step -1
        b = Int(d / (256# ^ i))
        d = d - b * (256# ^ i)
        s = s & Right$("0" & Hex$(CLng(b)), 2)
    Next i
    ToHex8 = s
End Function
' 统一基线解析内核（LoadBaseline / LoadBaselineStrict / SaveBaseline 共用）：
' status 返回 0=正常（可能为空字典）/ 1=基线文件不存在 / 2=读取或解析失败。
' 逐行规则：
'   2 字段（v1）       → 视作 哈希|1||；
'   5 字段（v2）       → 正常；
'   >5 字段（将来版本）→ 前 5 字段按 v2 语义读取（键/哈希/版本非空即可比对），
'                        但 isReadOnly=True——本版不得覆盖写该文件；
'   其它字段数、空键/空哈希/空版本、归一化后重复键 → status=2（errDetail 注明行号）。
' 任何异常（含文件被独占锁定、流读取失败）都如实上报 status=2，绝不把
' 「读取失败」伪装成「空基线」回写覆盖（oracle 复核 #3）。
Private Function ParseBaselineFile(ByVal fso As Object, ByVal filePath As String, _
                                   ByRef status As Long, ByRef isReadOnly As Boolean, _
                                   ByRef errDetail As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    status = 0: isReadOnly = False: errDetail = ""
    If Not fso.FileExists(filePath) Then
        status = 1
        Set ParseBaselineFile = d
        Exit Function
    End If
    On Error GoTo BadRead
    Dim stm As Object, text As String, lines() As String, i As Long
    Dim line As String, parts() As String, verS As String, tC As String, tU As String
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2                                   ' adTypeText
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        .Position = 0
        text = .ReadText
        .Close
    End With
    lines = Split(text, vbLf)
    Dim keyN As String
    For i = 0 To UBound(lines)
        line = Replace(lines(i), vbCr, "")
        If Len(line) > 0 Then
            parts = Split(line, "|")
            If UBound(parts) = 1 Then
                verS = "1": tC = "": tU = ""
            ElseIf UBound(parts) = 4 Then
                verS = parts(2): tC = parts(3): tU = parts(4)
            ElseIf UBound(parts) > 4 Then
                ' 将来版本写入：按 v2 前缀读取比对，标记只读（本版不得覆盖写）
                verS = parts(2): tC = parts(3): tU = parts(4)
                isReadOnly = True
            Else
                errDetail = "第 " & (i + 1) & " 行无法解析（字段数 " & (UBound(parts) + 1) & "）"
                GoTo BadFormat
            End If
            If Len(parts(0)) = 0 Or Len(parts(1)) = 0 Or Len(verS) = 0 Then
                errDetail = "第 " & (i + 1) & " 行无法解析（空键/空哈希/空版本）"
                GoTo BadFormat
            End If
            keyN = LCase$(parts(0))
            If d.Exists(keyN) Then
                errDetail = "第 " & (i + 1) & " 行无法解析（重复键 " & parts(0) & "）"
                GoTo BadFormat
            End If
            d(keyN) = parts(1) & "|" & verS & "|" & tC & "|" & tU
        End If
    Next i
    Set ParseBaselineFile = d
    Exit Function
BadRead:
    errDetail = "读取失败：" & Err.Description
BadFormat:
    status = 2
    Set d = CreateObject("Scripting.Dictionary")
    Set ParseBaselineFile = d
End Function

' 基线加载（带状态）：status / isReadOnly 语义见 ParseBaselineFile；
' status=2 时 lastBaselineError 带原因，供调用方如实提示
Private Function LoadBaselineEx(ByVal fso As Object, ByVal attachDir As String, _
                                ByRef status As Long, ByRef isReadOnly As Boolean) As Object
    Dim det As String
    Set LoadBaselineEx = ParseBaselineFile(fso, attachDir & "\" & BASELINE_FILE, status, isReadOnly, det)
    If status = 2 Then lastBaselineError = det
End Function

' 兼容包装：调用方不区分状态时使用（需要状态请用 LoadBaselineEx）
Private Function LoadBaseline(ByVal fso As Object, ByVal attachDir As String) As Object
    Dim st As Long, ro As Boolean
    Set LoadBaseline = LoadBaselineEx(fso, attachDir, st, ro)
End Function

' 严格版基线加载（删除等破坏性路径专用，oracle 复核 #3）：status 0/1/2 语义同上
Private Function LoadBaselineStrict(ByVal fso As Object, ByVal attachDir As String, ByRef status As Long) As Object
    Dim ro As Boolean
    Set LoadBaselineStrict = LoadBaselineEx(fso, attachDir, status, ro)
End Function

' 基线安全落盘（UTF-8）。返回 True=成功；False=失败且原因写入 lastBaselineError
' （调用方不得报告登记完成——oracle #3：静默吞错会让未持久化登记被当作已完成）。
' 安全契约：
'   (a) 已有基线先严格读取——损坏（status=2）或含更新版本字段（只读）时绝不覆盖；
'   (b) 先写 临时文件（…csv.tmp），不直接动正式文件；
'   (c) 写后严格回读临时文件，条目数须等于字典条目数，否则返回 False 且原文件保持完整；
'   (d) 任何失败路径都清理临时文件，不留半截正式文件。
Private Function SaveBaseline(ByVal fso As Object, ByVal attachDir As String, ByVal d As Object) As Boolean
    lastBaselineError = ""
    On Error GoTo Fail
    If Not EnsureFolder(fso, attachDir) Then
        lastBaselineError = "无法创建附件文件夹"
        Exit Function
    End If
    Dim p As String, tmp As String
    p = attachDir & "\" & BASELINE_FILE
    If fso.FileExists(p) Then
        Dim st As Long, ro As Boolean, det As String, chk As Object
        Set chk = ParseBaselineFile(fso, p, st, ro, det)
        If st = 2 Then
            lastBaselineError = "基线文件损坏（" & det & "），未覆盖"
            Exit Function
        End If
        If ro Then
            lastBaselineError = "基线含更新版本字段，本版不覆盖，请升级 AttachLink"
            Exit Function
        End If
    End If
    tmp = p & ".tmp"
    Dim stm As Object, k As Variant
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
    End With
    For Each k In d.Keys
        stm.WriteText k & "|" & d(k) & vbCrLf
    Next k
    stm.SaveToFile tmp, 2                                     ' adSaveCreateOverWrite
    stm.Close
    ' 写后严格回读校验（在替换正式文件之前）
    Dim st2 As Long, ro2 As Boolean, det2 As String, chk2 As Object
    Set chk2 = ParseBaselineFile(fso, tmp, st2, ro2, det2)
    If st2 <> 0 Or chk2.Count <> d.Count Then
        lastBaselineError = "临时文件回读校验失败（" & det2 & "），原基线未改动"
        On Error Resume Next
        fso.DeleteFile tmp, True
        Exit Function
    End If
    ' 安全替换（先备份再替换；失败自动回滚，正式路径绝不两失）
    If Not ReplaceFileSafely(fso, tmp, p) Then
        On Error Resume Next
        fso.DeleteFile tmp, True
        Exit Function                                  ' lastBaselineError 已由 ReplaceFileSafely 写明
    End If
    SaveBaseline = True
    Exit Function
Fail:
    lastBaselineError = Err.Description
    On Error Resume Next
    If Len(tmp) > 0 Then fso.DeleteFile tmp, True
End Function

' 安全替换（独立小函数，便于失败注入测试）：把已校验的 tmp 替换到 dst。
' 顺序：清残留 .bak → dst 存在则 MoveFile dst→dst.bak → MoveFile tmp→dst；
' 后一步失败把 .bak 改回 dst 并返回 False；成功后删除 .bak。
' 契约：任何失败路径下，dst 处要么是原文件、要么是完整新文件，绝不两者同时消失。
Private Function ReplaceFileSafely(ByVal fso As Object, ByVal tmp As String, ByVal dst As String) As Boolean
    Dim bak As String, hadOrig As Boolean, mvErr As String
    bak = dst & ".bak"
    hadOrig = fso.FileExists(dst)
    On Error Resume Next
    If fso.FileExists(bak) Then fso.DeleteFile bak, True
    If hadOrig Then fso.MoveFile dst, bak
    If Err.Number <> 0 Then
        lastBaselineError = "备份原基线失败（" & Err.Description & "），正式文件未改动"
        Err.Clear
        ReplaceFileSafely = False
        Exit Function
    End If
    fso.MoveFile tmp, dst
    If Err.Number <> 0 Then
        mvErr = Err.Description
        Err.Clear
        If hadOrig Then fso.MoveFile bak, dst            ' 回滚：备份改回正式文件
        If Err.Number <> 0 Then
            lastBaselineError = "替换正式文件失败（" & mvErr & "），且回滚失败（" & Err.Description & _
                                "）——原基线完整保留在 " & bak & "，请人工改回"
        Else
            lastBaselineError = "替换正式文件失败（" & mvErr & "），已回滚为原基线"
        End If
        Err.Clear
        ReplaceFileSafely = False
        Exit Function
    End If
    If hadOrig Then fso.DeleteFile bak, True
    If Err.Number <> 0 Then Err.Clear                    ' .bak 残留不影响替换成功事实
    ReplaceFileSafely = True
End Function

'============================ 内部辅助 ================================
' 递归体检文件夹内全部文件的基线：缺失键登记 v1（dirty 置位）；哈希不一致计入 badCnt；
' 哈希失败或目录遍历失败计入 failCnt（调用方显示「校验不可用」，不得显示「一致」）
Private Sub CheckFolderBaseline(ByVal fso As Object, ByVal base As Object, _
                                ByVal folderPath As String, ByVal attachDir As String, _
                                ByRef regCnt As Long, ByRef fileCnt As Long, ByRef badCnt As Long, _
                                ByRef failCnt As Long, ByRef dirty As Boolean)
    Dim f As Object, sf As Object, key As String, h As String, cur As String
    On Error Resume Next
    ' O-06 测试钩子（仅测试使用）：第 N 次进入时模拟用户在进度刷新时按 Esc——
    ' 以 On Error GoTo 0 原样抛出，专测「嵌套子文件夹里取消被父级 Resume Next 吞掉」的透传链路
    If testCancelInFolder > 0 Then
        testCancelInFolder = testCancelInFolder - 1
        If testCancelInFolder = 0 Then
            On Error GoTo 0
            Err.Raise 18
        End If
    End If
    For Each f In fso.GetFolder(folderPath).Files
        fileCnt = fileCnt + 1
        ' T14：文件夹体检进度（每 8 个文件刷新一次状态栏）
        If fileCnt Mod 8 = 0 Then Progress "文件夹 " & fso.GetFileName(folderPath) & "：已校验 " & fileCnt & " 个文件"
        key = LCase$(Mid$(f.Path, Len(attachDir) + 2))
        h = HashFileCached(f.Path)
        If Len(h) = 0 Then
            failCnt = failCnt + 1
        ElseIf Not base.Exists(key) Then
            base.Add key, h & "|1|" & Format$(Now, "yyyy-mm-dd hh:nn") & "|" & Format$(Now, "yyyy-mm-dd hh:nn")
            regCnt = regCnt + 1
            dirty = True
        Else
            cur = Split(base(key), "|")(0)
            If StrComp(cur, h, vbTextCompare) <> 0 Then badCnt = badCnt + 1
        End If
    Next f
    ' 用户在进度刷新时按 Esc：透传给体检主循环的取消分支，不得吞掉也不得误计为失败
    If Err.Number = 18 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 18
    End If
    If Err.Number <> 0 Then failCnt = failCnt + 1: Err.Clear
    For Each sf In fso.GetFolder(folderPath).SubFolders
        CheckFolderBaseline fso, base, sf.Path, attachDir, regCnt, fileCnt, badCnt, failCnt, dirty
        ' O-06：子目录递归返回后【首先】检查取消——不得落入下面的通用错误分支被吞掉
        If Err.Number = 18 Then
            Err.Clear
            On Error GoTo 0
            Err.Raise 18
        End If
        If Err.Number <> 0 Then failCnt = failCnt + 1: Err.Clear
    Next sf
    ' 子目录循环结束后同样先查取消再处理通用错误
    If Err.Number = 18 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 18
    End If
    If Err.Number <> 0 Then failCnt = failCnt + 1: Err.Clear
End Sub

' R5：附件夹前置提示的会话级去重记录（不含对话框，供回归直调）。
' 语义：NOTICE_ONCE_PER_FOLDER=False → 每次都需要提示（Needed 恒 True）；
' True → 每个「插入入口 × 附件夹」每次会话仅提示一次，三个入口相互独立。
' 键 = 入口标识 & "|" & 小写附件夹路径（路径大小写不敏感）。
Public Function AttachLink_FolderNoticeNeeded(ByVal attachDir As String, ByVal entryId As String) As Boolean
    If Not NOTICE_ONCE_PER_FOLDER Then
        AttachLink_FolderNoticeNeeded = True
        Exit Function
    End If
    EnsureNoticeDict
    AttachLink_FolderNoticeNeeded = Not notifiedAttachDirs.Exists(entryId & "|" & LCase$(attachDir))
End Function

' R5：登记「某入口对某附件夹已提示过一次」（已存在→提示、不存在→自动创建
' 两个分支共用；只标记传入的那个入口）。
Public Sub AttachLink_FolderNoticeMark(ByVal attachDir As String, ByVal entryId As String)
    EnsureNoticeDict
    Dim nKey As String
    nKey = entryId & "|" & LCase$(attachDir)
    If Not notifiedAttachDirs.Exists(nKey) Then notifiedAttachDirs.Add nKey, True
End Sub

' R5：清空提示记录（回归用；会话内回到「全部未提示」状态）。
Public Sub AttachLink_FolderNoticeReset()
    If Not notifiedAttachDirs Is Nothing Then notifiedAttachDirs.RemoveAll
End Sub

Private Sub EnsureNoticeDict()
    If notifiedAttachDirs Is Nothing Then Set notifiedAttachDirs = CreateObject("Scripting.Dictionary")
End Sub

' 前置检查（三个插入入口共用）：附件文件夹不存在则自动创建；已存在则
' 提示核对（NOTICE_ONCE_PER_FOLDER 语义，按「入口 × 文件夹」会话级去重）。
' entryId = 入口标识（insert / folder / crossref）；purposeText 说明本次归档
' 内容（附件/底稿副本）。返回附件夹路径；用户取消或创建失败返回 ""。
Private Function EnsureAttachDirPrompted(ByVal fso As Object, ByVal wb As Object, _
        ByVal entryId As String, ByVal purposeText As String) As String
    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME
    If fso.FolderExists(attachDir) Then
        If AttachLink_FolderNoticeNeeded(attachDir, entryId) Then
            If MsgBox("当前底稿同级已存在「" & ATTACH_FOLDER_NAME & "」文件夹：" & vbCrLf & _
                      attachDir & vbCrLf & vbCrLf & _
                      purposeText & "，请先核对是否为目标目录。" & vbCrLf & vbCrLf & _
                      "「确定」继续选择　　「取消」退出本次插入", _
                      vbOKCancel + vbExclamation, "AttachLink") <> vbOK Then Exit Function
            AttachLink_FolderNoticeMark attachDir, entryId
        End If
    Else
        If Not EnsureFolder(fso, attachDir) Then
            MsgBox "无法创建附件文件夹：" & attachDir, vbCritical, "AttachLink"
            Exit Function
        End If
        AttachLink_FolderNoticeMark attachDir, entryId
        MsgBox "当前底稿同级不存在「" & ATTACH_FOLDER_NAME & "」文件夹，已自动创建：" & vbCrLf & attachDir, _
               vbInformation, "AttachLink"
    End If
    EnsureAttachDirPrompted = attachDir
End Function
Private Function EnsureFolder(ByVal fso As Object, ByVal p As String) As Boolean
    On Error Resume Next
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    EnsureFolder = fso.FolderExists(p)
End Function

' V6：链接地址正斜杠归一。Excel 会把保存时归一化失败的相对链接改写成
' 「附件/a.pdf」形态（绝对地址经 GetAbsolutePathName 后天然为反斜杠，无需处理）；
' 所有判断「附件\」前缀、计算相对子路径的入口统一经本函数。
Private Function NormSlash(ByVal addr As String) As String
    If InStr(addr, "/") > 0 Then NormSlash = Replace(addr, "/", "\") Else NormSlash = addr
End Function

' 判断是否为本加载项插入的附件链接：
' 相对地址以「附件\」开头，或绝对路径的父目录名为附件文件夹，
' 或 ..\ 形态相对地址（鼎信诺归档另存会把链接重基为 ..\账套\...\附件\x，WORKLOG #14）
' fsoIn / wbPathIn：可选传入调用方已建的 FSO 实例 / 底稿路径（用于解析 ..\ 相对形态）
Private Function IsAttachmentLink(ByVal hl As Object, Optional ByVal fsoIn As Object, Optional ByVal wbPathIn As String = "") As Boolean
    Dim a As String
    a = NormSlash(hl.Address)                 ' V6：附件/a.pdf 正斜杠形态同样识别
    If Len(a) = 0 Then Exit Function
    ' 快速预过滤：地址中不含附件文件夹名的不可能是附件链接
    If InStr(1, a, ATTACH_FOLDER_NAME, vbTextCompare) = 0 Then Exit Function
    If StrComp(Left$(a, Len(ATTACH_FOLDER_NAME) + 1), ATTACH_FOLDER_NAME & "\", vbTextCompare) = 0 Then
        IsAttachmentLink = True
    ElseIf InStr(a, ":") > 0 Or Left$(a, 2) = "\\" Then
        Dim fso As Object, parentDir As String
        If fsoIn Is Nothing Then
            Set fso = CreateObject("Scripting.FileSystemObject")
        Else
            Set fso = fsoIn
        End If
        parentDir = fso.GetParentFolderName(a)
        If Len(parentDir) > 0 Then
            If StrComp(fso.GetFileName(parentDir), ATTACH_FOLDER_NAME, vbTextCompare) = 0 Then IsAttachmentLink = True
        End If
    ElseIf Left$(a, 2) = ".." And Len(wbPathIn) > 0 Then
        ' ..\ 相对形态：基于底稿路径解析出完整路径后看父目录名（T12）
        Dim fso2 As Object, fullPath As String
        If fsoIn Is Nothing Then
            Set fso2 = CreateObject("Scripting.FileSystemObject")
        Else
            Set fso2 = fsoIn
        End If
        fullPath = ResolveLinkPath(fso2, wbPathIn, a)
        If Len(fullPath) > 0 Then
            parentDir = fso2.GetParentFolderName(fullPath)
            If Len(parentDir) > 0 Then
                If StrComp(fso2.GetFileName(parentDir), ATTACH_FOLDER_NAME, vbTextCompare) = 0 Then IsAttachmentLink = True
            End If
        End If
    End If
End Function

' 把地址解析为可检查的完整路径（绝对地址原样返回，相对地址基于底稿目录）
Private Function ResolveLinkPath(ByVal fso As Object, ByVal wbPath As String, ByVal addr As String) As String
    addr = NormSlash(addr)                    ' V6：附件/a.pdf → 附件\a.pdf 再解析
    If InStr(addr, ":") > 0 Or Left$(addr, 2) = "\\" Then
        ResolveLinkPath = addr
    Else
        ResolveLinkPath = wbPath & "\" & addr
    End If
End Function

' O-03：附件链接写操作的统一解析闸口——任何复制/改指/登记之前必须过本函数。
' 用 GetAbsolutePathName 规范化（大小写、重复分隔符、. 与 .. 全部解析），
' 用带目录边界的 IsSubPathOf 判定是否真在本底稿附件夹内（前缀匹配会把
' 「附件\..\outside.pdf」误判为夹内）。返回规范化完整路径；不在附件夹内
' 或 relSub 非法时返回 ""——调用方必须放弃一切写操作。
' relSub 输出：附件夹内相对子路径（供落盘路径、基线键与相对链接地址共用）。
Private Function ResolveAttachTarget(ByVal fso As Object, ByVal wbPath As String, _
                                     ByVal attachDir As String, ByVal addr As String, _
                                     ByRef relSub As String) As String
    relSub = ""
    Dim full As String
    full = fso.GetAbsolutePathName(ResolveLinkPath(fso, wbPath, addr))
    If Not IsSubPathOf(full, attachDir) Then Exit Function
    If StrComp(full, attachDir, vbTextCompare) = 0 Then Exit Function
    relSub = Mid$(full, Len(attachDir) + 2)
    ' 纵深防御（V11a 改按路径段判定）：文件名含连续两个点（a..b.pdf）合法，
    ' 只有规范化后仍存在「..」整段才算异常（GetAbsolutePathName 已消解真 .. 段，
    ' 此处是双保险）；也不得有盘符或前导分隔符
    Dim segs() As String, si As Long, badSeg As Boolean
    segs = Split(relSub, "\")
    For si = 0 To UBound(segs)
        If StrComp(segs(si), "..", vbTextCompare) = 0 Then badSeg = True: Exit For
    Next si
    If badSeg Or InStr(relSub, ":") > 0 Or Left$(relSub, 1) = "\" Then
        relSub = ""
        Exit Function
    End If
    ResolveAttachTarget = full
End Function

' V10：网页路径底稿（SharePoint/OneDrive「在桌面应用中打开」，wb.Path 以
' https:// 开头）下相对路径附件链接不可用。返回闸文案（"" = 可通过）；
' 各入口界面宏弹框、Core 返回「失败：…」，文案统一出自本函数。
Private Function WebPathGateText(ByVal pathText As String) As String
    If InStr(1, pathText, "://", vbTextCompare) > 0 Then
        WebPathGateText = "底稿保存在网页路径（SharePoint/OneDrive 网页版），相对路径附件链接不可用。" & vbCrLf & _
                          "请先把底稿另存到本地磁盘或局域网共享盘后再执行。"
    End If
End Function

Private Function IsWebPathWorkbook(ByVal wb As Object) As Boolean
    On Error Resume Next
    IsWebPathWorkbook = (Len(WebPathGateText(CStr(wb.Path))) > 0)
    Err.Clear
End Function

' V4：规范化完整路径是否在本次删除目标集合中（大小写不敏感）
Private Function PathInArray(ByVal fullPath As String, ByRef arr() As String) As Boolean
    Dim j As Long
    For j = 0 To UBound(arr)
        If StrComp(fullPath, CStr(arr(j)), vbTextCompare) = 0 Then
            PathInArray = True
            Exit Function
        End If
    Next j
End Function

' V4："地址|表名" 引用是否在本次删除链接集合中
Private Function RefInArray(ByVal ref As String, ByRef arr() As String) As Boolean
    Dim j As Long
    For j = 0 To UBound(arr)
        If StrComp(ref, CStr(arr(j)), vbTextCompare) = 0 Then
            RefInArray = True
            Exit Function
        End If
    Next j
End Function

' V11b：复制前快照源文件大小/修改时间（VerifyCopiedHash 的输入）
Private Sub SnapshotSrcFile(ByVal fso As Object, ByVal srcPath As String, _
                            ByRef sz As Variant, ByRef mt As Variant)
    On Error Resume Next
    Dim f As Object
    Set f = fso.GetFile(srcPath)
    If f Is Nothing Then
        sz = -1: mt = 0
    Else
        sz = f.Size
        mt = f.DateLastModified
    End If
    Err.Clear
End Sub

' N-06 复制后校验：调用方在复制前快照源大小/修改时间；复制后源被改动、
' 或目标大小与源不符（复制截断/被截获）时，源哈希已不代表目标内容——
' 对目标重新哈希返回；全部一致则沿用源哈希（不重复计算）。
Private Function VerifyCopiedHash(ByVal fso As Object, ByVal srcPath As String, ByVal dstPath As String, _
                                  ByVal srcSize As Variant, ByVal srcMTime As Variant, _
                                  ByVal srcHash As String) As String
    On Error Resume Next
    Dim ok As Boolean, sf As Object, df As Object
    ok = True
    Set sf = fso.GetFile(srcPath)
    Set df = fso.GetFile(dstPath)
    If sf Is Nothing Or df Is Nothing Then
        Err.Clear
        VerifyCopiedHash = ""
        Exit Function
    End If
    If sf.Size <> srcSize Then ok = False
    If ok Then
        If Abs(DateDiff("s", CDate(sf.DateLastModified), CDate(srcMTime))) > 1 Then ok = False
    End If
    If ok Then If df.Size <> srcSize Then ok = False
    ' Err 卫生：调用方靠 Err.Number 做逐链接故障隔离计数，函数不得遗留 Err
    Err.Clear
    If ok Then
        VerifyCopiedHash = srcHash
    Else
        VerifyCopiedHash = Sha256File(dstPath)
    End If
    Err.Clear
End Function

' 展开选区为单元格列表（行优先，合并格只保留左上角格）
Private Function FlattenSelection(ByVal sel As Object) As Collection
    Dim out As New Collection, area As Object, c As Object
    For Each area In sel.Areas
        For Each c In area.Cells
            If c.MergeCells Then
                If StrComp(c.MergeArea.Cells(1, 1).Address, c.Address, vbTextCompare) = 0 Then
                    out.Add c.MergeArea.Cells(1, 1)
                End If
            Else
                out.Add c
            End If
        Next c
    Next area
    Set FlattenSelection = out
End Function

' 向列表末尾所在列的下方追加一个空格（跳过非空格与合并格，最多向下找 200 行）
Private Function ExtendCellList(ByVal cellList As Collection) As Boolean
    ExtendCellList = False
    If cellList.Count = 0 Then Exit Function
    Dim lastCell As Object, r As Long, cand As Object
    Set lastCell = cellList(cellList.Count)
    For r = lastCell.Row + 1 To lastCell.Row + 200
        Set cand = lastCell.Worksheet.Cells(r, lastCell.Column)
        If Not cand.MergeCells Then
            If Application.WorksheetFunction.CountA(cand) = 0 Then
                cellList.Add cand
                ExtendCellList = True
                Exit Function
            End If
        End If
    Next r
End Function

' 归档判定（T2 附件身份核心，插入/底稿引用共用）：返回最终落盘完整路径
' （"" = 跳过/失败），category 输出 CAT_* 类别，detail 输出另存/复用的文件名。
' 规则：
'   - 源即目标（已在附件夹内）：不复制，确保基线已登记；
'   - 目标不存在：复制并【立即登记基线 v1】；
'   - 目标已存在：先比文件大小、再比 SHA-256——
'       内容相同 → 静默复用（不复制不弹窗），确保基线已登记；
'       内容不同 → 自动另存 主名(n).ext（n 从 2 起找第一个「不存在，或存在但
'                  内容相同」的名字；内容相同则复用该名字），不覆盖原文件、不弹窗；
'       哈希取不到 → 退回既有弹窗（覆盖/保留两者/跳过），绝不把「算不出哈希」
'                    当成「内容相同」。
' base/baseDirty：基线字典与脏标记（调用方整批一次落盘）；登记仅补缺失键。
Private Function ArchiveFileEx(ByVal fso As Object, ByVal srcPath As String, ByVal attachDir As String, _
                               ByVal base As Object, ByRef baseDirty As Boolean, _
                               ByRef category As Long, ByRef detail As String) As String
    category = 0
    detail = ""
    Dim fName As String, dest As String
    fName = fso.GetFileName(srcPath)
    dest = attachDir & "\" & fName
    ' 源文件本身就在附件文件夹中：不重复复制，确保登记
    If StrComp(srcPath, dest, vbTextCompare) = 0 Then
        Dim hIn As String
        hIn = HashFileCached(dest)
        ' O-05：夹内副本与既有基线不符（疑似被篡改）→ 不得复用扩散，另存新名
        If BaselineMismatch(base, LCase$(fName), hIn) Then
            ArchiveFileEx = SaveAsAnomalyCopy(fso, srcPath, attachDir, fName, category, detail, hIn)
            If Len(ArchiveFileEx) > 0 Then
                CacheHash ArchiveFileEx, hIn
                EnsureBaselineKey base, LCase$(fso.GetFileName(ArchiveFileEx)), hIn, baseDirty
            End If
            Exit Function
        End If
        category = CAT_INSITU
        EnsureBaselineKey base, LCase$(fName), hIn, baseDirty
        ArchiveFileEx = dest
        Exit Function
    End If
    If Not fso.FileExists(dest) Then
        Dim szN As Variant, mtN As Variant
        SnapshotSrcFile fso, srcPath, szN, mtN
        fso.CopyFile srcPath, dest, False
        ' 测试钩子（生产恒 0）：首次复制完成后向源追加字节（源大小变化），
        ' 验证登记以目标实际哈希为准（V11b）
        If testGrowAfterCopy > 0 Then
            testGrowAfterCopy = testGrowAfterCopy - 1
            On Error Resume Next
            Dim gFn As Integer
            gFn = FreeFile
            Open srcPath For Append As #gFn
            Print #gFn, "grow"
            Close #gFn
            Err.Clear
        End If
        category = CAT_NEW
        Dim hNew As String
        hNew = HashFileCached(srcPath)
        ' V11b：复制后校验——源被改动/被截获或目标大小不符时，源哈希已不代表
        ' 目标内容，以目标实际哈希登记
        hNew = VerifyCopiedHash(fso, srcPath, dest, szN, mtN, hNew)
        If Len(hNew) = 0 Then hNew = Sha256File(dest)
        CacheHash dest, hNew
        EnsureBaselineKey base, LCase$(fName), hNew, baseDirty
        ArchiveFileEx = dest
        Exit Function
    End If
    ' 目标已存在：先比大小，再比哈希
    Dim sameContent As Long    ' 1=相同 / 0=不同 / -1=无法判定（哈希不可用）
    sameContent = -1
    Dim hSrc As String, hDst As String
    If fso.GetFile(srcPath).Size <> fso.GetFile(dest).Size Then
        sameContent = 0                                  ' 大小不同必为异内容，无需哈希
    Else
        hSrc = HashFileCached(srcPath)
        hDst = HashFileCached(dest)
        If Len(hSrc) > 0 And Len(hDst) > 0 Then
            If StrComp(hSrc, hDst, vbTextCompare) = 0 Then sameContent = 1 Else sameContent = 0
        End If
    End If
    If sameContent = 1 Then
        ' O-05：复用前做三方哈希校验——copy==baseline==source 才允许复用；
        ' 副本与既有基线不符（疑似被篡改，体检会标红）时不得复用扩散，另存新名
        If BaselineMismatch(base, LCase$(fName), hDst) Then
            ArchiveFileEx = SaveAsAnomalyCopy(fso, srcPath, attachDir, fName, category, detail, hSrc)
            If Len(ArchiveFileEx) > 0 Then
                CacheHash ArchiveFileEx, hSrc
                EnsureBaselineKey base, LCase$(fso.GetFileName(ArchiveFileEx)), hSrc, baseDirty
            End If
            Exit Function
        End If
        category = CAT_REUSED
        EnsureBaselineKey base, LCase$(fName), hDst, baseDirty
        ArchiveFileEx = dest
        Exit Function
    End If
    If sameContent = 0 Then
        ' 同名异内容：自动另存 主名(n).ext（不覆盖原文件、不弹窗）
        Dim baseName As String, extName As String, n As Long, cand As String, hCand As String
        baseName = fso.GetBaseName(fName)
        extName = fso.GetExtensionName(fName)
        n = 2
        Do
            cand = attachDir & "\" & baseName & "(" & n & ")" & IIf(Len(extName) > 0, "." & extName, "")
            If Not fso.FileExists(cand) Then
                Dim szR As Variant, mtR As Variant
                SnapshotSrcFile fso, srcPath, szR, mtR
                fso.CopyFile srcPath, cand, False
                If Len(hSrc) = 0 Then hSrc = HashFileCached(srcPath)
                ' V11b：(n) 另存接复制后校验——异常时以目标实际哈希登记
                hSrc = VerifyCopiedHash(fso, srcPath, cand, szR, mtR, hSrc)
                If Len(hSrc) = 0 Then hSrc = Sha256File(cand)
                CacheHash cand, hSrc
                EnsureBaselineKey base, LCase$(fso.GetFileName(cand)), hSrc, baseDirty
                category = CAT_RENAMED
                detail = fso.GetFileName(cand)
                ArchiveFileEx = cand
                Exit Function
            End If
            ' 候选已存在：内容相同且与基线三方一致则复用该名字（O-05：基线不符不扩散）；
            ' 大小不同直接判异，免哈希
            If fso.GetFile(cand).Size = fso.GetFile(srcPath).Size Then
                If Len(hSrc) = 0 Then hSrc = HashFileCached(srcPath)
                hCand = HashFileCached(cand)
                If Len(hSrc) > 0 And Len(hCand) > 0 Then
                    If StrComp(hCand, hSrc, vbTextCompare) = 0 Then
                        If BaselineMismatch(base, LCase$(fso.GetFileName(cand)), hCand) Then GoTo NextCand
                        EnsureBaselineKey base, LCase$(fso.GetFileName(cand)), hCand, baseDirty
                        category = CAT_REUSED
                        detail = fso.GetFileName(cand)
                        ArchiveFileEx = cand
                        Exit Function
                    End If
                End If
            End If
NextCand:
            n = n + 1
        Loop While n < 1000
        ' 极端：1000 个候选全为异内容——拒绝继续，绝不覆盖
        category = 0
        ArchiveFileEx = ""
        Exit Function
    End If
    ' 哈希取不到：退回既有弹窗（覆盖/保留两者/跳过），绝不把「算不出哈希」当成「内容相同」
    Dim popResult As String
    popResult = ArchiveFile(fso, srcPath, attachDir)
    category = CAT_POPUP
    If Len(popResult) > 0 Then
        If StrComp(popResult, dest, vbTextCompare) <> 0 Then detail = fso.GetFileName(popResult)
    End If
    ArchiveFileEx = popResult
End Function

' 归档复制（哈希不可用时的弹窗回退路径）：返回最终落盘路径；重名时弹窗确认
' （覆盖 / 自动重命名 / 跳过）。返回 "" 表示用户选择跳过
Private Function ArchiveFile(ByVal fso As Object, ByVal srcPath As String, ByVal attachDir As String) As String
    Dim fName As String, dest As String
    fName = fso.GetFileName(srcPath)
    dest = attachDir & "\" & fName
    ' 源文件本身就在附件文件夹中：不重复复制
    If StrComp(srcPath, dest, vbTextCompare) = 0 Then
        ArchiveFile = dest
        Exit Function
    End If
    If Not fso.FileExists(dest) Then
        fso.CopyFile srcPath, dest, False
        ArchiveFile = dest
        Exit Function
    End If
    Dim ans As VbMsgBoxResult
    ans = MsgBox("附件已存在同名文件：" & vbCrLf & fName & vbCrLf & vbCrLf & _
                 "「是」覆盖旧文件　　「否」保留两者（自动重命名）　　「取消」跳过此文件", _
                 vbYesNoCancel + vbQuestion, "AttachLink")
    If ans = vbYes Then
        fso.CopyFile srcPath, dest, True
        ArchiveFile = dest
    ElseIf ans = vbNo Then
        Dim baseName As String, extName As String, n As Long
        baseName = fso.GetBaseName(fName)
        extName = fso.GetExtensionName(fName)
        n = 2
        Do
            dest = attachDir & "\" & baseName & "(" & n & ")" & IIf(Len(extName) > 0, "." & extName, "")
            n = n + 1
        Loop While fso.FileExists(dest)
        fso.CopyFile srcPath, dest, False
        ArchiveFile = dest
    Else
        ArchiveFile = ""
    End If
End Function

Private Sub InsertLink(ByVal targetCell As Object, ByVal relAddr As String, ByVal disp As String)
    ' 目标格已含超链接时 Hyperlinks.Add 会报 1004，先移除旧链接（保持“一格一链接”约定）
    If targetCell.Hyperlinks.Count > 0 Then targetCell.Hyperlinks(1).Delete
    targetCell.Hyperlinks.Add Anchor:=targetCell, Address:=relAddr, TextToDisplay:=disp
End Sub

Private Function DisplayName(ByVal fName As String, ByVal fso As Object) As String
    If SHOW_EXTENSION Then
        DisplayName = fName
    Else
        DisplayName = fso.GetBaseName(fName)
    End If
    If Len(DisplayName) = 0 Then DisplayName = fName
End Function

' 取得（或创建）附件清单表。
' 同名表存在但缺少本工具归属标记（CustomProperties，WPS 不支持时回退
' sheet 级名称标记 _AttachLinkIndex）时返回 Nothing，
' 防止整表重建误清用户或鼎信诺生成的同名表。
Private Function GetOrCreateIndexSheet(ByVal wb As Object) As Object
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(INDEX_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = INDEX_SHEET_NAME
        MarkIndexSheet ws
    Else
        If Not IndexSheetMarked(ws) Then
            Exit Function
        End If
    End If
    Set GetOrCreateIndexSheet = ws
End Function

' 写入归属标记：CustomProperties（Excel）+ sheet 级名称（WPS 兼容回退），两者都试
Private Sub MarkIndexSheet(ByVal ws As Object)
    On Error Resume Next
    ws.CustomProperties.Add "AttachLinkIndex", "1"
    Err.Clear
    ws.Names.Add "_AttachLinkIndex", "=""1"""
    Err.Clear
    On Error GoTo 0
End Sub

' 读取归属标记：CustomProperties 优先；宿主不支持或未写入时查名称标记
Private Function IndexSheetMarked(ByVal ws As Object) As Boolean
    Dim cp As Object
    On Error Resume Next
    For Each cp In ws.CustomProperties
        If StrComp(cp.Name, "AttachLinkIndex", vbTextCompare) = 0 Then
            IndexSheetMarked = True
            Exit Function
        End If
    Next cp
    Err.Clear
    Dim nm As Object
    Set nm = Nothing
    Set nm = ws.Names("_AttachLinkIndex")
    IndexSheetMarked = Not (nm Is Nothing)
End Function

' 归档迁移：把工作簿中所有「指向别处」的附件链接改写为相对地址（附件\文件名）。
' 关键实证：鼎信诺另存归档时，Excel 会把链接重基为 ../ 相对路径指回账套原件
' （也可能保持绝对路径），两种形态都要处理——凡解析后不落在本底稿同级附件夹
' 内的附件链接，一律按文件名改写为「附件\文件名」（与本工具「附件平铺、
' 一格一链」的写入形态一致）。
' 返回改写数量；遇到受保护工作表等错误时抛出，由调用方记入失败清单。
Private Function RewriteAbsoluteAttachLinks(ByVal wb As Object, ByVal fso As Object) As Long
    Dim ws As Object, hl As Object, a As String, full As String
    Dim parentDir As String, localAttach As String, n As Long
    On Error GoTo Fail
    localAttach = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)
    For Each ws In wb.Worksheets
        For Each hl In ws.Hyperlinks
            a = hl.Address
            If Len(a) > 0 Then
                If InStr(1, a, ATTACH_FOLDER_NAME, vbTextCompare) > 0 Then
                    ' O-03：先规范化再判断父目录（.. 与重复分隔符一律解析，防误判）
                    full = fso.GetAbsolutePathName(ResolveLinkPath(fso, wb.Path, a))
                    parentDir = fso.GetParentFolderName(full)
                    If Len(parentDir) > 0 Then
                        If StrComp(fso.GetFileName(parentDir), ATTACH_FOLDER_NAME, vbTextCompare) = 0 Then
                            If StrComp(parentDir, localAttach, vbTextCompare) <> 0 Then
                                ' V5：只改目标文件确实已出现在归档附件夹里的链接——
                                ' 移动失败的保留原地址（迁移方汇总单列，不是断链漏报）
                                Dim relUnder As String
                                relUnder = Mid$(full, Len(parentDir) + 2)
                                If fso.FileExists(localAttach & "\" & relUnder) Then
                                    hl.Address = ATTACH_FOLDER_NAME & "\" & relUnder
                                    n = n + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next hl
    Next ws
    RewriteAbsoluteAttachLinks = n
    Exit Function
Fail:
    Dim ctx As String
    If Not ws Is Nothing Then ctx = ws.Name & "："
    Err.Raise vbObjectError + 513, , ctx & Err.Description
End Function

' 归档迁移：递归把源目录全部文件【移动】到目标目录（目标已有同名文件时
' 删除旧文件后移入，即以源=账套原件为准）。全部移空后仅删除空目录
' （非空=有移动失败残留，原样保留，绝不递归强删）。
' 返回成功移动的文件数；单文件失败记入 errs 不中断，rel 同时记入 moveFail
' （V5：迁移方后续据此不改写这些文件的链接）。
Private Function MoveTreeInto(ByVal fso As Object, ByVal srcDir As String, ByVal dstDir As String, _
                              ByVal errs As Collection, ByRef moveFail As Collection) As Long
    Dim f As Object, sf As Object, rel As String, dstPath As String, n As Long
    On Error GoTo Fail
    If Not EnsureFolder(fso, dstDir) Then
        errs.Add dstDir & "：无法创建目标文件夹"
        Exit Function
    End If
    For Each f In fso.GetFolder(srcDir).Files
        rel = Mid$(f.Path, Len(srcDir) + 2)
        dstPath = dstDir & "\" & rel
        On Error Resume Next
        If fso.FileExists(dstPath) Then fso.DeleteFile dstPath, True
        fso.MoveFile f.Path, dstPath
        If Err.Number <> 0 Then
            errs.Add rel & "：移动失败 - " & Err.Description
            moveFail.Add rel
            Err.Clear
        Else
            n = n + 1
        End If
        On Error GoTo Fail
    Next f
    For Each sf In fso.GetFolder(srcDir).SubFolders
        n = n + MoveTreeInto(fso, sf.Path, dstDir & "\" & sf.Name, errs, moveFail)
    Next sf
    ' 子目录递归返回后，源目录已无子目录且无文件才删除（留下空壳无害，但顺手清理）
    On Error Resume Next
    If (fso.GetFolder(srcDir).Files.Count = 0) And (fso.GetFolder(srcDir).SubFolders.Count = 0) Then
        fso.DeleteFolder srcDir, False
    End If
    On Error GoTo 0
    MoveTreeInto = n
    Exit Function
Fail:
    errs.Add srcDir & "：读取失败 - " & Err.Description
End Function

' 归档迁移：扫描工作簿中绝对地址的附件链接，返回出现次数最多的附件目录
' （带尾部反斜杠，作为源目录选择器的建议值）；没有则返回空串
Private Function SuggestAttachSource(ByVal wb As Object, ByVal fso As Object) As String
    On Error Resume Next
    Dim ws As Object, hl As Object, dirs As Object, p As String
    Set dirs = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        For Each hl In ws.Hyperlinks
            If IsAttachmentLink(hl, fso, wb.Path) Then
                If InStr(hl.Address, ":") > 0 Or Left$(hl.Address, 2) = "\\" Then
                    p = LCase$(fso.GetParentFolderName(hl.Address))
                    If Len(p) > 0 Then dirs(p) = dirs(p) + 1
                End If
            End If
        Next hl
    Next ws
    Dim k As Variant, best As String, bestN As Long
    For Each k In dirs.Keys
        If dirs(k) > bestN Then
            bestN = dirs(k)
            best = k
        End If
    Next k
    If Len(best) > 0 Then SuggestAttachSource = best & "\"
End Function

' 按完整路径在已打开工作簿中查找（找不到返回 Nothing）。
' 归档迁移用它区分「归档件已打开（就地处理）」与「账套同名原件已打开（互不影响）」
Private Function WorkbookByPath(ByVal fullPath As String) As Object
    Dim w As Object
    For Each w In Application.Workbooks
        If StrComp(w.FullName, fullPath, vbTextCompare) = 0 Then
            Set WorkbookByPath = w
            Exit Function
        End If
    Next w
End Function

' 清除单个单元格上的附件链接；清完无残留链接时恢复字体格式
Private Function ClearCellLinks(ByVal c As Object) As Long
    Dim wb As Object: Set wb = ActiveWorkbook
    Dim i As Long, removed As Long
    For i = c.Hyperlinks.Count To 1 Step -1
        If IsAttachmentLink(c.Hyperlinks(i), , wb.Path) Then
            c.Hyperlinks(i).Delete
            removed = removed + 1
        End If
    Next i
    If removed > 0 Then
        If c.Hyperlinks.Count = 0 Then
            With c.Font
                .Underline = (-4142)
                .ColorIndex = (-4105)
            End With
        End If
    End If
    ClearCellLinks = removed
End Function
