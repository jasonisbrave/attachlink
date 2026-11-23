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
' True = 「附件文件夹已存在」提示对同一文件夹每次会话仅弹一次；False = 每次插入都提示
Public Const NOTICE_ONCE_PER_FOLDER As Boolean = True
' 快捷键：插入附件链接 / 附件清单（改为 "" 可禁用）
Public Const HOTKEY_INSERT As String = "^+a"    ' Ctrl+Shift+A
Public Const HOTKEY_INDEX As String = "^+k"     ' Ctrl+Shift+K
' 附件清单工作表名
Public Const INDEX_SHEET_NAME As String = "附件清单"

' 版本号（单一事实源：安装器/看门狗日志/支持排障都以此为准；语义化递增）
' 1.4.0：附件删除（删链接+删文件+基线清理）/ 底稿引用（跨底稿交叉索引）/
'       IsAttachmentLink 识别 ..\ 形态 / 批量操作逐文件故障隔离
'       （导出交付包曾实现后于 2026-11-23 按用户决策下线（代码归档移出，回退任务 T1）；
'        配置外置 INI 未实施，配置仍为源码常量）
' 1.3.1：右键菜单再浓缩（4 顶项：归档子菜单收编体检/断链修复，归档前后步骤统一入口）
' 1.3.0：WPS 表格支持（全量晚绑定改造，需宿主启用 VBA 环境）+ 重命名多选批量 + 安装器墓碑进程误报修复
' 1.2.1：风险处置（Issue #1）——看门狗父进程白名单（配置化默认开放）+ 装载版本日志
' 1.2.0：文件夹引用（复制入附件夹+夹内基线）/ 重命名链接显示名 / 插入对话框记忆上次源目录 / 右键菜单二级化
' ■ 发布流程约定：版本发布（GitHub Release + git push）必须先经用户确认
Public Const ATTACHLINK_VERSION As String = "1.4.2"
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
' 上次已提示过「已存在」的附件文件夹路径（会话级记忆，避免重复打扰）
Private lastNotifiedAttachDir As String
' 上次选取附件/文件夹的源目录（会话级记忆：插入对话框优先定位到这里，不再每次落在附件夹）
Private lastSourceDir As String

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
        btn.Caption = "重命名链接显示名"
        btn.OnAction = "AttachLink_RenameLink"
        btn.Tag = MENU_TAG

        Set btn = cb.Controls.Add(1, , , 4, True)
        btn.Caption = "插入底稿引用"
        btn.OnAction = "AttachLink_InsertCrossRef"
        btn.Tag = MENU_TAG

        ' 二级菜单：归档（迁移 / 体检 / 修复 / 版本更新 / 删除——归档前后步骤统一入口，1.3.1 浓缩，1.4.0 加删除）
        Dim pop As Object, subBtn As Object
        Set pop = cb.Controls.Add(10, , , 5, True)
        pop.Caption = "归档"
        pop.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 1, True)
        subBtn.Caption = "归档迁移（批量）"
        subBtn.OnAction = "AttachLink_ArchiveMigrate"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 2, True)
        subBtn.Caption = "附件清单 / 体检"
        subBtn.OnAction = "AttachLink_BuildIndex"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 3, True)
        subBtn.Caption = "附件断链修复"
        subBtn.OnAction = "AttachLink_Repair"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(1, , , 4, True)
        subBtn.Caption = "附件版本更新"
        subBtn.OnAction = "AttachLink_UpdateVersion"
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
    attachDir = EnsureAttachDirPrompted(fso, wb)
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

    Dim cellList As Collection
    Set cellList = FlattenSelection(Selection)

    Dim inserted As Long, skipped As Long, idx As Long, failed As Long
    Dim i As Long
    Dim destPath As String, targetCell As Object, failList As String

    For i = 1 To fd.SelectedItems.Count
        ' T4：逐文件故障隔离——单个文件处理失败不中断整批
        On Error Resume Next
        Err.Clear
        destPath = ""
        destPath = ArchiveFile(fso, fd.SelectedItems(i), attachDir)
        If Err.Number <> 0 Then
            failed = failed + 1
            failList = failList & "  " & fso.GetFileName(fd.SelectedItems(i)) & "：" & Err.Description & vbCrLf
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
                    InsertLink targetCell, _
                               ATTACH_FOLDER_NAME & "\" & fso.GetFileName(destPath), _
                               DisplayName(fso.GetFileName(destPath), fso)
                    idx = idx + 1
                    inserted = inserted + 1
                End If
            End If
        End If
    Next i

    Dim msg As String
    msg = "已插入 " & inserted & " 个附件链接，文件已归档到同级「" & ATTACH_FOLDER_NAME & "」文件夹。"
    If skipped > 0 Then msg = msg & vbCrLf & "跳过 " & skipped & " 个文件（已取消或可用单元格不足）。"
    If failed > 0 Then
        msg = msg & vbCrLf & "失败 " & failed & " 个文件：" & vbCrLf & failList
    End If
    MsgBox msg, IIf(failed > 0, vbExclamation, vbInformation), "AttachLink"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "插入附件链接时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

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
    MsgBox AttachLink_BuildIndexCore(wb), vbInformation, "AttachLink 附件清单"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "生成附件清单时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 清单核心（独立公开函数：便于自动化测试），返回汇总文本
Public Function AttachLink_BuildIndexCore(ByVal wb As Object) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME

    ' 内容基线（SHA-256 经 WinCrypt；单文件失败显示「校验不可用」，不影响清单）
    Dim base As Object
    Set base = LoadBaseline(fso, attachDir)
    Dim baseDirty As Boolean

    Application.ScreenUpdating = False
    Dim wsIdx As Object
    Set wsIdx = GetOrCreateIndexSheet(wb)
    If wsIdx Is Nothing Then
        Application.ScreenUpdating = True
        AttachLink_BuildIndexCore = "已存在同名「" & INDEX_SHEET_NAME & "」工作表，但它不是本加载项生成的（缺少归属标记）。" & vbCrLf & _
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
    wsIdx.Range("A3:G3").Value = Array("编号", "附件名", "相对路径", "所在位置", "状态", "内容基线", "更替记录")

    Dim ws As Object, hl As Object
    Dim r As Long, total As Long, brokenCnt As Long, changedCnt As Long, unavailCnt As Long
    Dim cnt As Long, g As Long
    Dim linkPath As String, ok As Boolean
    Dim h As String, key As String, cellF As String, cellG As String, meta() As String

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
                        wsIdx.Cells(r, 1).Value = ws.Name & "-" & Format(g, "00")
                        wsIdx.Cells(r, 2).Value = hl.TextToDisplay
                        wsIdx.Cells(r, 3).Value = hl.Address
                        wsIdx.Cells(r, 4).Value = ws.Name & "!" & hl.Range.Address(False, False)
                        linkPath = ResolveLinkPath(fso, wb.Path, hl.Address)
                        ok = fso.FileExists(linkPath)
                        cellF = "—"
                        cellG = ""
                        Dim isFolder As Boolean
                        isFolder = False
                        If Not ok Then isFolder = fso.FolderExists(linkPath)
                        If isFolder Then
                            wsIdx.Cells(r, 5).Value = "文件夹"
                            wsIdx.Cells(r, 5).Font.Color = RGB(0, 102, 204)
                            If StrComp(Left$(linkPath, Len(attachDir)), attachDir, vbTextCompare) <> 0 Then
                                ' 外部文件夹（如归档前的指回链接）：仅查存在性，不校验基线（键无从落在本附件夹内）
                                cellF = "外部，未校验"
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
                            ' 内容基线校验（SHA-256；值=哈希|版本|登记时间|更新时间）
                            h = Sha256File(linkPath)
                            ' 键=附件夹内相对子路径；指向外部附件夹的链接
                            ' （迁移前形态）按文件名兜底
                            If StrComp(Left$(linkPath, Len(attachDir)), attachDir, vbTextCompare) = 0 Then
                                key = LCase$(Mid$(linkPath, Len(attachDir) + 2))
                            Else
                                key = LCase$(fso.GetFileName(linkPath))
                            End If
                            cellG = ""
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
                        Else
                            wsIdx.Cells(r, 5).Value = "断链"
                            wsIdx.Cells(r, 5).Font.Color = RGB(192, 0, 0)
                            wsIdx.Cells(r, 5).Font.Bold = True
                            brokenCnt = brokenCnt + 1
                        End If
                        wsIdx.Cells(r, 6).Value = cellF
                        If Left$(cellF, 6) = "■ 内容变更" Then
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
                    End If
                Next hl
            End If
        End If
    Next ws

    ' 基线落盘（有新登记才写）；失败时汇总警告（登记未持久化，下次体检将重新登记）
    Dim persistFail As Boolean
    If baseDirty Then
        If Not SaveBaseline(fso, attachDir, base) Then persistFail = True
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

    Application.ScreenUpdating = True
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
        Else
            msg = msg & "，内容基线全部一致。"
        End If
    End If
    If persistFail Then
        msg = msg & vbCrLf & "■ 基线文件写入失败：本次新登记未持久化（" & BASELINE_FILE & "），下次体检将重新登记，请检查附件夹写权限。"
    End If
    AttachLink_BuildIndexCore = msg
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

    ' 收集断链（oracle #1：文件夹的 FileExists 恒为 False，必须同时排除存在的文件夹，
    ' 否则健康文件夹会被误列入修复并用源目录合并覆盖）
    Dim broken As Collection
    Set broken = New Collection
    Dim ws As Object, hl As Object, rp As String
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, INDEX_SHEET_NAME, vbTextCompare) <> 0 Then
            For Each hl In ws.Hyperlinks
                If IsAttachmentLink(hl, fso, wb.Path) Then
                    rp = ResolveLinkPath(fso, wb.Path, hl.Address)
                    If Not fso.FileExists(rp) And Not fso.FolderExists(rp) Then broken.Add hl
                End If
            Next hl
        End If
    Next ws
    If broken.Count = 0 Then
        MsgBox "未发现断链附件，无需修复。", vbInformation, "AttachLink"
        Exit Sub
    End If

    ' 选择附件文件夹当前位置
    Dim fd As Object
    Set fd = Application.FileDialog(4)
    With fd
        .Title = "选择附件文件当前所在的文件夹（共 " & broken.Count & " 个断链待修复）"
        .InitialFileName = wb.Path & "\"
        If .Show <> -1 Then Exit Sub
    End With
    Dim srcDir As String
    srcDir = fd.SelectedItems(1)
    If Right$(srcDir, 1) = "\" Then srcDir = Left$(srcDir, Len(srcDir) - 1)

    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME
    EnsureFolder fso, attachDir

    Dim fixedCnt As Long, missCnt As Long, i As Long
    Dim fName As String, srcPath As String, finalPath As String
    For i = 1 To broken.Count
        Set hl = broken(i)
        fName = fso.GetFileName(hl.Address)
        srcPath = srcDir & "\" & fName
        If fso.FileExists(srcPath) Then
            finalPath = ArchiveFile(fso, srcPath, attachDir)
            If Len(finalPath) > 0 Then
                hl.Address = ATTACH_FOLDER_NAME & "\" & fso.GetFileName(finalPath)
                fixedCnt = fixedCnt + 1
            Else
                missCnt = missCnt + 1
            End If
        ElseIf fso.FolderExists(srcPath) Then
            ' 文件夹引用的断链重连：整体复制（合并覆盖）到附件夹
            On Error Resume Next
            fso.CopyFolder srcPath, attachDir & "\" & fName, True
            If Err.Number <> 0 Then
                Err.Clear
                missCnt = missCnt + 1
            Else
                hl.Address = ATTACH_FOLDER_NAME & "\" & fName
                fixedCnt = fixedCnt + 1
            End If
            On Error GoTo Fail
        Else
            missCnt = missCnt + 1
        End If
    Next i

    MsgBox "修复完成：重连 " & fixedCnt & " 个，未找到 " & missCnt & " 个。" & vbCrLf & _
           "建议运行「附件清单 / 体检」核对状态。", vbInformation, "AttachLink"
    Exit Sub
Fail:
    MsgBox "修复时出错：" & Err.Description, vbCritical, "AttachLink"
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
Public Sub NormalizeAttachmentLinks(ByVal wb As Object)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME
    Dim ws As Object, hl As Object, parentDir As String
    For Each ws In wb.Worksheets
        For Each hl In ws.Hyperlinks
            If IsAttachmentLink(hl, fso, wb.Path) Then
                If InStr(hl.Address, ":") > 0 Or Left$(hl.Address, 2) = "\\" Then
                    parentDir = fso.GetParentFolderName(hl.Address)
                    If StrComp(parentDir, attachDir, vbTextCompare) = 0 Then
                        hl.Address = ATTACH_FOLDER_NAME & "\" & fso.GetFileName(hl.Address)
                    End If
                End If
            End If
        Next hl
    Next ws
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
    attachDir = EnsureAttachDirPrompted(fso, wb)
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
    If fso.FolderExists(destDir) Then
        If MsgBox("附件夹内已存在同名文件夹：" & fName & vbCrLf & vbCrLf & _
                  "「是」合并覆盖（同名文件以新复制的为准）　　「否」取消", _
                  vbYesNo + vbExclamation, "AttachLink") <> vbYes Then Exit Sub
    End If

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

    ' 链接插入当前选区的第一个可用格（沿用插入的扩展约定：一格一链接）
    Dim cellList As Collection
    Set cellList = FlattenSelection(Selection)
    If cellList.Count = 0 Then
        MsgBox "选区内没有可放置链接的单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If
    InsertLink cellList(1), ATTACH_FOLDER_NAME & "\" & fName, fName
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "插入文件夹引用时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 文件夹引用核心（独立公开函数：便于自动化测试）。以当前活动工作簿所在
' 目录为基准，复制源文件夹到附件夹（合并覆盖）并把夹内全部文件登记入
' 内容基线（仅登记缺失键，不改既有版本）。返回汇总文本。
Public Function AttachLink_InsertFolderCore(ByVal srcDir As String) As String
    Dim wbHost As Object
    Set wbHost = ActiveWorkbook
    If wbHost Is Nothing Then
        AttachLink_InsertFolderCore = "失败：没有活动工作簿。"
        Exit Function
    End If
    If Len(wbHost.Path) = 0 Then
        AttachLink_InsertFolderCore = "失败：活动工作簿尚未保存。"
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcDir) Then
        AttachLink_InsertFolderCore = "失败：源文件夹不存在：" & srcDir
        Exit Function
    End If
    srcDir = fso.GetFolder(srcDir).Path                ' oracle #5：Core 直调同样规范化（含大小写归一）

    Dim attachDir As String, fName As String, destDir As String
    attachDir = wbHost.Path & "\" & ATTACH_FOLDER_NAME
    fName = fso.GetFileName(srcDir)
    destDir = attachDir & "\" & fName
    ' 路径防护（oracle #4，Core 直调同样校验）
    If Len(fso.GetParentFolderName(srcDir)) = 0 Then
        AttachLink_InsertFolderCore = "失败：不能引用磁盘根目录。"
        Exit Function
    End If
    If IsSubPathOf(destDir, srcDir) Or StrComp(destDir, srcDir, vbTextCompare) = 0 Then
        AttachLink_InsertFolderCore = "失败：目标位置不能位于源文件夹内部（或与源相同）。"
        Exit Function
    End If
    If IsSubPathOf(srcDir, attachDir) Then
        AttachLink_InsertFolderCore = "失败：源文件夹已位于本底稿附件夹内，无需引用。"
        Exit Function
    End If
    If Not EnsureFolder(fso, attachDir) Then
        AttachLink_InsertFolderCore = "失败：无法创建附件文件夹：" & attachDir
        Exit Function
    End If
    If Not EnsureFolder(fso, destDir) Then
        AttachLink_InsertFolderCore = "失败：无法创建目标文件夹：" & destDir
        Exit Function
    End If
    On Error GoTo Fail
    fso.CopyFolder srcDir, destDir, True

    ' 夹内全部文件登记基线（仅补缺失键；已有键的版本与哈希不动）。
    ' 登记失败（哈希/遍历错误）时如实报告部分完成（oracle #2）
    Dim base As Object
    Set base = LoadBaseline(fso, attachDir)
    Dim regCnt As Long, fileCnt As Long, failCnt As Long
    RegisterFolderFilesBaseline fso, base, destDir, attachDir, regCnt, fileCnt, failCnt
    If Not SaveBaseline(fso, attachDir, base) Then
        AttachLink_InsertFolderCore = "失败：基线文件写入不成功（文件夹已复制，但登记未持久化，体检将重新登记） - " & attachDir & "\" & BASELINE_FILE
        Exit Function
    End If

    Dim tail As String
    If failCnt > 0 Then tail = vbCrLf & "■ 有 " & failCnt & " 个文件登记失败（不可读/被占用），体检将显示「校验不可用」"
    AttachLink_InsertFolderCore = "文件夹引用完成：" & vbCrLf & _
        "· 已复制到 " & destDir & "（共 " & fileCnt & " 个文件）" & vbCrLf & _
        "· 新登记内容基线 " & regCnt & " 个文件（夹内篡改可被体检检测）" & tail & vbCrLf & _
        "· 链接指向附件夹内副本，点击即在资源管理器打开"
    Exit Function
Fail:
    AttachLink_InsertFolderCore = "失败：插入文件夹引用未完成 - " & Err.Description
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
            h = Sha256File(f.Path)
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

    ' 执行交给 Core（无对话框，供自动化测试直调）：先删文件/文件夹，再删链接，最后清基线
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
    If InStr(summary, "部分失败") > 0 Or InStr(summary, "失败：") > 0 Then msgIcon = vbExclamation Else msgIcon = vbInformation
    MsgBox summary, msgIcon, "AttachLink"
    Exit Sub
Fail:
    MsgBox "附件删除时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

' 删除执行核心（无对话框，供自动化测试直调）：删文件/文件夹 → 删链接清样式 → 清基线，
' 返回汇总文本（含「部分失败」/「失败：」字样时 wrapper 以警示图标显示）。
' linksJoined / pathsJoined 为按行拼接的 "地址|表名" / 完整路径串（vbLf 分隔，两侧一一对应）。
' 安全契约（oracle 复核 #1）：不信任调用方传入的路径——附件夹由 wb 自身推导，
' 全部路径先经一次性预检（附件夹内、非根、无控制字符、数据一致），任一项
' 不合法即整体拒绝、零删除；首项合法后项越界同样不得删除首项。
Public Function AttachLink_DeleteAttachmentCore(ByVal wb As Object, ByVal linksJoined As String, _
    ByVal pathsJoined As String, Optional ByVal fsoIn As Object) As String

    Dim fso As Object
    If fsoIn Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    Else
        Set fso = fsoIn
    End If

    If wb Is Nothing Then
        AttachLink_DeleteAttachmentCore = "失败：未提供工作簿，未执行任何删除。"
        Exit Function
    End If
    If Len(wb.Path) = 0 Then
        AttachLink_DeleteAttachmentCore = "失败：底稿未保存，无法定位附件文件夹，未执行任何删除。"
        Exit Function
    End If

    ' 附件夹从 wb 自推导（不信任调用方），GetAbsolutePathName 规范化防 ../ 与重复分隔符
    Dim attachDir As String
    attachDir = fso.GetAbsolutePathName(wb.Path & "\" & ATTACH_FOLDER_NAME)

    Dim linkArr() As String, pathArr() As String
    linkArr = Split(linksJoined, vbLf)
    pathArr = Split(pathsJoined, vbLf)
    If UBound(linkArr) < 0 Or UBound(linkArr) <> UBound(pathArr) Then
        AttachLink_DeleteAttachmentCore = "失败：删除内部数据不一致，未执行任何删除。"
        Exit Function
    End If

    ' ---- 预检：全部通过才开始删除；任一项不合法 → 整体拒绝、零删除 ----
    Dim i As Long, p As Variant, norm As String
    For i = 0 To UBound(pathArr)
        p = pathArr(i)
        If Len(CStr(p)) = 0 Then
            AttachLink_DeleteAttachmentCore = "失败：第 " & (i + 1) & " 项路径为空，未执行任何删除。"
            Exit Function
        End If
        ' 拒绝换行/NUL 注入（不用 Asc 逐字符扫描：中文字符的 ANSI 码转换会假阳性）
        If InStr(CStr(p), vbCr) > 0 Or InStr(CStr(p), vbLf) > 0 Or InStr(CStr(p), Chr$(0)) > 0 Then
            AttachLink_DeleteAttachmentCore = "失败：第 " & (i + 1) & " 项路径含换行/NUL 字符，未执行任何删除。"
            Exit Function
        End If
        norm = fso.GetAbsolutePathName(CStr(p))
        If Not IsSubPathOf(norm, attachDir) Or StrComp(norm, attachDir, vbTextCompare) = 0 Then
            AttachLink_DeleteAttachmentCore = "失败：第 " & (i + 1) & " 项越界（不在当前底稿附件文件夹内或为附件夹根），未执行任何删除：" & vbCrLf & norm
            Exit Function
        End If
        pathArr(i) = norm
    Next i
    For i = 0 To UBound(linkArr)
        If InStr(CStr(linkArr(i)), "|") = 0 Then
            AttachLink_DeleteAttachmentCore = "失败：第 " & (i + 1) & " 项链接引用格式非法，未执行任何删除。"
            Exit Function
        End If
    Next i

    Dim delCnt As Long, delFail As Long
    Dim linkDelCnt As Long, linkDelFail As Long
    Dim cleanFail As Long
    Dim missCnt As Long

    ' 1) 删除文件/文件夹（missCnt 在此自行统计，不信任调用方预计算）
    For i = 0 To UBound(pathArr)
        p = pathArr(i)
        Err.Clear
        On Error Resume Next
        If fso.FileExists(p) Then
            Err.Clear
            fso.DeleteFile p, True
            If Err.Number = 0 Then delCnt = delCnt + 1 Else delFail = delFail + 1: Err.Clear
        ElseIf fso.FolderExists(p) Then
            Err.Clear
            fso.DeleteFolder p, True
            If Err.Number = 0 Then delCnt = delCnt + 1 Else delFail = delFail + 1: Err.Clear
        Else
            ' 断链：目标本就不存在，不算文件删除失败
            missCnt = missCnt + 1
        End If
        Err.Clear
        On Error GoTo 0
    Next i

    ' 2) 删链接 + 清内容（保留单元格边框/填充格式；清除超链接字体样式和文本）
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

    ' 3) 清基线：文件删自身键；文件夹连带删除 "文件夹\..." 全部后代键
    '    严格加载（oracle 复核 #3）：区分「基线文件不存在」与「读取失败」——
    '    读取失败必须如实报 baseReadFail，且绝不把部分字典回写（防覆盖丢键）；
    '    写回失败单独报 baseWriteFail（写回失败不保证文件未改动，措辞区分）
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
            For Each p In pathArr
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
    If delFail > 0 Or linkDelFail > 0 Or cleanFail > 0 Or baseReadFail Or baseWriteFail Then
        msg = "附件删除结束（部分失败）："
    Else
        msg = "附件删除完成："
    End If
    msg = msg & "已删除 " & delCnt & " 个文件/文件夹、" & linkDelCnt & " 个链接"
    If delFail > 0 Then msg = msg & "，" & delFail & " 个文件/文件夹删除失败"
    If linkDelFail > 0 Then msg = msg & "，" & linkDelFail & " 个链接删除失败"
    If cleanFail > 0 Then msg = msg & "，" & cleanFail & " 个单元格内容/样式清理失败"
    If baseReadFail Then msg = msg & "，基线读取失败（文件未改动，未执行基线清理，请人工核对）"
    If baseWriteFail Then msg = msg & "，基线清理结果保存失败（请人工核对）"
    If missCnt > 0 Then msg = msg & "，" & missCnt & " 个链接目标本就不存在"
    AttachLink_DeleteAttachmentCore = msg
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
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要放置底稿引用的单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

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

    ' 第三步：复制目标底稿到附件夹（自包含原则）
    Dim attachDir As String
    attachDir = EnsureAttachDirPrompted(fso, wb)
    If Len(attachDir) = 0 Then Exit Sub
    Dim destFile As String, dispName As String
    destFile = attachDir & "\" & fso.GetFileName(targetPath)
    dispName = fso.GetBaseName(targetPath) & "-" & targetSheet
    ' 同名不同内容检测（oracle 复核）：附件夹已有同名副本时，内容一致才静默复用；
    ' 不一致必须提示——防旧版或异源同名底稿被静默引用（自包含原则被破坏）
    Dim needCopy As Boolean
    If fso.FileExists(destFile) Then
        Dim hSrc As String, hDst As String
        hSrc = Sha256File(targetPath)
        hDst = Sha256File(destFile)
        If Len(hSrc) > 0 And Len(hDst) > 0 And StrComp(hSrc, hDst, vbBinaryCompare) = 0 Then
            needCopy = False
        Else
            If MsgBox("附件夹内已存在同名但内容不同的底稿副本：" & vbCrLf & _
                      fso.GetFileName(destFile) & vbCrLf & vbCrLf & _
                      "「是」= 用所选底稿覆盖副本（已有指向它的引用将看到新内容）；" & vbCrLf & _
                      "「否」= 取消本次插入。", vbYesNo + vbExclamation, "AttachLink 底稿引用") <> vbYes Then Exit Sub
            needCopy = True
        End If
    Else
        needCopy = True
    End If
    If needCopy Then fso.CopyFile targetPath, destFile, True

    ' 插入链接：Address=附件夹内相对路径，SubAddress=工作表!A1
    Dim cellList As Collection
    Set cellList = FlattenSelection(Selection)
    If cellList.Count = 0 Then
        MsgBox "选区内没有可放置链接的单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If
    Dim target As Object, relAddr As String, subAddr As String
    Set target = cellList(1)
    If target.Hyperlinks.Count > 0 Then target.Hyperlinks(1).Delete
    relAddr = ATTACH_FOLDER_NAME & "\" & fso.GetFileName(destFile)
    ' SubAddress 转义（oracle 复核）：统一单引号包裹（Excel 对简单表名同样合法），
    ' 表名内的单引号按 Excel 规则双写——空格/引号/括号/数字开头等全部安全
    subAddr = "'" & Replace(targetSheet, "'", "''") & "'!A1"
    target.Hyperlinks.Add Anchor:=target, Address:=relAddr, SubAddress:=subAddr, TextToDisplay:=dispName

    MsgBox "底稿引用已插入：" & vbCrLf & _
           "  显示名：" & dispName & vbCrLf & _
           "  底稿副本：" & relAddr & vbCrLf & _
           "  工作表：" & targetSheet & vbCrLf & _
           "点击链接即打开副本并跳转到该工作表。", vbInformation, "AttachLink"
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
Public Function AttachLink_ArchiveMigrateCore(ByVal srcDir As String) As String
    Dim wbHost As Object
    Set wbHost = ActiveWorkbook
    If wbHost Is Nothing Then
        AttachLink_ArchiveMigrateCore = "失败：没有活动工作簿（请在归档目录中打开任一底稿后运行）。"
        Exit Function
    End If
    If Len(wbHost.Path) = 0 Then
        AttachLink_ArchiveMigrateCore = "失败：活动工作簿尚未保存。"
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcDir) Then
        AttachLink_ArchiveMigrateCore = "失败：源文件夹不存在：" & srcDir
        Exit Function
    End If

    Dim archiveRoot As String, dstDir As String
    archiveRoot = wbHost.Path
    dstDir = archiveRoot & "\" & ATTACH_FOLDER_NAME
    If Not EnsureFolder(fso, dstDir) Then
        AttachLink_ArchiveMigrateCore = "失败：无法创建目标文件夹：" & dstDir
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

    ' 1) 整体移动（目标已有同名文件以账套原件覆盖；单文件失败记入清单不中断）
    movedCnt = MoveTreeInto(fso, srcDir, dstDir, errs)

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
    s = s & vbCrLf & vbCrLf & "建议对归档底稿运行「附件清单 / 体检」复核。"
    AttachLink_ArchiveMigrateCore = s
    Exit Function
Fail:
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    AttachLink_ArchiveMigrateCore = "归档迁移失败（未完成）：" & Err.Description & vbCrLf & _
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
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME
    fullPath = ResolveLinkPath(fso, wb.Path, target.Address)
    fName = fso.GetFileName(fullPath)
    If StrComp(Left$(fullPath, Len(attachDir)), attachDir, vbTextCompare) <> 0 Then
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

    ' 基线必须已登记（版本号与历史来自基线）
    Dim base As Object, key As Variant, meta() As String
    Set base = LoadBaseline(fso, attachDir)
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
              "旧版本文件将被【删除】（更替动作将记入体检清单与基线，可留痕追溯）。" & vbCrLf & vbCrLf & _
              "「是」执行更替　　「否」取消", vbYesNo + vbExclamation, "AttachLink 附件版本更新") <> vbYes Then Exit Sub

    ' 文件操作全部成功后才写基线（避免中途失败造成状态不一致）
    Dim tCreate As String
    tCreate = meta(2)
    If Len(tCreate) = 0 Then tCreate = Format$(Now, "yyyy-mm-dd hh:nn")
    fso.DeleteFile fullPath, True
    fso.CopyFile newPath, fullPath, False
    base(key) = hNew & "|" & (CLng(meta(1)) + 1) & "|" & tCreate & "|" & Format$(Now, "yyyy-mm-dd hh:nn")
    SaveBaseline fso, attachDir, base

    MsgBox "附件版本更新完成：" & fName & " 已更替为 v" & (CLng(meta(1)) + 1) & "。" & vbCrLf & _
           "链接地址不变（同名替换）。建议重新运行「附件清单 / 体检」核对。", vbInformation, "AttachLink"
    Exit Sub
Fail:
    MsgBox "附件版本更新时出错（文件可能被占用，请关闭相关程序后重试；基线未被修改）：" & Err.Description, vbCritical, "AttachLink"
End Sub

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
' SHA-256 经系统 certutil 实现（WPS 兼容决策，2026-11-21）：WPS 的 VBA 环境缺少
' LongPtr/PtrSafe 生态，CAPI Declare 触发「用户定义类型未定义」编译错误且按需
' 编译会连带击穿调用链（清单/插入文件夹）。certutil 为 Windows 内置命令
' （Win7+），经 WScript.Shell 隐藏窗口执行、输出落临时文件解析，无任何宿主
' 类型依赖。代价：单文件约 50ms 进程开销（体检几十个附件数秒级，可接受；
' 后续优化见开发交接 T14 哈希缓存）。
Private Function Sha256File(ByVal filePath As String) As String
    ' 宿主分派（2026-11-21）：WPS 宏安全拦截子进程创建（0x80070005）→ 纯 VBA；
    ' Excel 走 certutil（快）。互为回退：任一路径失败自动尝试另一条（oracle：宿主误判/环境异常时兜底）
    Dim h As String
    If IsWpsHost() Then
        h = Sha256VBA(filePath)
        If Len(h) = 0 Then h = Sha256Certutil(filePath)
    Else
        h = Sha256Certutil(filePath)
        If Len(h) = 0 Then h = Sha256VBA(filePath)
    End If
    Sha256File = h
End Function

' 宿主判定：WPS 安装路径含 kingsoft/wps（Application.Name 在 WPS 伪装为 Excel 不可靠）
Private Function IsWpsHost() As Boolean
    On Error Resume Next
    Dim p As String
    p = Application.Path
    IsWpsHost = (InStr(1, p, "kingsoft", vbTextCompare) > 0) Or (InStr(1, p, "wps", vbTextCompare) > 0)
End Function

Private Function Sha256Certutil(ByVal filePath As String) As String
    On Error GoTo Fail
    Sha256Certutil = ""
    Dim fso As Object, tmp As String, sh As Object, cmd As String, rc As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
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
Private Function LoadBaseline(ByVal fso As Object, ByVal attachDir As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo Done
    Dim p As String
    p = attachDir & "\" & BASELINE_FILE
    If Not fso.FileExists(p) Then GoTo Done
    Dim stm As Object, text As String, lines() As String, i As Long
    Dim line As String, parts() As String, verS As String, tC As String, tU As String
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2                                   ' adTypeText
        .Charset = "UTF-8"
        .Open
        .LoadFromFile p
        .Position = 0
        text = .ReadText
        .Close
    End With
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        line = Replace(lines(i), vbCr, "")
        If Len(line) > 0 Then
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                verS = "1": tC = "": tU = ""
                If UBound(parts) >= 4 Then
                    verS = parts(2): tC = parts(3): tU = parts(4)
                End If
                If Len(parts(0)) > 0 And Len(parts(1)) > 0 And Len(verS) > 0 Then
                    d(LCase$(parts(0))) = parts(1) & "|" & verS & "|" & tC & "|" & tU
                End If
            End If
        End If
    Next i
Done:
    Set LoadBaseline = d
End Function

' 严格版基线加载（删除等破坏性路径专用，oracle 复核 #3）：
' status 返回 0=正常（可能为空字典）/ 1=基线文件不存在 / 2=读取或解析失败。
' 与 LoadBaseline 的区别：任何异常（含文件被独占锁定、流读取失败）都如实
' 上报 status=2，绝不把「读取失败」伪装成「空基线=无需清理」。
Private Function LoadBaselineStrict(ByVal fso As Object, ByVal attachDir As String, ByRef status As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    status = 0
    Dim p As String
    p = attachDir & "\" & BASELINE_FILE
    If Not fso.FileExists(p) Then
        status = 1
        Set LoadBaselineStrict = d
        Exit Function
    End If
    On Error GoTo BadRead
    Dim stm As Object, text As String, lines() As String, i As Long
    Dim line As String, parts() As String, verS As String, tC As String, tU As String
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile p
        .Position = 0
        text = .ReadText
        .Close
    End With
    lines = Split(text, vbLf)
    ' 严格逐行校验（oracle 二轮 #3）：任何非空行无法按受支持格式完整解析
    ' （无分隔符 / 字段数非 v1(2) 或 v2(5) / 空键空哈希空版本 / 归一化重复键）
    ' 均判读取失败——绝不让部分字典伪装成完整基线回写覆盖
    Dim keyN As String
    For i = 0 To UBound(lines)
        line = Replace(lines(i), vbCr, "")
        If Len(line) > 0 Then
            parts = Split(line, "|")
            If UBound(parts) <> 1 And UBound(parts) <> 4 Then GoTo BadRead
            verS = "1": tC = "": tU = ""
            If UBound(parts) = 4 Then
                verS = parts(2): tC = parts(3): tU = parts(4)
            End If
            If Len(parts(0)) = 0 Or Len(parts(1)) = 0 Or Len(verS) = 0 Then GoTo BadRead
            keyN = LCase$(parts(0))
            If d.Exists(keyN) Then GoTo BadRead
            d(keyN) = parts(1) & "|" & verS & "|" & tC & "|" & tU
        End If
    Next i
    Set LoadBaselineStrict = d
    Exit Function
BadRead:
    status = 2
    Set d = CreateObject("Scripting.Dictionary")
    Set LoadBaselineStrict = d
End Function

' 基线落盘（UTF-8 覆盖写）。返回 True=成功；False=失败（调用方不得报告登记完成——
' oracle #3：静默吞错会让未持久化登记被当作已完成）
Private Function SaveBaseline(ByVal fso As Object, ByVal attachDir As String, ByVal d As Object) As Boolean
    On Error GoTo Fail
    If Not EnsureFolder(fso, attachDir) Then Exit Function
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
    stm.SaveToFile attachDir & "\" & BASELINE_FILE, 2             ' adSaveCreateOverWrite
    stm.Close
    SaveBaseline = True
    Exit Function
Fail:
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
    For Each f In fso.GetFolder(folderPath).Files
        fileCnt = fileCnt + 1
        key = LCase$(Mid$(f.Path, Len(attachDir) + 2))
        h = Sha256File(f.Path)
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
    If Err.Number <> 0 Then failCnt = failCnt + 1: Err.Clear
    For Each sf In fso.GetFolder(folderPath).SubFolders
        CheckFolderBaseline fso, base, sf.Path, attachDir, regCnt, fileCnt, badCnt, failCnt, dirty
    Next sf
    If Err.Number <> 0 Then failCnt = failCnt + 1: Err.Clear
End Sub

' 前置检查（插入文件/文件夹共用）：附件文件夹不存在则自动创建；已存在则
' 提示核对（NOTICE_ONCE_PER_FOLDER 语义）。返回附件夹路径；用户取消或
' 创建失败返回 ""。
Private Function EnsureAttachDirPrompted(ByVal fso As Object, ByVal wb As Object) As String
    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME
    If fso.FolderExists(attachDir) Then
        If (Not NOTICE_ONCE_PER_FOLDER) Or StrComp(lastNotifiedAttachDir, attachDir, vbTextCompare) <> 0 Then
            If MsgBox("当前底稿同级已存在「" & ATTACH_FOLDER_NAME & "」文件夹：" & vbCrLf & _
                      attachDir & vbCrLf & vbCrLf & _
                      "本次插入的附件将归档复制到该文件夹，请先核对是否为目标目录。" & vbCrLf & vbCrLf & _
                      "「确定」继续选择　　「取消」退出本次插入", _
                      vbOKCancel + vbExclamation, "AttachLink") <> vbOK Then Exit Function
            lastNotifiedAttachDir = attachDir
        End If
    Else
        If Not EnsureFolder(fso, attachDir) Then
            MsgBox "无法创建附件文件夹：" & attachDir, vbCritical, "AttachLink"
            Exit Function
        End If
        lastNotifiedAttachDir = attachDir
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

' 判断是否为本加载项插入的附件链接：
' 相对地址以「附件\」开头，或绝对路径的父目录名为附件文件夹，
' 或 ..\ 形态相对地址（鼎信诺归档另存会把链接重基为 ..\账套\...\附件\x，WORKLOG #14）
' fsoIn / wbPathIn：可选传入调用方已建的 FSO 实例 / 底稿路径（用于解析 ..\ 相对形态）
Private Function IsAttachmentLink(ByVal hl As Object, Optional ByVal fsoIn As Object, Optional ByVal wbPathIn As String = "") As Boolean
    Dim a As String
    a = hl.Address
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
    If InStr(addr, ":") > 0 Or Left$(addr, 2) = "\\" Then
        ResolveLinkPath = addr
    Else
        ResolveLinkPath = wbPath & "\" & addr
    End If
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

' 归档复制：返回最终落盘路径；重名时弹窗确认（覆盖 / 自动重命名 / 跳过）
' 返回 "" 表示用户选择跳过
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
    localAttach = wb.Path & "\" & ATTACH_FOLDER_NAME
    For Each ws In wb.Worksheets
        For Each hl In ws.Hyperlinks
            a = hl.Address
            If Len(a) > 0 Then
                If InStr(1, a, ATTACH_FOLDER_NAME, vbTextCompare) > 0 Then
                    full = ResolveLinkPath(fso, wb.Path, a)
                    parentDir = fso.GetParentFolderName(full)
                    If Len(parentDir) > 0 Then
                        If StrComp(fso.GetFileName(parentDir), ATTACH_FOLDER_NAME, vbTextCompare) = 0 Then
                            If StrComp(parentDir, localAttach, vbTextCompare) <> 0 Then
                                hl.Address = ATTACH_FOLDER_NAME & "\" & fso.GetFileName(a)
                                n = n + 1
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
' 返回成功移动的文件数；单文件失败记入 errs 不中断。
Private Function MoveTreeInto(ByVal fso As Object, ByVal srcDir As String, ByVal dstDir As String, ByVal errs As Collection) As Long
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
            Err.Clear
        Else
            n = n + 1
        End If
        On Error GoTo Fail
    Next f
    For Each sf In fso.GetFolder(srcDir).SubFolders
        n = n + MoveTreeInto(fso, sf.Path, dstDir & "\" & sf.Name, errs)
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
