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
' 构建与安装：见 build\build.ps1 与 build\install.ps1
'=====================================================================
Option Explicit

'============================ 配置区 =================================
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
' 1.2.1：风险处置（Issue #1）——看门狗父进程白名单（配置化默认开放）+ 装载版本日志
' 1.2.0：文件夹引用（复制入附件夹+夹内基线）/ 重命名链接显示名 / 插入对话框记忆上次源目录 / 右键菜单二级化
' ⚠ 发布流程约定：版本发布（GitHub Release + git push）必须先经用户确认
Public Const ATTACHLINK_VERSION As String = "1.2.1"
'=====================================================================

Private Const MENU_TAG As String = "AttachLinkCtx"
' 内容基线文件名（位于附件文件夹内，随附件夹整体迁移/同步；UTF-8，格式：相对子路径|SHA-256）
Private Const BASELINE_FILE As String = "_attachlink_baseline.csv"

' ⚠ VBA 位置规则：模块级 Declare 必须放在【第一条过程之前】（本声明区）。
' 曾放在模块中段导致全量编译失败——SaveAs(xlam) 触发全项目编译时报
' 0x800A03EC 或产物宏不可用，而按需编译的 Run 却正常，极具迷惑性
' （见 WORKLOG #15）。新增任何 Declare/模块级变量一律放这里。
' SHA-256 经 WinCrypt CAPI（advapi32）：系统原生组件，无 COM/.NET 依赖
' （.NET COM 路线已实测排除：SHA256Managed 在 .NET 4.x 无 ComputeHash_2、
'  SHA256CryptoServiceProvider 未注册）。
#If VBA7 Then
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef phProv As LongPtr, ByVal szContainer As LongPtr, ByVal szProvider As LongPtr, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal hProv As LongPtr, ByVal algId As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hHash As LongPtr, ByVal pbData As LongPtr, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hHash As LongPtr, ByVal dwParam As Long, ByVal pbData As LongPtr, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hHash As LongPtr) As Long
#Else
Private Declare Function CryptAcquireContextW Lib "advapi32" (ByRef phProv As Long, ByVal szContainer As Long, ByVal szProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal algId As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, ByVal pbData As Long, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As Long, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
#End If

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
    Dim cb As CommandBar, c As CommandBarControl, found As Boolean
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
    Dim cb As CommandBar, ctl As CommandBarControl, btn As CommandBarButton
    On Error Resume Next
    Set cb = Application.CommandBars("Cell")
    If cb Is Nothing Then Exit Sub
    For Each ctl In cb.Controls
        If ctl.Tag = MENU_TAG Then ctl.Delete
    Next ctl
    If adding Then
        Set btn = cb.Controls.Add(msoControlButton, , , 1, True)
        btn.Caption = "插入附件链接（可多选）"
        btn.OnAction = "AttachLink_Insert"
        btn.Tag = MENU_TAG
        btn.BeginGroup = True

        Set btn = cb.Controls.Add(msoControlButton, , , 2, True)
        btn.Caption = "插入文件夹引用"
        btn.OnAction = "AttachLink_InsertFolder"
        btn.Tag = MENU_TAG

        Set btn = cb.Controls.Add(msoControlButton, , , 3, True)
        btn.Caption = "重命名链接显示名"
        btn.OnAction = "AttachLink_RenameLink"
        btn.Tag = MENU_TAG

        ' 二级菜单：附件清单（体检 / 断链修复）
        Dim pop As CommandBarPopup, subBtn As CommandBarButton
        Set pop = cb.Controls.Add(msoControlPopup, , , 4, True)
        pop.Caption = "附件清单"
        pop.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(msoControlButton, , , 1, True)
        subBtn.Caption = "附件清单 / 体检"
        subBtn.OnAction = "AttachLink_BuildIndex"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(msoControlButton, , , 2, True)
        subBtn.Caption = "附件断链修复"
        subBtn.OnAction = "AttachLink_Repair"
        subBtn.Tag = MENU_TAG

        ' 二级菜单：归档（迁移 / 版本更新）
        Set pop = cb.Controls.Add(msoControlPopup, , , 5, True)
        pop.Caption = "归档"
        pop.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(msoControlButton, , , 1, True)
        subBtn.Caption = "归档迁移（批量）"
        subBtn.OnAction = "AttachLink_ArchiveMigrate"
        subBtn.Tag = MENU_TAG
        Set subBtn = pop.Controls.Add(msoControlButton, , , 2, True)
        subBtn.Caption = "附件版本更新"
        subBtn.OnAction = "AttachLink_UpdateVersion"
        subBtn.Tag = MENU_TAG
    End If
End Sub

'============================ 1. 插入附件链接 =========================
Public Sub AttachLink_Insert()
    On Error GoTo Fail
    Dim wb As Workbook
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

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
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

    Dim inserted As Long, skipped As Long, idx As Long
    Dim i As Long
    Dim destPath As String, targetCell As Range

    For i = 1 To fd.SelectedItems.Count
        destPath = ArchiveFile(fso, fd.SelectedItems(i), attachDir)
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
    Next i

    Dim msg As String
    msg = "已插入 " & inserted & " 个附件链接，文件已归档到同级「" & ATTACH_FOLDER_NAME & "」文件夹。"
    If skipped > 0 Then msg = msg & vbCrLf & "跳过 " & skipped & " 个文件（已取消或可用单元格不足）。"
    MsgBox msg, vbInformation, "AttachLink"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "插入附件链接时出错：" & Err.Description, vbCritical, "AttachLink"
End Sub

'======================== 2. 附件清单 / 体检 ==========================
' 状态列（E）：正常 / 断链（文件不存在）
' 内容基线列（F）：一致 / ⚠ 内容变更 / 已登记基线 / 校验不可用
'   内容基线 = 附件夹内 _attachlink_baseline.csv（文件名|SHA-256，UTF-8），
'   随附件夹一起归档迁移/网盘同步。哈希不一致（附件被替换、篡改或同步端
'   同名覆盖）持续红字警示且【不自动更新基线】；确为有意替换时，删除该
'   基线文件后重新生成清单即可重建。
Public Sub AttachLink_BuildIndex()
    On Error GoTo Fail
    Dim wb As Workbook
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
Public Function AttachLink_BuildIndexCore(ByVal wb As Workbook) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME

    ' 内容基线（SHA-256 经 WinCrypt；单文件失败显示「校验不可用」，不影响清单）
    Dim base As Object
    Set base = LoadBaseline(fso, attachDir)
    Dim baseDirty As Boolean

    Application.ScreenUpdating = False
    Dim wsIdx As Worksheet
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

    Dim ws As Worksheet, hl As Hyperlink
    Dim r As Long, total As Long, brokenCnt As Long, changedCnt As Long, unavailCnt As Long
    Dim cnt As Long, g As Long
    Dim linkPath As String, ok As Boolean
    Dim h As String, key As String, cellF As String, cellG As String, meta() As String

    r = 4
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, INDEX_SHEET_NAME, vbTextCompare) <> 0 Then
            cnt = 0
            For Each hl In ws.Hyperlinks
                If IsAttachmentLink(hl, fso) Then cnt = cnt + 1
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
                    If IsAttachmentLink(hl, fso) Then
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
        If r > 4 Then .Range("A3:G" & (r - 1)).Borders.LineStyle = xlContinuous
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
        msg = msg & vbCrLf & "⚠ 基线文件写入失败：本次新登记未持久化（" & BASELINE_FILE & "），下次体检将重新登记，请检查附件夹写权限。"
    End If
    AttachLink_BuildIndexCore = msg
End Function

'============================ 3. 断链修复 =============================
Public Sub AttachLink_Repair()
    On Error GoTo Fail
    Dim wb As Workbook
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
    Dim ws As Worksheet, hl As Hyperlink, rp As String
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, INDEX_SHEET_NAME, vbTextCompare) <> 0 Then
            For Each hl In ws.Hyperlinks
                If IsAttachmentLink(hl, fso) Then
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
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
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

    Dim cnt As Long, area As Range, c As Range
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
Public Sub NormalizeAttachmentLinks(ByVal wb As Workbook)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    If Len(wb.Path) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attachDir As String
    attachDir = wb.Path & "\" & ATTACH_FOLDER_NAME
    Dim ws As Worksheet, hl As Hyperlink, parentDir As String
    For Each ws In wb.Worksheets
        For Each hl In ws.Hyperlinks
            If IsAttachmentLink(hl, fso) Then
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
    Dim wb As Workbook
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

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
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
    Dim wbHost As Workbook
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
    If failCnt > 0 Then tail = vbCrLf & "⚠ 有 " & failCnt & " 个文件登记失败（不可读/被占用），体检将显示「校验不可用」"
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
Public Sub AttachLink_RenameLink()
    On Error GoTo Fail
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要重命名的附件链接单元格。", vbInformation, "AttachLink"
        Exit Sub
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim c As Range, hl As Hyperlink, target As Hyperlink
    Set target = Nothing
    For Each c In Selection.Cells
        For Each hl In c.Hyperlinks
            If IsAttachmentLink(hl, fso) Then Set target = hl: Exit For
        Next hl
        If Not target Is Nothing Then Exit For
    Next c
    If target Is Nothing Then
        MsgBox "选区内未发现附件链接。", vbInformation, "AttachLink"
        Exit Sub
    End If

    Dim newName As String
    newName = InputBox("请输入新的显示名（仅改变链接显示文本，不改动文件本身）：", _
                       "AttachLink 重命名链接显示名", target.TextToDisplay)
    newName = Trim$(newName)
    If Len(newName) = 0 Then Exit Sub
    target.TextToDisplay = newName
    Exit Sub
Fail:
    MsgBox "重命名时出错：" & Err.Description, vbCritical, "AttachLink"
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
    Dim wb As Workbook
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

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
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
    Dim wbHost As Workbook
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
    Dim f As Object, wb As Workbook, n As Long, opened As Boolean
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
    Dim wb As Workbook
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

    ' 取选区中第一个附件链接
    Dim c As Range, hl As Hyperlink, target As Hyperlink
    Set target = Nothing
    For Each c In Selection.Cells
        For Each hl In c.Hyperlinks
            If IsAttachmentLink(hl, fso) Then Set target = hl: Exit For
        Next hl
        If Not target Is Nothing Then Exit For
    Next c
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
    Dim base As Object, key As String, meta() As String
    Set base = LoadBaseline(fso, attachDir)
    key = LCase$(Mid$(fullPath, Len(attachDir) + 2))
    If Not base.Exists(key) Then
        MsgBox "该附件尚未登记内容基线。请先运行「附件清单 / 体检」登记基线，再进行版本更新。", vbExclamation, "AttachLink"
        Exit Sub
    End If
    meta = Split(base(key), "|")
    ' meta: 0=哈希 1=版本 2=登记时间 3=更新时间

    ' 选择新文件
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
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
' 声明见模块顶部声明区（WinCrypt Declare 必须位于第一条过程之前，位置规则
' 见顶部注释与 WORKLOG #15）

' 计算文件 SHA-256（十六进制小写64位）；失败返回 ""（调用方显示「校验不可用」）
Private Function Sha256File(ByVal filePath As String) As String
    Dim hProv As LongPtr, hHash As LongPtr, stm As Object
    Dim bytes As Variant, arr() As Byte, n As Long, hashBytes() As Byte, hashLen As Long
    Dim i As Long, s As String
    Const PROV_RSA_AES As Long = 24                ' SHA-256 需要 AES CSP；PROV_RSA_FULL 会报 NTE_BAD_ALG_ID(0x80090008)
    Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
    Const CALG_SHA_256 As Long = &H800C&           ' 注意：尾部 & 必不可少。VBA 把 >= &H8000 的字面量解析为
                                                    ' 负 Integer，符号扩展后变 0xFFFF800C，导致 NTE_BAD_ALG_ID（实测踩坑）
    Const HP_HASHVAL As Long = 2
    Const CHUNK As Long = 65536
    On Error GoTo Fail
    Sha256File = ""
    If CryptAcquireContextW(hProv, 0, 0, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then GoTo RelAll
    If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) = 0 Then GoTo RelProv
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 1                                   ' adTypeBinary
        .Open
        .LoadFromFile filePath
        .Position = 0
        Do While Not .EOS
            bytes = .Read(CHUNK)
            arr = bytes                             ' Variant(Byte()) → 类型化数组：VarPtr 必须作用于类型化数组元素
            n = UBound(arr) - LBound(arr) + 1
            If CryptHashData(hHash, VarPtr(arr(LBound(arr))), n, 0) = 0 Then GoTo RelAll
        Loop
        .Close
    End With
    hashLen = 32
    ReDim hashBytes(0 To 31)
    If CryptGetHashParam(hHash, HP_HASHVAL, VarPtr(hashBytes(0)), hashLen, 0) = 0 Then GoTo RelAll
    For i = 0 To 31
        s = s & Right$("0" & Hex$(hashBytes(i)), 2)
    Next i
    Sha256File = LCase$(s)
RelAll:
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    On Error GoTo Fail
    If hHash <> 0 Then CryptDestroyHash hHash
RelProv:
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    Exit Function
Fail:
    Resume RelAll
End Function

' 读取内容基线为字典（键=附件夹内相对子路径小写，值=「哈希|版本|登记时间|更新时间」）。
' 兼容旧 v1 格式（仅「路径|哈希」两字段）自动按版本 1 迁移。
' 基线不存在/损坏时返回空字典（不中断清单）
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
Private Function EnsureAttachDirPrompted(ByVal fso As Object, ByVal wb As Workbook) As String
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
' 相对地址以「附件\」开头，或绝对路径的父目录名为附件文件夹
' fsoIn：可选传入调用方已建的 FSO 实例（保存归一化等高频路径下避免反复 CreateObject）
Private Function IsAttachmentLink(ByVal hl As Hyperlink, Optional ByVal fsoIn As Object) As Boolean
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
Private Function FlattenSelection(ByVal sel As Range) As Collection
    Dim out As New Collection, area As Range, c As Range
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
    Dim lastCell As Range, r As Long, cand As Range
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

Private Sub InsertLink(ByVal targetCell As Range, ByVal relAddr As String, ByVal disp As String)
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
' 同名表存在但缺少本工具归属标记（CustomProperties）时返回 Nothing，
' 防止整表重建误清用户或鼎信诺生成的同名表。
Private Function GetOrCreateIndexSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, cp As CustomProperty, marked As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(INDEX_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = INDEX_SHEET_NAME
        ws.CustomProperties.Add "AttachLinkIndex", "1"
    Else
        For Each cp In ws.CustomProperties
            If StrComp(cp.Name, "AttachLinkIndex", vbTextCompare) = 0 Then marked = True
        Next cp
        If Not marked Then
            Exit Function
        End If
    End If
    Set GetOrCreateIndexSheet = ws
End Function

' 归档迁移：把工作簿中所有「指向别处」的附件链接改写为相对地址（附件\文件名）。
' 关键实证：鼎信诺另存归档时，Excel 会把链接重基为 ../ 相对路径指回账套原件
' （也可能保持绝对路径），两种形态都要处理——凡解析后不落在本底稿同级附件夹
' 内的附件链接，一律按文件名改写为「附件\文件名」（与本工具「附件平铺、
' 一格一链」的写入形态一致）。
' 返回改写数量；遇到受保护工作表等错误时抛出，由调用方记入失败清单。
Private Function RewriteAbsoluteAttachLinks(ByVal wb As Workbook, ByVal fso As Object) As Long
    Dim ws As Worksheet, hl As Hyperlink, a As String, full As String
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
Private Function SuggestAttachSource(ByVal wb As Workbook, ByVal fso As Object) As String
    On Error Resume Next
    Dim ws As Worksheet, hl As Hyperlink, dirs As Object, p As String
    Set dirs = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        For Each hl In ws.Hyperlinks
            If IsAttachmentLink(hl, fso) Then
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
Private Function WorkbookByPath(ByVal fullPath As String) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, fullPath, vbTextCompare) = 0 Then
            Set WorkbookByPath = w
            Exit Function
        End If
    Next w
End Function

' 清除单个单元格上的附件链接；清完无残留链接时恢复字体格式
Private Function ClearCellLinks(ByVal c As Range) As Long
    Dim i As Long, removed As Long
    For i = c.Hyperlinks.Count To 1 Step -1
        If IsAttachmentLink(c.Hyperlinks(i)) Then
            c.Hyperlinks(i).Delete
            removed = removed + 1
        End If
    Next i
    If removed > 0 Then
        If c.Hyperlinks.Count = 0 Then
            With c.Font
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlColorIndexAutomatic
            End With
        End If
    End If
    ClearCellLinks = removed
End Function
