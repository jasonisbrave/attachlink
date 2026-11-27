' 由构建脚本写入加载项的 ThisWorkbook 模块（请勿放入标准模块）
' v1.6.4：本模块不再有 Workbook_Open，也不再处理「打开工作簿」事件。
'   - 外部装载（看门狗经 Workbooks.Open 打开本加载项）：完全靠看门狗显式调用 AttachLink_Initialize
'   - 交互式启动（双击打开 Excel、在加载项列表里勾选）：靠标准模块的 Auto_Open
'   两条路都经 AttachLink_BindEvents 挂上这里的应用事件（右键自愈、保存前归一化）。
'   为什么要去掉「打开」相关的两个入口：它们的处理一执行，Excel 就会把宿主的 DisplayAlerts /
'   ScreenUpdating 复位为 True（2026-11-26 / 27 实测，与加载项代码无关）——别的程序拉起的 Excel 里，
'   对方关掉的提示会被重新打开，隐藏的自动化实例会卡在提示框上。
Option Explicit

Private WithEvents app As Application

' 挂应用事件：由 AttachLink_Initialize 调用，幂等。事件对象声明在本文档模块，绑定只能在这里做。
' 这里不吞错误：绑定失败要让初始化宏报错，看门狗的「三项齐备」判定才会发现
Public Sub AttachLink_BindEvents()
    Set app = Application
End Sub

Private Sub app_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' 右键瞬间自愈：右键菜单被其他加载项清除/重置时，当场补回
    On Error Resume Next
    AttachLink_EnsureMenu
    ' T8：该工作簿上次保存若有归一化失败，在这里提醒。v1.6.4 起不再处理「打开工作簿」事件，
    ' 原先「打开时再提醒」改在用户于该工作簿里右键时出现（保存后的那次状态栏提醒不变）
    AttachLink_NotifyNormalizeFailure Sh.Parent
End Sub

Private Sub app_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' 保存前把被 Excel 改写为绝对路径的附件链接归一化回相对路径
    On Error Resume Next
    Dim normFailCnt As Long
    normFailCnt = NormalizeAttachmentLinks(Wb)
    ' T8/N-02：仅当归一化确有失败时才调度 OnTime 提醒（否则每次保存都空跑提醒过程，
    ' 且死记录会让提醒永远驻留）；NormalizeAttachmentLinks 内部已清理已关闭工作簿的记录
    If normFailCnt > 0 Then Application.OnTime Now, "AttachLink_NotifyNormalizeFailure"
End Sub
