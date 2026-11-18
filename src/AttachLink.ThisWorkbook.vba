' 由构建脚本写入加载项的 ThisWorkbook 模块（请勿放入标准模块）
Option Explicit

Private WithEvents app As Application

Private Sub Workbook_Open()
    Set app = Application
    ' 无论 Excel 以何种方式启动（交互式 / 鼎信诺通过 COM 自动化拉起），
    ' 加载项装载时都会触发本事件；Auto_Open 在自动化启动下会被跳过，
    ' 因此初始化统一在这里兜底执行
    On Error Resume Next
    AttachLink_Initialize
End Sub

Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
    ' 打开任何工作簿时确认菜单在位（兜底自愈）
    On Error Resume Next
    AttachLink_EnsureMenu
End Sub

Private Sub app_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' 右键瞬间自愈：右键菜单被其他加载项清除/重置时，当场补回
    On Error Resume Next
    AttachLink_EnsureMenu
End Sub

Private Sub app_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' 保存前把被 Excel 改写为绝对路径的附件链接归一化回相对路径
    On Error Resume Next
    NormalizeAttachmentLinks Wb
End Sub
