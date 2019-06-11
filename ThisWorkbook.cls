VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim isUpdate As String
    On Error Resume Next
    ThisWorkbook.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    Call 允许事件触发
    Call 远程更新代码
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    isUpdate = Range("H8").Value
    Worksheets("专业矩阵状态").Visible = False
    NoMsgBox = True
    If isUpdate = "更新公式" Then
        Call 修订公式
    End If
    NoMsgBox = False
    Call 其他操作
End Sub
'【修订日期】：2019年5月9日

