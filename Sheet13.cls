VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim LastLine As String
    Dim TargetValue As String
    Dim TargetRow As String
    Dim TargetColumn As String
    Dim Mode As String
    TargetRow = CStr(Target.Row)
    TargetColumn = CStr(Target.Column)
    Application.EnableEvents = False
    If TargetRow = 3 And TargetColumn = 17 Then
        Mode = "认证状态"
        If (Target.Value <> "") Then
            TargetValue = CStr(Target.Value)
        End If
    ElseIf TargetRow = 3 And TargetColumn = 2 Then
        Mode = "课程序号"
        If (Target.Value <> "") Then
            TargetValue = CStr(Target.Value)
        End If
    ElseIf TargetRow = 7 And TargetColumn = 2 Then
        Mode = "认证专业"
        If (Target.Value <> "") Then
            TargetValue = CStr(Target.Value)
        End If
    Else
        Mode = "无操作"
    End If
    If Mode <> "无操作" Then
        Call 课程目标和综合分析(Mode, TargetValue, TargetRow, TargetColumn)
    End If
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Application.Intersect(Target, Range("A1:AG200")) Is Nothing Then
        Application.CellDragAndDrop = False
    Else
        Application.CellDragAndDrop = True
    End If
End Sub
Private Sub Worksheet_Deactivate()
    Application.CellDragAndDrop = True
End Sub
'修订日期：2019年5月6日
