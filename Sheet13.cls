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
        Mode = "��֤״̬"
        If (Target.Value <> "") Then
            TargetValue = CStr(Target.Value)
        End If
    ElseIf TargetRow = 3 And TargetColumn = 2 Then
        Mode = "�γ����"
        If (Target.Value <> "") Then
            TargetValue = CStr(Target.Value)
        End If
    ElseIf TargetRow = 7 And TargetColumn = 2 Then
        Mode = "��֤רҵ"
        If (Target.Value <> "") Then
            TargetValue = CStr(Target.Value)
        End If
    Else
        Mode = "�޲���"
    End If
    If Mode <> "�޲���" Then
        Call �γ�Ŀ����ۺϷ���(Mode, TargetValue, TargetRow, TargetColumn)
    End If
    Application.EnableEvents = True
End Sub
'�޶����ڣ�2019��5��28��