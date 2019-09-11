VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Row >= 2 And Target.Row <= 182 Then
        If Target.Column >= 16 And Target.Column <= 24 Then
            Range("L3:R182").Select
            With Selection
                Selection.NumberFormat = "General"
                .Value = .Value
            End With
        End If
    End If
End Sub
