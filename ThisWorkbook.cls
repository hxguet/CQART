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
    On Error Resume Next
    ThisWorkbook.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    Call �����¼�����
    Call Զ�̸��´���
    NoMsgBox = True
    If isUpdate Then
        Call �޶���ʽ
    End If
    NoMsgBox = False
    Call ��������
End Sub
'���޶����ڡ���2019��5��9��


