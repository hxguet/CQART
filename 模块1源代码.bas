Attribute VB_Name = "ģ��1"
    Public Const MaxRecord As String = "185"
    Public Const MaxLineCout As String = "403"
    Public Const OldPassword As String = "dpt8"
    Public Const Password As String = "dpt8hx"
    Public Const MajorLastRow As String = "12"
    Public Const CSummary = 0
    Public Const CModuleCount = 0
    Public Const CSumLines = 1
    Public Const CStartLine = 0
    Public Const CModuleName = 1
    Public Const CFileName = 2
    Public Const CVersion = 3
    Public Const CRiviseDate = 4
    Public Const CUpdateInfo = 5
    Public Const CStatus As Integer = 0
    Public Const CMName As Integer = 1
    Public Const CBackup As Integer = 2
    Public Const CRelease As Integer = 3
    Public Const CCodeFileCount As Integer = 3
    Public School As String
    Public MajorCount As Integer
    Public MajorList(4) As String
    Public ModuleLastRivise() As String
    Public NoMsgBox As Boolean
    Public CodeFileName(0 To CCodeFileCount, 0 To 3) As String
    Public isOpenAfterPublish As Boolean
    Public Update As String
    

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim ThisSheet As String
    ThisSheet = ActiveSheet.Name
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H10").Value = "������Ϣ��"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
    Worksheets(ThisSheet).Activate
    Call �����¼�����
End Sub
    
''רҵ���޶�
Sub �޶���ʽ()
Dim TempWorkSheetVisible As Boolean
Dim SumRow As Integer
    Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("AC2:AI403").Select
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Range("T4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",RC[15],IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0)))"
    Selection.AutoFill Destination:=Range("T4:T403"), Type:=xlFillDefault
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "=RC[-15]"
    Range("AC3").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX('2-�γ�Ŀ����ۺϷ�������д��'!R3,MATCH('1-�Ծ�ɼ��ǼǱ���д��'!R2C,'2-�γ�Ŀ����ۺϷ�������д��'!R2,0))"
    Range("AC4").Select
    ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(RC[-15]),RC[-15]*R3C/100,0)"
    Range("AC2:AC4").Select
    Selection.AutoFill Destination:=Range("AC2:AH4"), Type:=xlFillDefault
    Range("AI4").Select
    ActiveCell.FormulaR1C1 = "=ROUND(SUM(RC[-6]:RC[-1]),0)"
    Range("AC4:AI4").Select
    Selection.AutoFill Destination:=Range("AC4:AI403"), Type:=xlFillDefault
    Range("AC4:AI403").Select
    Columns("AB:AI").Select
    Selection.ColumnWidth = 4
    Selection.EntireColumn.Hidden = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    TempWorkSheetVisible = Worksheets("4-�����������棨��д+��ӡ��").Visible
    Worksheets("4-�����������棨��д+��ӡ��").Visible = True
    Worksheets("4-�����������棨��д+��ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    
    Range("F6:G6").Select
    ActiveCell.FormulaR1C1 = _
        "=RC[-4]-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[14],""����"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[14],""ȡ��"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[14],""����"")"
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Range("P12").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[4],""ȡ��"")=R6C2,COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,""<60"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[4],""����"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[4],""����""),COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,""<60""))"
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Range("F16:H16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("I16:J16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("K16:L16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("M16:O16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("P15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""<60"")+COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")+COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""ȡ��"")"

    Range("P16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("H4:K5").Select
    Selection.NumberFormatLocal = "yyyy""��""m""��""d""��"";@"
    Range("A6:P16").Select
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Range("P10,F13:P13,F16:P16").Select
    Selection.NumberFormatLocal = "0.00_ "
    Range("D8:E9,J8:K9,P8:P10").Select
    Range("P8").Activate
    Selection.NumberFormatLocal = "0%"
    Range("P10").Select
    Selection.NumberFormatLocal = "0.00%"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("4-�����������棨��д+��ӡ��").Visible = TempWorkSheetVisible
    Call �޶�רҵ����״̬
    Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Rows("11:21").Select
    Selection.EntireRow.Hidden = False
    SumRow = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
    Rows(SumRow + 11 & ":20").Select
    Selection.EntireRow.Hidden = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub ��������()
    'ɾ��רҵ�������ఴť
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    'ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    'Dim sh As Shape
    'For Each sh In ActiveSheet.Shapes
    '    If sh.Name = "Drop Down 5606" Then
    '        sh.Delete
    '    End If
    'Next
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Call �����¼�����
    Call �޶���ʽ
    If Update = vbYes Then
        Worksheets("רҵ����״̬").Visible = True
        Worksheets("רҵ����״̬").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    
        If Range("H8").Value = "���¹�ʽ" Then
            Call �������ù�ʽ��ť
        End If
        Worksheets("רҵ����״̬").Activate
        Range("H8").Value = Range("H1").Value
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("רҵ����״̬").Visible = False
    End If
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    ActiveWorkbook.BreakLink Name:="E:\01-ѧ��\������������ģ�����Ϣ��ʾ\������������ģ��V5.xls", Type _
        :=xlExcelLinks
    ActiveSheet.Shapes.Range(Array("Button 4209")).Select
    Selection.OnAction = "�����ļ�"
    ActiveSheet.Shapes.Range(Array("Button 5")).Select
    Selection.OnAction = "��ӡ"
    ActiveSheet.Shapes.Range(Array("Button 4210")).Select
    Selection.OnAction = "CreateRecordWorkBook"
    ActiveSheet.Shapes.Range(Array("Button 4211")).Select
    Selection.OnAction = "�������ù�ʽ��ť"
    ActiveSheet.Shapes.Range(Array("Button 4212")).Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Selection.OnAction = "�����¼�����"
End Sub
Sub ���������()
    On Error Resume Next
    Dim temp As Boolean
    For Each sht In Sheets
        temp = Worksheets(sht.Name).Visible
        Worksheets(sht.Name).Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=OldPassword
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets(sht.Name).Visible = temp
    Next
End Sub
Sub ���ñ����ļ���Ϣ()
Dim BackupFilePath As String
Dim ReleaseFile As String
Dim RiviseDate As String
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    BackupFilePath = Range("H7").Value
    ReleaseFile = "ģ��1Դ����.bas"
    RiviseDate = Format(Now, "yyyy-mm-dd")
    CodeFileName(0, CStatus) = "����"
    CodeFileName(0, CMName) = "ģ��1"
    CodeFileName(0, CBackup) = BackupFilePath & "\ģ��1Դ����-" & Range("H1").Value & "-" & Format(RiviseDate, "YYYYMMDD") & ".bas"
    CodeFileName(0, CRelease) = ReleaseFile
    
    CodeFileName(1, CStatus) = "����"
    CodeFileName(1, CMName) = "Sheet13"
    CodeFileName(1, CBackup) = BackupFilePath & "\Sheet13-" & Range("H1").Value & "-" & Format(RiviseDate, "YYYYMMDD") & ".cls"
    CodeFileName(1, CRelease) = "Sheet13.cls"
    
    CodeFileName(2, CStatus) = "���"
    CodeFileName(2, CCMName) = "Sheet20"
    CodeFileName(2, CBackup) = BackupFilePath & "\Sheet20-" & Range("H1").Value & "-" & Format(RiviseDate, "YYYYMMDD") & ".cls"
    CodeFileName(2, CRelease) = "Sheet20.cls"
    
    CodeFileName(3, CStatus) = "���"
    CodeFileName(3, CMName) = "Sheet3"
    CodeFileName(3, CBackup) = BackupFilePath & "\Sheet3-" & Range("H1").Value & "-" & Format(RiviseDate, "YYYYMMDD") & ".cls"
    CodeFileName(3, CRelease) = "Sheet3.cls"
    
End Sub
Sub ���ɰ汾��()
    On Error Resume Next
    Dim MainVer As Integer
    Dim SubVer As Integer
    Dim RiviseVer As Integer
    Dim ModuleFileName As String
    Dim Version As String
    Dim RiviseDate As String
    Dim LineCount As Integer
    Dim LastLineCode$
    Dim Vbc As Object
    Dim ModuleName As String
    Dim ModuleCount As Integer
    Dim ReleaseFilePath As String
    Dim BackupFilePath As String
    
    Dim ReleaseFile As String
    Dim Commit As String
    Dim n As Integer
    Dim BatFile As String
    Dim TempStr As String
    Dim TestCode As String
    Dim ReadMeisEmpty As Boolean
    Call �޶�רҵ����״̬
    BatFile = "��������.bat"
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    MainVer = Range("H3").Value
    SubVer = Range("H4").Value
    RiviseVer = Range("H5").Value
    ReleaseFilePath = Range("H6").Value
    BackupFilePath = Range("H7").Value
    ReleaseFile = "ģ��1Դ����.bas"
    BatFile = "��������.bat"
    TestCode = Range("H9").Value
    If (TestCode = "�����汾") Then
        If (RiviseVer < 40) Then
            RiviseVer = RiviseVer + 1
        Else
            RiviseVer = 1
            Range("H5").Value = 1
            If (SubVer < 20) Then
                SubVer = SubVer + 1
            Else
                SubVer = 1
                Range("H4").Value = 1
                If (MainVer < 10) Then
                    MainVer = MainVer + 1
                 End If
           End If
        End If
    End If
    Version = "V" & MainVer & "." & Format(SubVer, "00") & "." & Format(RiviseVer, "00")
    RiviseDate = Format(Now, "yyyy-mm-dd")
    Commit = Format(Now, "yyyy-mm-dd hh:mm:ss") & "  Commit"
    ModuleCount = 0
    For Each Vbc In ThisWorkbook.VBProject.VBComponents
        If Vbc.Type = 1 And Mid(Vbc.Name, 1, 2) = "ģ��" Then
            ModuleCount = ModuleCount + 1
            If (Vbc.Name <> "ģ��1") Then
                Vbc.Name = "ģ��1"
            End If
        End If
    Next Vbc
    If ModuleCount = 1 Then
        Range("H1").Select
        ActiveCell.FormulaR1C1 = _
            "=""V""&R[2]C&"".""&TEXT(R[3]C,""00"")&"".""&TEXT(R[4]C,""00"")"
        Range("H2").Value = RiviseDate
        Range("H3").Value = MainVer
        Range("H4").Value = SubVer
        Range("H5").Value = RiviseVer
        Call ���ñ����ļ���Ϣ
        For i = 0 To CCodeFileCount
            If Dir(BackupFilePath & "\") = "" Then
                MkDir BackupFilePath
            End If
            If Dir(ReleaseFilePath & "\") = "" Then
                MkDir ReleaseFilePath
            End If
            If Dir(CodeFileName(i, CBackup)) <> "" Then
                Kill CodeFileName(i, CBackup)
            End If
            If Dir(CodeFileName(i, CRelease)) <> "" Then
                Kill CodeFileName(i, CRelease)
            End If
            If (CodeFileName(i, CStatus) = "����") Then
                Application.VBE.ActiveVBProject.VBComponents(CodeFileName(i, CMName)).Export (CodeFileName(i, CBackup))
                If TestCode = "�����汾" Then
                    Application.VBE.ActiveVBProject.VBComponents(CodeFileName(i, CMName)).Export (ReleaseFilePath & "\" & CodeFileName(i, CRelease))
                End If
            End If
        Next i
        Application.VBE.ActiveVBProject.VBComponents(CodeFileName(i, CMName)).Export (ReleaseFilePath & "\" & CodeFileName(i, CRelease))
        If TestCode = "�����汾" Then
            Call WriteLastLine(CodeFileName(0, CBackup), "'[�汾��]" & Range("H1").Value)
            Call WriteLastLine(ReleaseFilePath & "\" & CodeFileName(0, CRelease), "'[�汾��]" & Range("H1").Value)
            
            ReadMeisEmpty = ����Readme(ReleaseFilePath, "Readme.txt", BackupFilePath, "Readme.txt", "ģ��1", ReleaseFile, Version, RiviseDate)
            
            '����Git�����������ļ�
            If Not ReadMeisEmpty Then
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set MyTxtObj = fso.CreateTextFile(ReleaseFilePath & "\" & BatFile, True, False)
                MyTxtObj.WriteLine (Mid(ReleaseFilePath, 1, 2))
                MyTxtObj.WriteLine ("cd " & Mid(ReleaseFilePath, 3, Len(ReleaseFilePath) - 2))
                MyTxtObj.WriteLine ("git add .")
                TempStr = "git commit -m """
                TempStr = TempStr & Commit
                TempStr = TempStr & """"
                MyTxtObj.WriteLine (TempStr)
                MyTxtObj.WriteLine ("git pull origin master")
                MyTxtObj.WriteLine ("git push -u origin master")
                MyTxtObj.WriteLine ("exit")
                MyTxtObj.Close
                Shell (ReleaseFilePath & "\" & BatFile)
            Else
                Call MsgInfo(NoMsgBox, "ReadMe�ļ�Ϊ�գ��汾����ʧ�ܣ�")
            End If
        End If
    End If
    Set MyTxtObj = Nothing
    Set fso = Nothing
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
End Sub
Function ����Readme(ReleaseFilePath As String, ReleaseReadmeFile As String, BackupFilePath As String, BackupReadmeFile As String, ModuleName As String, ReleaseFile As String, Version As String, RiviseDate As String)
    On Error Resume Next
    Dim ModuleCount As Integer
    Dim i As Integer
    Dim k As Integer
    Dim LineCount As Integer
    Dim UpdateInfo As String
    Dim Status As String
    Dim isEmpty As Boolean
    Dim isError As String
    On Error Resume Next
    Status = DownFile(ThisWorkbook.Path, ReleaseReadmeFile, True)
    If Status = False Then
        Exit Function
    End If
    isEmpty = GetVersionFromFile(ThisWorkbook.Path & "\" & ReleaseReadmeFile)
    If Not isEmpty Then
        'ɾ����ʱReadme.txt
        If Dir(ThisWorkbook.Path & "\" & ReleaseReadmeFile) <> "" Then
            Open ThisWorkbook.Path & "\" & ReleaseReadmeFile For Input As #1
            Close #1
            Kill ThisWorkbook.Path & "\" & ReleaseReadmeFile
        End If
        ModuleCount = ModuleLastRivise(CSummary, CModuleCount)
        LineCount = ModuleLastRivise(CSummary, CSumLines)
        k = 1
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set MyTxtObj = fso.CreateTextFile(ReleaseFilePath & "\" & ReleaseReadmeFile, True, False)
        ModuleLastRivise(1, CModuleName) = "ģ��1"
        ModuleLastRivise(1, CFileName) = ReleaseFile
        ModuleLastRivise(1, CVersion) = Version
        ModuleLastRivise(1, CRiviseDate) = RiviseDate
        For i = 1 To ModuleCount
            MyTxtObj.WriteLine ("[ģ������]" & ModuleLastRivise(i, CModuleName))
            MyTxtObj.WriteLine ("[�ļ�����]" & ModuleLastRivise(i, CFileName))
            MyTxtObj.WriteLine ("[�޶��汾]" & ModuleLastRivise(i, CVersion))
            MyTxtObj.WriteLine ("[�޶�����]" & ModuleLastRivise(i, CRiviseDate))
            UpdateInfo = InputBox("������" & ModuleLastRivise(i, CModuleName) & "�˴θ���˵��" & vbCrLf & ModuleLastRivise(i, CUpdateInfo))
            MyTxtObj.WriteLine ("[����˵��]" & vbCrLf & ModuleLastRivise(i, CUpdateInfo)) & UpdateInfo
        Next i
        fso.CopyFile ReleaseFilePath & "\" & ReleaseReadmeFile, BackupFilePath & "\" & BackupReadmeFile
        MyTxtObj.Close
        Set fso = Nothing
        Set MyTxtObj = Nothing
    End If
    ����Readme = isEmpty
End Function
Sub ImportCode(Workbook As String, CodeFileName As String, wbSheetName As String)
    Dim fso As Object
    Dim Txtfile As Object
    Dim Str$
    Dim StrTxt
    Dim i, j As Integer
    Dim LineCount As Integer
    Dim CodeTXT$
    Dim CodeLineCount As Integer
    Dim ModuleCount As Integer
    On Error Resume Next
    Dim xlsApp As New Excel.Application '��Ҫ�ڹ���������EXCEL����Ŷ
    Dim xlsWorkBook As Excel.Workbook
    Workbooks(ThisWorkBookName).Activate
    Dim vbPro
    Dim WorkBookName As String
    Dim LineStart As Integer
    WorkBookName = ActiveWorkbook.Name
    Set vbPro = ActiveWorkbook.VBProject
    ModuleCount = 0
    With vbPro
         For i = .VBComponents.Count To 1 Step -1
            If (Mid(.VBComponents(i).Name, 1, 2) = "ģ��") Then
                ModuleCount = ModuleCount + 1
            End If
            If .VBComponents(i).Name = wbSheetName Then
                Set fso = CreateObject("scripting.filesystemobject")
                Set Txtfile = fso.OpenTextFile(CodeFileName, 1)
                Str = Txtfile.ReadAll
                Txtfile.Close
                StrTxt = Split(Str, vbLf)
                LineCount = UBound(StrTxt)
                j = LineCount
                While (StrTxt(j) = "")
                    LineCount = LineCount - 1
                    j = j - 1
                Wend
                LineStart = 1
                While (InStr(1, StrTxt(LineStart), "Private Sub") <> 1)
                    LineStart = LineStart + 1
                Wend
                CodeTXT = ""
                For j = LineStart To LineCount
                    CodeTXT = CodeTXT & StrTxt(j) & vbLf
                Next j
                .VBComponents(i).CodeModule.AddFromString CodeTXT
            End If
        Next i
        If ModuleCount = 0 Then
            If Dir(CodeFileName) <> "" Then
                ActiveWorkbook.VBProject.VBComponents.Import CodeFileName
            End If
        End If
    End With
End Sub
Sub ���¹��������(FilePath As String)
Dim FileList() As String
Dim SheetList() As String
Dim VBSheetName As String
Dim Status As String
Set vbPro = ActiveWorkbook.VBProject
    j = 1
    With vbPro
        For i = .VBComponents.Count To 1 Step -1
            LCount = .VBComponents(i).CodeModule.CountOfLines
            If .VBComponents(i).Name = CodeFileName(j, CMName) Then
                .VBComponents(i).CodeModule.DeleteLines 1, LCount
                '.VBComponents.Remove .VBComponents(i)
                If (CodeFileName(j, CStatus) = "����") Then
                    Status = DownFile(ThisWorkbook.Path, CodeFileName(j, CRelease), True)
                    If Status = True And Dir(ThisWorkbook.Path & "\" & CodeFileName(j, CRelease)) <> "" Then
                        Call ImportCode(ThisWorkbook.Name, FilePath & "\" & CodeFileName(j, CRelease), CodeFileName(j, CMName))
                        j = j + 1
                    End If
                End If
            End If
        Next i
    End With
End Sub
Sub Զ�̸��´���()
    On Error Resume Next
    Dim ModuleCount As Integer
    Dim ModuleName As String
    Dim ModuleFile As String
    Dim CurrentVersion As String
    Dim CurrentRiviseDate As String
    Dim UpdateInfo As String
    Dim wbList() As String
    Dim FileName As String
    Dim FileType As String
    Dim LastVersion As String
    Dim LastRiviseDate As String
    Dim CtrResult As String
    Dim Vbc As Object
    Dim Result As String
    Dim VersionFilePath As String
    Dim LastFilePath As String
    Dim LastReadme As String
    Dim LastBasFile As String
    Dim Status As Boolean
    Dim RemoteVersion As String
    Dim DownComplete As String
    Dim isError As String
    Dim AutoUpdate As String
    Application.ScreenUpdating = False
    Call �޶�רҵ����״̬
    Call ���ñ����ļ���Ϣ
    Application.ScreenUpdating = False
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    AutoUpdate = Range("H12").Value
    If (Range("H10").Value = "������Ϣ��") Then
        NoMsgBox = False
    ElseIf (Range("H10").Select = "��������Ϣ��") Then
        NoMsgBox = True
    End If
    LastFilePath = ThisWorkbook.Path
    LastReadme = "Readme.txt"
    If Dir((LastFilePath & "\" & LastReadme)) <> "" Then
        Open LastFilePath & "\" & LastReadme For Input As #1
        Close #1
        Kill LastFilePath & "\" & LastReadme
    End If
    If AutoUpdate = "�Զ�����" Then
        Update = vbYes
    Else
        Update = MsgBox("��������Զ�̷����������������°汾��" & vbCrLf & "��ʼ���´�����", vbYesNo, "Զ���Զ����´���")
    End If
    If Update = vbYes Then
        'Call MsgInfo(NoMsgBox, "��������Զ�̷����������������°汾��")
        Status = DownFile(LastFilePath, LastReadme, True)
        If Status = False Or Dir(LastFilePath & "\" & LastReadme) = "" Or GetLastLine(LastFilePath & "\" & LastReadme) = "�ļ�Ϊ��" Then
            GoTo ErrorSub
        End If
        Call GetVersionFromFile(LastFilePath & "\" & LastReadme)
        CurrentVersion = Range("H1").Value
        CurrentRiviseDate = Range("H2").Value
        CtrResult = StrComp(CurrentVersion, ModuleLastRivise(1, CVersion), vbTextCompare)
         'Զ�̴���汾�űȵ�ǰ����汾����
        If CtrResult = "-1" Then
            ModuleFile = ModuleLastRivise(1, CFileName)
            Status = DownFile(ThisWorkbook.Path, ModuleFile, False)
            If GetLastLine(ThisWorkbook.Path & "\" & ModuleFile) = "�ļ�Ϊ��" Then
                DownComplete = "1"
            Else
                RemoteVersion = Replace(GetLastLine(ThisWorkbook.Path & "\" & ModuleFile), "'[�汾��]", "")
                If Status = False Or Dir(ThisWorkbook.Path & "\" & ModuleFile) = "" Then
                    GoTo ErrorSub
                End If
                DownComplete = StrComp(ModuleLastRivise(1, CVersion), RemoteVersion, vbTextCompare)
            End If
            If DownComplete <> "0" Then
                Call MsgInfo(NoMsgBox, "�汾Ϊ" & LastVersion & "�Ĵ���δ���سɹ��������´��ļ��Զ��������´���!")
            Else
                Call ���¹��������(ThisWorkbook.Path)
                ModuleName = ModuleLastRivise(1, CModuleName)
                ModuleFile = ModuleLastRivise(1, CFileName)
                LastVersion = ModuleLastRivise(1, CVersion)
                LastRiviseDate = ModuleLastRivise(1, CRiviseDate)
                UpdateInfo = ModuleLastRivise(1, CUpdateInfo)
                For Each Vbc In ThisWorkbook.VBProject.VBComponents
                    If Vbc.Type = 1 And Mid(Vbc.Name, 1, 2) = "ģ��" Then
                        ThisWorkbook.VBProject.VBComponents.Remove Vbc
                    End If
                Next Vbc
                If Dir(ThisWorkbook.Path & "\" & ModuleFile) <> "" Then
                    ActiveWorkbook.VBProject.VBComponents.Import ThisWorkbook.Path & "\" & ModuleFile
                    ModuleCount = 0
                    For Each Vbc In ThisWorkbook.VBProject.VBComponents
                        If Vbc.Type = 1 And Mid(Vbc.Name, 1, 2) = "ģ��" Then
                            ModuleCount = ModuleCount + 1
                        End If
                        If ModuleCount = 1 Then
                            Vbc.Name = "ģ��1"
                        Else
                            Exit For
                        End If
                    Next Vbc
                    
                    Range("H1").Select
                    ActiveCell.FormulaR1C1 = _
                        "=""V""&R[2]C&"".""&TEXT(R[3]C,""00"")&"".""&TEXT(R[4]C,""00"")"
                    Range("H2").Value = LastRiviseDate
                    Range("H3").Value = Val(Mid(LastVersion, 2, 3))
                    Range("H4").Value = Val(Mid(LastVersion, 4, 2))
                    Range("H5").Value = Val(Mid(LastVersion, 7, 2))
                End If
                Call MsgInfo(NoMsgBox, "�Ѹ��´���汾Ϊ��" & LastVersion & "�޶����ڣ�" & LastRiviseDate)
                Worksheets("רҵ����״̬").Activate
                Range("H8").Value = "���¹�ʽ"
            End If
        Else
            Call MsgInfo(NoMsgBox, "��ģ�����汾�Ѿ�Ϊ���°汾!")
        End If
        If Dir(LastFilePath & "\" & LastReadme) <> "" Then
            Open LastFilePath & "\" & LastReadme For Input As #1
            Close #1
            Kill LastFilePath & "\" & LastReadme
        End If
        For i = 0 To 3
            If Dir(ThisWorkbook.Path & "\" & CodeFileName(i, CRelease)) <> "" Then
                Open ThisWorkbook.Path & "\" & CodeFileName(i, CRelease) For Input As #1
                Close #1
                Kill ThisWorkbook.Path & "\" & CodeFileName(i, CRelease)
            End If
        Next i
    End If
ErrorSub:
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
End Sub
Function ShellAndWait(cmdStr As String, isHide As Boolean) As String
    On Error Resume Next
    Dim oShell As Object, oExec As Object
    Dim oRun
    Dim isError As String
    Set oShell = CreateObject("WScript.Shell")
    If isHide Then
        oRun = oShell.Run(cmdStr, vbHide, True)
        isError = Err.Description
        If isError <> "" Then
            ShellAndWait = Err.Description
        ElseIf oRun = 1 Then
            ShellAndWait = "����ʧ��"
        End If
    Else
        Set oExec = oShell.exec(cmdStr)
        isError = Err.Description
        If isError <> "" Then
            ShellAndWait = Err.Description
        Else
            ShellAndWait = oExec.StdOut.ReadAll
        End If
        Set oExec = Nothing
    End If
    Set oShell = Nothing
End Function
Function GetLastLine(FileName As String)
    On Error Resume Next
    Dim Buf As String
    Dim LastRow As Long
    Dim fso As Object
    Dim TempBuf As String
    Dim isError As String
    Open FileName For Input As #1
    If EOF(1) Then
        TempBuf = "�ļ�Ϊ��"
        Close #1
    Else
        Close #1
        Set fso = CreateObject("Scripting.FileSystemObject")
        With fso.OpenTextFile(FileName, 1)
            Buf = .ReadAll
            .Close
        End With
        Set fso = Nothing
        LastRow = UBound(Split(Buf, vbLf))
        TempBuf = Split(Buf, vbLf)(LastRow)
        While (TempBuf = "")
            LastRow = LastRow - 1
            TempBuf = Split(Buf, vbLf)(LastRow)
        Wend
    End If
    GetLastLine = TempBuf
End Function
Sub WriteLastLine(FileName As String, WriteStr As String)
    Dim fso As Object
    Dim Txtfile As Object
    Dim Str$
    Set fso = CreateObject("scripting.filesystemobject")
    Set Txtfile = fso.OpenTextFile(FileName, 1)
    Str = Txtfile.ReadAll
    Txtfile.Close
    temp = Split(Str, vbLf)
    i = UBound(temp)
    While Not (InStr(1, temp(i), "�汾��") > 0 Or InStr(1, temp(i), "End Sub") > 0)
        i = i - 1
    Wend
    If InStr(1, temp(i), "�汾��") > 0 Then
        temp(i) = WriteStr
    ElseIf InStr(1, temp(i), "End Sub") > 0 Then
        temp(i + 1) = WriteStr
    End If
    Str = Join(temp, vbLf)
    Set Txtfile = fso.OpenTextFile(FileName, 2)
    Txtfile.Write Str
    Txtfile.Close
End Sub
Function GetVersionFromLocal(LocalFileName As String)
    Dim StrTxt() As String
    Dim n As Integer
    Dim StrTemp As String
    Dim i As Integer, x As Integer, y As Integer
    Dim Module As String
    Dim ModuleCount As Integer
    Dim LastVersion As String
    
    ModuleCount = 0
    UpdateInfo = ""
    Open LocalFileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, StrTemp
        n = n + 1
    Loop
    Close #1
    x = InStr(1, StrTemp, vbLf, vbTextCompare)
    If x <> 0 Then
        StrTxt = Split(StrTemp, vbLf)
        n = UBound(StrTxt) - LBound(StrTxt)
    Else
        Open LocalFileName For Input As #1
        ReDim Preserve StrTxt(1 To n)
        i = 1
        Do While Not EOF(1)
            Line Input #1, StrTxt(i)
            i = i + 1
        Loop
        Close #1
    End If
    For i = 1 To n
        If InStr(1, StrTxt(i), "[�޶��汾]") > 0 Then
            LastVersion = Replace(StrTxt(i), "[�޶��汾]", "")
        End If
    Next i
    GetVersionFromFile = LastVersion
End Function
Function DownFile(FilePath As String, FileName As String, isHide As Boolean)
    Dim TempFileName As String
    Dim Result As String
    Dim VersionFilePath As String
    Dim isError As String
    On Error Resume Next
    If Dir(ThisWorkbook.Path & "\wget.exe") = "" Then
        Call MsgInfo(NoMsgBox, "����" & ThisWorkbook.Path & "\wget.exe �ļ��Ƿ���ڣ�")
        DownFile = False
        GoTo ErrorSub
    End If
    RemoteFile = "https://raw.githubusercontent.com/hxguet/CQART/master/" & FileName
    TempFileName = ThisWorkbook.Path & "\wget.exe -O " & FilePath & "\" & FileName & " " & RemoteFile
    If (Dir(FilePath & "\" & FileName) <> "") Then
        Open FilePath & "\" & FileName For Input As #1
        Close #1
        Kill FilePath & "\" & FileName
    End If
    If (FileName = "Readme.txt") Then
        Result = ShellAndWait(TempFileName, isHide)
    Else
        Result = ShellAndWait(TempFileName, isHide)
    End If
    If Result <> "" Then
        DownFile = False
    Else
        DownFile = True
    End If
ErrorSub:
End Function
Function GetVersionFromFile(LocalFileName As String)
    Dim StrTxt() As String
    Dim n As Integer
    Dim StrTemp As String
    Dim i As Integer, x As Integer, y As Integer
    Dim Module As String
    Dim ModuleCount As Integer
    Dim UpdateInfo As String
    Dim isError As String
    Dim isEmpty As Boolean
    ModuleCount = 0
    UpdateInfo = ""
    isEmpty = False
    Open LocalFileName For Input As #1
    n = 0
    Do While Not EOF(1)
        Line Input #1, StrTemp
        n = n + 1
    Loop
    If EOF(1) And n = 0 Then
        isEmpty = True
        GoTo Error
    End If
    Close #1
    x = InStr(1, StrTemp, vbLf)
    '�ı��ļ�ÿ���Ի��з�����
    If x <> 0 Then
        StrTxt = Split(StrTemp, vbLf)
        n = UBound(StrTxt) - LBound(StrTxt)
    '�ı��ļ�ÿ���ѻس����н���
    Else
        Open LocalFileName For Input As #1
        ReDim Preserve StrTxt(1 To n)
        i = 1
        Do While Not EOF(1)
            Line Input #1, StrTxt(i)
            i = i + 1
        Loop
        Close #1
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set MyTxtObj = fso.CreateTextFile(LocalFileName, True, False)
    For i = 0 To n - 1
        x = InStr(1, StrTxt(i), "[ģ������]")
        If x > 0 Then
            ModuleCount = ModuleCount + 1
        End If
        MyTxtObj.WriteLine (StrTxt(i))
    Next i
    MyTxtObj.Close
    ReDim Preserve ModuleLastRivise(0 To ModuleCount, 0 To 5)
    ModuleLastRivise(CSummary, CModuleCount) = ModuleCount
    ModuleLastRivise(CSummary, CSumLines) = n
    k = 1
    For i = 0 To n - 1
        If InStr(1, StrTxt(i), "[ģ������]") > 0 Then
            ModuleLastRivise(k, CStartLine) = i + 1
            ModuleLastRivise(k, CModuleName) = Replace(StrTxt(i), "[ģ������]", "")
        ElseIf InStr(1, StrTxt(i), "[�ļ�����]") > 0 Then
            ModuleLastRivise(k, CFileName) = Replace(StrTxt(i), "[�ļ�����]", "")
        ElseIf InStr(1, StrTxt(i), "[�޶��汾]") > 0 Then
            ModuleLastRivise(k, CVersion) = Replace(StrTxt(i), "[�޶��汾]", "")
        ElseIf InStr(1, StrTxt(i), "[�޶�����]") > 0 Then
            ModuleLastRivise(k, CRiviseDate) = Replace(StrTxt(i), "[�޶�����]", "")
        ElseIf InStr(1, StrTxt(i), "[����˵��]") > 0 Then
            i = i + 1
            Do While i <= n - 1
                If InStr(1, StrTxt(i), "[�ļ�����]") = 0 Then
                    If StrTxt(i) <> "" Then
                        UpdateInfo = UpdateInfo & StrTxt(i) & vbCrLf
                    End If
                    i = i + 1
                End If
            Loop
            ModuleLastRivise(k, CUpdateInfo) = UpdateInfo
            k = k + 1
        End If
    Next i
Error: GetVersionFromFile = isEmpty
End Function
Sub MsgInfo(NoMsg As Boolean, Msg As String)
    If Not NoMsg Then
        MsgBox (Msg)
    End If
End Sub
Sub �޶��γ�Ŀ����ۺϷ�����ʽ()
    On Error Resume Next
    Dim MyShapes As Shapes
    Dim Shp As Shape
    Dim SchoolName As String
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    SchoolName = Range("B2").Value
    Worksheets("רҵ����״̬").Visible = False
    '�޶�"2-�γ�Ŀ����ۺϷ�������д��"���������ۻ��ڿγ̱������ҵ�ɼ���ʽ
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Application.EnableEvents = False
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    '2019.5.3�޶������ʵ��γ�ʵ��1��ʵ��2��ʵ��3������Ϊ100�֣��ϼƿ��˷ֳ���100�ֵ����
    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C="""","""",COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C24,R9C4&""-""&R7C2&""-""&""��֤""))"
    Range("R7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(COUNTBLANK(RC[-14]:RC[-1])=14,SUM(R6C4:R6C17)=0),"""",SUM('2-�γ�Ŀ����ۺϷ�������д��'!R7C4:R7C17)*100/SUM(R6C4:R6C17))"
    Range("R8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R7C18="""",R7C18=0),"""",ROUND(R7C18*100/R6C18,1))"
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Call �γ�Ŀ������༭����
    Range("N5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2]="""","""",'1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2])"
    Range("Q5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2]="""","""",'1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2])"
    If SchoolName = "�������Ϣ�밲ȫѧԺ" Then
        If Range("A26").Value = "��4����ҵҪ���ɶ�����" Then
            Rows("26:27").Select
            Selection.EntireRow.Hidden = True
            Range("A28").Select
            ActiveCell.FormulaR1C1 = "��4���Ľ���ʩ"
        End If
    End If
    Range("AH4:AQ4").Select
    Selection.Merge
    Range("AH5:AQ5").Select
    Selection.Merge
    Range("AH4:AQ4").Select
    ActiveCell.FormulaR1C1 = "=רҵ����״̬!R[-3]C[-27]&""��""&רҵ����״̬!R[-3]C[-26]"
    Range("AH5:AQ5").Select
    ActiveCell.FormulaR1C1 = _
        "=רҵ����״̬!R[-3]C[-27]&""��""&TEXT(רҵ����״̬!R[-3]C[-26],""YYYY��MM��DD��"")"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub �޶���ѧ���̵ǼǱ�ʽ()
    On Error Resume Next
    Dim temp As Boolean
    '"0-��ѧ���̵ǼǱ���д+��ӡ)"�������޶����⣬ѧ�ţ������ؼ��ʣ��������壬��������༭����
    Application.ScreenUpdating = False
    temp = Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Visible
    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Visible = True
    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A2:AG2").Select
    ActiveCell.FormulaR1C1 = _
        "=רҵ����״̬!RC[1]&"" ""&'2-�γ�Ŀ����ۺϷ�������д��'!R[2]C[1]&"" �γ�(����/����/ѡ��)��ѧ���̵ǼǱ�"""
    Range("B4:B5").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "ѧ��"
    Range("C4:C5").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "����"
    Range("D4:L4").Select
    Selection.Merge
    Range("M4:U4").Select
    Selection.Merge
    Range("Y4:Y5").Select
    Selection.Merge
    Range("Z4:Z5").Select
    Selection.Merge
    Range("AA4:AA5").Select
    Selection.Merge
    Range("AB4:AB5").Select
    Selection.Merge
    Range("AC4:AC5").Select
    Selection.Merge
    Range("AD4:AD5").Select
    Selection.Merge
    Range("AE4:AE5").Select
    Selection.Merge
    Range("AF4:AF5").Select
    Selection.Merge
    Range("AG4:AG5").Select
    Selection.Merge
    Xueqi = Range("AN1").Value
    SumCount = Range("AM1").Value
    If (SumCount <> 0) Then
        Range("A6:AG" & SumCount + 5).Select
        With Selection.Font
            .Name = "����"
            .Name = "Calibri"
            .Size = 9
        End With
        Range("A4:AG" & SumCount).Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
    End If
    Range("AB3:AE3").Select
    If Selection.MergeCells = True Then
        Range("AB3:AG3").Select
        Selection.Merge
    End If
    Columns("AH:AR").Select
    Selection.EntireColumn.Hidden = True
    Call ���ø�����ɫ
    '��ѧ���̵ǼǱ�����༭��������
    AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
    If (AllowEditCount <> 0) Then
        For i = 1 To AllowEditCount
            Sheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Protection.AllowEditRanges(1).Delete
        Next i
    End If
    ActiveSheet.Protection.AllowEditRanges.Add Title:="��¼��", Range:=Range("$D$6:$Y$185")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="���гɼ�", Range:=Range("$AA$6:$AA$185")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�ɼ��ȼ���", Range:=Range("$AV$6:$AW$21")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�ɼ������", Range:=Range("$AF$6:$AF$185")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="���������", Range:=Range("$AG$6:$AG$185")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�������", Range:=Range("$AJ$1")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="���۷�ʽ", Range:=Range("$AX$6")
    Range("$D$6:$Y$185").Select
    Selection.FormulaHidden = False
    Range("$AA$6:$AA$185").Select
    Selection.FormulaHidden = False
    
    Range("AT20:AU20").Select
    ActiveCell.FormulaR1C1 = "�����ٵ�"
    Range("AV20").Select
    ActiveCell.FormulaR1C1 = "��"
    Range("AW20").Value = 70
    Range("AT21:AU21").Select
    ActiveCell.FormulaR1C1 = "�������"
    Range("AV21").Select
    ActiveCell.FormulaR1C1 = "��"
    Range("AW21").Value = 80
    Range("AT20:AU20").Select
    Selection.Merge
    Range("AT21:AU21").Select
    Selection.Merge
    Call ���ñ����("AT6", "AW21", 9)
    Range("AT6:AW21").Select
    Selection.Font.Bold = False
    With Selection.Font
        .Name = "����"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    '������ҵ�Ǽ����������У��
    Range("D6:T" & (SumCount + 5)).Select
    Selection.FormulaHidden = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$AV$6:$AV$22"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��ҵ����͵����ǼǴ���"
        .InputMessage = ""
        .ErrorMessage = "�밴��ѧ���̵ǼǱ����Ϸ��ɼ��ȼ�������д���������ͳٵ����Ų�������ҵ�ȼ�������ͬ��"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Visible = temp
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.ScreenUpdating = True
End Sub
Sub �޶�ƽʱ�ɼ���()
    On Error Resume Next
    Application.ScreenUpdating = False
    Worksheets("ƽʱ�ɼ���").Visible = True
    Worksheets("ƽʱ�ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC25+RC29=0,0,ROUND((RC32)/(R5C25+R5C29),0))"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("ƽʱ�ɼ���").Visible = False
    Application.ScreenUpdating = True
End Sub
Sub �޶�רҵ����״̬()
    On Error Resume Next
     '�޶�רҵ����״̬������
    Dim MyShapes As Shapes
    Dim Shp As Shape
    Dim LastVersion As String
    Application.ScreenUpdating = False
    Set MyShapes = Worksheets("רҵ����״̬").Shapes
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:I12").Select
    With Selection.Validation
        .Delete
    End With
    Range("H1").Select
    ActiveCell.FormulaR1C1 = _
        "=""V""&R[2]C&"".""&TEXT(R[3]C,""00"")&"".""&TEXT(R[4]C,""00"")"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "�汾��"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "�޶�����"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "���汾��"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "���汾��"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "�޸��汾��"
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "���뷢��·��"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "���뱸��·��"
    
    Range("G1:G12").Select
    Selection.Font.Bold = True
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Set NewShp = ActiveSheet.Buttons.Add(745, 2, 86, 25) '��λ�ø߶ȣ�λ�ÿ�ȣ���ť�߶ȣ���ť��ȣ�
    NewShp.Characters.Text = "�����汾"
    NewShp.OnAction = "���ɰ汾��"
    NewShp.Font.Name = "΢���ź�"
    NewShp.Font.Size = 14
    NewShp.Font.ColorIndex = 3
    
    Range("K6").Select
    AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
    If (AllowEditCount <> 0) Then
        For i = 1 To AllowEditCount
            Sheets("רҵ����״̬").Protection.AllowEditRanges(1).Delete
        Next i
    End If
    ActiveSheet.Protection.AllowEditRanges.Add Title:="רҵ", Range:=Range("B4:C12")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="ѧԺ", Range:=Range("B2:D2")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�޸��汾��", Range:=Range("H5")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="���뷢���汾", Range:=Range("H9")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="��Ϣ��״̬", Range:=Range("H10")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="����PDF����ĵ�", Range:=Range("H11")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�Զ����´���", Range:=Range("H12")
    Range("G9").Select
    ActiveCell.FormulaR1C1 = "���뷢���汾"
    Range("H9").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="���԰汾,�����汾"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("G10").Select
    ActiveCell.FormulaR1C1 = "��Ϣ��״̬"
    Range("H10").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="������Ϣ��,��������Ϣ��"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("G11").Select
    ActiveCell.FormulaR1C1 = "����PDF����ĵ�"
    Range("H11").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="��PDF,����PDF"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("G12").Select
    ActiveCell.FormulaR1C1 = "�Զ����´���"
    Range("H12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="�Զ�����,�ֶ�����"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("B4:C12,B2:F2").Select
    Range("B2").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
        .PatternTintAndShade = 0
    End With
    Rows("1:12").Select
    Selection.RowHeight = 30
    Call ���ñ����("A2", "H12", 12)
    Columns("H:H").Select
    Selection.ColumnWidth = 20
    Range("G1:H12").Select
    With Selection.Font
        .Name = "����"
        .Size = 12
    End With
    Selection.Font.Bold = True
    Range("A2:F12").Select
    With Selection.Font
        .Name = "����"
        .Size = 12
    End With
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
    Application.ScreenUpdating = True
End Sub
Sub �޶���ҵҪ���ɶ����۱�()
    On Error Resume Next
    Dim SchoolName As String
    Application.ScreenUpdating = False
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    SchoolName = Range("B2").Value
    Worksheets("רҵ����״̬").Visible = False
    Worksheets("3-�ۺϷ�������ӡ��").Visible = True
    Worksheets("3-�ۺϷ�������ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    If SchoolName = "�������Ϣ�밲ȫѧԺ" Then
        If Range("A9").Value = "��4����ҵҪ���ɶ�����" Then
            Rows("9:10").Select
            Selection.Delete Shift:=xlUp
            Range("A11").Select
            ActiveCell.FormulaR1C1 = "��4���Ľ���ʩ"
        End If
    End If
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("3-�ۺϷ�������ӡ��").Visible = False
    Application.ScreenUpdating = True
End Sub
Sub �γ�Ŀ������༭����()
    On Error Resume Next
    Dim AllowEditCount As Integer
    Dim i As Integer
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
    If (AllowEditCount <> 0) Then
        For i = 1 To AllowEditCount
            Sheets("2-�γ�Ŀ����ۺϷ�������д��").Protection.AllowEditRanges(1).Delete
        Next i
    End If
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�κ�", Range:=Range("B3")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="����", Range:=Range("D3:O3")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�Ƿ���֤", Range:=Range("Q3")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="��֤רҵ", Range:=Range("B7")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="��д����", Range:=Range("B8")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�γ�Ŀ��", Range:=Range("B11:B20")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="���ۻ���֧�ű���", Range:=Range("D11:Q20")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="���˽������", Range:=Range("B22")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="��Ч��ʩ", Range:=Range("B23")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�γ�Ŀ���ɶ�", Range:=Range("B25")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="��ҵҪ���ɶ�", Range:=Range("B27")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="�Ľ���ʩ", Range:=Range("B28")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="��ʽ����", Range:=Range("Y1")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="����", Range:=Range("AE2:AE7")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="���ۻ�������1", Range:=Range("D2:E2")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="���ۻ�������2", Range:=Range("L2:M2")
    Range("D2:E2").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Range("L2:M2").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub
Sub ���ĵ�()
On Error Resume Next
Dim Grade As String
Dim Major As String
Dim i As Integer
Dim CourseCount As Integer
Dim PointCount As Integer
Dim MatrixSheet As String
On Error Resume Next
    Application.EnableEvents = True
    Call MsgInfo(NoMsgBox, "���µ����ѧ����ѧ��������רҵ�������Ϣ�����Եȡ�����")
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Grade = Range("D9").Value
    Major = Range("B7").Value
    On Error Resume Next
    If Sheets("ѧ������") Is Nothing Then
        Worksheets("ʹ�ð���").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "ѧ������"
    End If
    Sheets("ѧ������").Visible = True
    Worksheets("ѧ������").Activate
    If (Application.WorksheetFunction.CountIf(Range("E:E"), Grade) < 1) Or (Application.WorksheetFunction.CountIf(Range("F:F"), Major) < 1) Then
        Call ����ѧ������
    End If

    Sheets("ѧ������").Visible = False
    
    If Sheets("רҵ����״̬") Is Nothing Then
        Worksheets("ʹ�ð���").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "רҵ����״̬"
        Call �½�רҵ����״̬������
    End If
    Call �������
    Sheets("רҵ����״̬").Visible = False
    Call ָ������ݱ�ʽ
    Call �γ�Ŀ����ۺϷ�����ʽ
    Worksheets("4-�����������棨��д+��ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    ActiveSheet.Shapes.Range(Array("Button 5")).Select
    Selection.OnAction = "��ӡ"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
End Sub
Sub �������()
On Error Resume Next
Dim Major As String
Dim MatrixSheet As String
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Major = Range("B7").Value
    Worksheets("רҵ����״̬").Activate
    Call ��ȡ��ҵҪ�������Ϣ(Major)
    Sheets("רҵ����״̬").Visible = True
    On Error Resume Next
    Worksheets("רҵ����״̬").Activate
    MatrixSheet = Application.Index(Range("C4:C" & MajorLastRow), Application.Match(Major, Range("B4:B" & MajorLastRow), 0))
    If Sheets(MatrixSheet) Is Nothing Then
        Worksheets("��ҵҪ��-ָ������ݱ�").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = MatrixSheet
    End If
    Sheets(MatrixSheet).Visible = True
    Worksheets(MatrixSheet).Activate
    CourseCount = Application.CountA(Range("D4:D200"))
    PointCount = Application.CountA(Range("E4:AS200"))
    Sheets(MatrixSheet).Visible = False
    Worksheets("רҵ����״̬").Activate
    If (Range("D" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = "����") Then
        'If (Range("E" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value <> CourseCount) Or (Range("F" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value <> PointCount) Then
            Call �����ҵҪ�����(Major)
        'End If
    End If
    Sheets("רҵ����״̬").Visible = False
End Sub
Sub �����¼�����()
    On Error Resume Next
    Application.EnableEvents = True
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("B7").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=OFFSET(רҵ����״̬!$B$4,,,COUNTA(רҵ����״̬!$B$4:$B$12))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("Q3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$X$3:$X$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Call �γ�Ŀ������༭����
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H10").Value = "������Ϣ��"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
End Sub
Sub �ύǰ���()
    Dim CourseName As String
    Dim Term As String
    Dim Major As String
    Dim CourseNum As String
    Dim ShoolName As String
    Dim Teacher As String
    Dim DateCompleted As String
    Dim CourseTargetCount As Integer
    Dim RequirementCount As Integer
    Dim RequirementReachCount As Integer
    Dim i As Integer
    Dim LinkCount As Integer
    Dim ErrorMsg  As String
    Dim ErrNum As Integer
    Dim NoError As Boolean
    Dim ThisFileName As String
    Dim IdentifyStatus As String
    Dim isError As String
    On Error Resume Next
    NoError = True
    Application.ScreenUpdating = False
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ShoolName = Range("B2").Value
    ErrNum = 0
    ErrorMsg = ""
    ThisFileName = Mid(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1)
    '���2-�γ�Ŀ����ۺϷ�������д�� ������
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    CourseTargetCount = Application.CountA(Range("B11:B20"))
    IdentifyStatus = Range("$Q$3").Value
    If Range("$B$3").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д��������δ��д���γ���š�" & vbCrLf
    End If
    If IdentifyStatus = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д��������δѡ����֤״̬��" & vbCrLf
    End If
    If Range("$B$7").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д��������δѡ����֤רҵ��" & vbCrLf
    End If
    If CourseTargetCount = 0 Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д��������δ��д���γ�Ŀ�꡿" & vbCrLf
    End If
    If Range("$B$8").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д��������δ��д����д���ڡ�" & vbCrLf
    End If

    For i = 1 To 5
        Mark = Application.Index(Range("M7:Q7"), Application.Match(Cells(2, 2 * i + 2).Value, Range("M5:Q5"), 0))
        If Cells(3, 2 * i + 2).Value <> "" And Cells(3, 2 * i + 2).Value <> "0" Then
            If Mark = "" Or Mark = 0 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д�����������ۻ���" & Cells(2, 2 * i + 2).Value & "ƽ���÷�Ϊ��,���飺��1�����û�и����ۻ��ڣ�ɾ����������2���Ծ�ɼ��ǼǱ���ȱ�ٸ���ɼ�" & vbCrLf
            End If
        ElseIf Cells(3, 2 * i + 2).Value = "" Then
            If Mark <> "" Or Mark = 0 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д�����������ۻ���" & Cells(2, 2 * i + 2).Value & "����Ϊ�գ���ƽ���÷ֲ�Ϊ��,���飺��1���Ƿ�ȱ�ٸ����ۻ��ڣ���2���Ծ�ɼ��ǼǱ���ɼ�ʱ�Ƿ�ർ���˸���ɼ�" & vbCrLf
            End If
        End If
    Next i
    LinkCount = Application.CountA(Range("D5:Q5"))
    For i = 4 To LinkCount + 4
        If Cells(7, i).Value <> "" Then
            If Application.CountBlank(Cells(11, i).Resize(10, 1)) = 10 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "�����ۻ���" & Cells(5, i).Value & "��ƽ���ɼ����������ۻ���δ֧�ſγ�Ŀ�ꡣ" & vbCrLf
            End If
        Else
            If Application.CountBlank(Cells(11, i).Resize(10, 1)) <> 10 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "�����ۻ���" & Cells(5, i).Value & "û��ƽ���ɼ����������ۻ���֧���˿γ�Ŀ�ꡣ" & vbCrLf
            End If
        End If
    Next i
    
    For i = 11 To 20
        If Range("B" & i).Value <> "" Then
            If (Range("R" & i).Value = "") Or (Range("R" & i).Value <> 100) Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "���γ�Ŀ��" & i - 10 & "֧�ű����ϼƲ���100%��" & vbCrLf
            End If
        ElseIf (Range("R" & i).Value <> "") Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "��û�пγ�Ŀ��" & i - 10 & "����ɾ����Ӧ��֧�ű���" & vbCrLf
        End If
    Next i
    If Range("$B$22").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��ȱ�٣�1�����˽������" & vbCrLf
    End If
    If Range("$B$23").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��ȱ�٣�2����Ч�Ľ�ѧ�����ʹ�ʩ" & vbCrLf
    End If
    If Range("$B$25").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��ȱ�٣�3���γ�Ŀ���ɶ�����" & vbCrLf
    End If
    If ShoolName = "���ӹ������Զ���ѧԺ" Then
        If Range("$B$27").Value = "" Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "��ȱ�٣�4����ҵҪ���ɶ�����" & vbCrLf
        End If
    End If
    If Range("$B$28").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��ȱ�٣�5���Ľ���ʩ" & vbCrLf
    End If

    For i = 0 To CourseTargetCount - 1
        Worksheets("�γ�Ŀ���ɶȻ���������").Activate
        If (Not IsNumeric(Cells(2, 2 * i + 9).Value) Or Cells(2, 2 * i + 9).Value = "" Or Cells(2, 2 * i + 9).Value = 0) Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "���γ�Ŀ��" & i + 1 & "����������ȷ���顣" & vbCrLf
        End If
    Next i
    
    If (ShoolName = "���ӹ������Զ���ѧԺ") Then
        Worksheets("3-��ҵҪ�����ݱ���д��").Visible = True
        Worksheets("3-��ҵҪ�����ݱ���д��").Activate
        RequirementCount = Application.CountA(Range("C7:C18")) - Application.CountBlank(Range("C7:C18"))
        RequirementReachCount = Application.Count(Range("D7:D18"))
        For i = 0 To Application.CountA(Range("B7:B18")) - 1
            If Range("O" & i + 11).Value <> "" And Range("O" & i + 11).Value <> 100 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "��" & Cells(i + 7, 2).Value & "֧�ű����ϼƲ�Ϊ100��" & vbCrLf
            End If
            If (Cells(i + 7, 3).Value = "��") And (Not IsNumeric(Cells(i + 7, 4).Value) Or Cells(i + 7, 4).Value = "" Or Cells(i + 7, 4).Value = 0) Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "��" & Cells(i + 7, 2).Value & "����������ȷ" & vbCrLf
            End If
        Next i
    End If

    Worksheets("�γ�Ŀ���ɶȻ���������").Visible = True
    Worksheets("�γ�Ŀ���ɶȻ���������").Activate
    If (Range("B2").Value = "") Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "���γ���Ϣ��ȡ����������������Դ-��ѧ����.xls�в��䡾ѧ�ڡ�" & vbCrLf
    ElseIf Range("C2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "���γ���Ϣ��ȡ����������������Դ-��ѧ����.xls�в��䡾�γ����ơ�" & vbCrLf
    ElseIf Range("E2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "���γ���Ϣ��ȡ����������������Դ-��ѧ����.xls�в��䡾������ʦ��" & vbCrLf
    ElseIf Range("G2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "���γ���Ϣ��ȡ����������������Դ-��ѧ����.xls�в��䡾ѧ�֡�" & vbCrLf
    End If
    Worksheets("�γ�Ŀ���ɶȻ���������").Visible = False
    Worksheets("��ҵҪ���ɶȻ���������").Visible = False
    

    If IdentifyStatus = "����֤" Or IdentifyStatus = "��֤δ�ύ�ɼ�" Then
        Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
        If (Application.WorksheetFunction.CountIf(Range("Z6:Z185"), "ȡ��") <> Application.WorksheetFunction.CountIf(Range("AD6:AD185"), "ȡ��")) Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "������ͬѧƽʱ�ɼ�Ϊ0�����ʵ�Ƿ�ȡ�������ʸ���ȡ�������ڳɼ������ѡ��ȡ����" & vbCrLf
            For i = 6 To 185
                If (Range("B" & i).Value <> "") And (Range("Z" & i).Value = "ȡ��") Then
                    If (Range("AD" & i).Value = "") Then
                        ErrorMsg = ErrorMsg & Range("B" & i).Value & "  "
                    End If
                End If
            Next i
            ErrorMsg = ErrorMsg & vbCrLf
        End If
    End If
    '��顰1-�Ծ�ɼ��ǼǱ���д����������
    Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
    For i = 1 To 9
        If Cells(3, i + 4) = "" Then
            If Application.WorksheetFunction.Count(Cells(4, i + 4).Resize(403, 1)) <> 0 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "���Ծ�ɼ��ǼǱ���" & Cells(2, i + 4) & "����Ϊ�գ��������гɼ�������д���֡�" & vbCrLf
            End If
        Else
            If Application.WorksheetFunction.Count(Cells(4, i + 4).Resize(403, 1)) = 0 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "���Ծ�ɼ��ǼǱ���" & Cells(2, i + 4) & "���ֲ�Ϊ�գ�������û�гɼ���" & vbCrLf
            End If
        End If
    Next i
 
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Term = Range("$B$2").Value
    Term = Mid(Term, 3, 2) & "-" & Mid(Term, 8, 2) & "-" & Mid(Term, 14, 1)
    CourseNum = Range("$B$3").Value
    CourseName = Range("$B$4").Value
    Teacher = Range("$B$5").Value
    Major = Range("$B$7").Value
    If Dir(ThisWorkbook.Path & "\���󱨸�\" & ThisFileName & "-�����鱨��.txt") <> "" Then
        Open ThisWorkbook.Path & "\���󱨸�\" & ThisFileName & "-�����鱨��.txt" For Input As #1
        Close #1
        Kill ThisWorkbook.Path & "\���󱨸�\" & ThisFileName & "-�����鱨��.txt"
    End If
    If ErrorMsg <> "" Then
        If Dir(ThisWorkbook.Path & "\���󱨸�\") = "" Then
            MkDir ThisWorkbook.Path & "\���󱨸�\"
        End If
        Call CreateTXTfile(ThisWorkbook.Path & "\���󱨸�\" & ThisFileName & "-�����鱨��.txt", ErrorMsg, True)
        NoError = False
    End If
End Sub
Sub ��ӡ()
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H11").Value = "��PDF"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
    Call ����PDF
End Sub
Sub ����PDF()
    On Error Resume Next
    '���ø�ʽ ��
    Dim CurrentWorksheet As String
    Dim CourseName As String
    Dim Term As String
    Dim Major As String
    Dim CourseNum As String
    Dim Teacher As String
    Dim PDFFileName As String
    Dim SumRow As Integer
    Dim DateCompleted As String
    Dim CourseTargetCount As Integer
    Dim RequirementCount As Integer
    Dim RequirementReachCount As Integer
    Dim i As Integer
    Dim SchoolName As String
    Dim LinkCount As Integer
    Dim ErrorMsg  As String
    Dim ErrNum As Integer
    Dim NoError As Boolean
    Dim PDFFilePath As String
    CurrentWorksheet = ActiveSheet.Name
    Application.ScreenUpdating = False
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    If (Range("H11").Value = "��PDF") Then
        isOpenAfterPublish = True
    ElseIf (Range("H11").Value = "����PDF") Then
        isOpenAfterPublish = False
    End If
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
    
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Term = Range("$B$2").Value
    Term = Mid(Term, 3, 2) & "-" & Mid(Term, 8, 2) & "-" & Mid(Term, 14, 1)
    CourseNum = Range("$B$3").Value
    CourseName = Range("$B$4").Value
    Teacher = Range("$B$5").Value
    Major = Range("$B$7").Value
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    SchoolName = Range("B2").Value
    Worksheets("רҵ����״̬").Visible = False
    PDFFilePath = ThisWorkbook.Path & "\PDF��"
    If Dir(PDFFilePath) = "" Then
        MkDir PDFFilePath
    End If
    PDFFileName = Term & "-" & CourseNum & "-" & Major & "-" & Teacher & "-" & CourseName
    '�����鱨�治���ڣ�û�д���
    If (Dir(ThisWorkbook.Path & "\���󱨸�\" & PDFFileName & "-�����鱨��.txt") = "") Then
        Call ��������ʽ
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        If Range("$Q$3").Value = "����֤" Then
            Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
            ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
            Call Excel2PDF("0-��ѧ���̵ǼǱ���д+��ӡ)", PDFFilePath, PDFFileName & "--��ѧ���̵ǼǱ�.pdf", isOpenAfterPublish)
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
            
            Worksheets("4-�����������棨��д+��ӡ��").Activate
            Call ȡ����������������д������ɫ
            
            Call Excel2PDF("4-�����������棨��д+��ӡ��", PDFFilePath, PDFFileName & "--������������.pdf", isOpenAfterPublish)
            Call ������������������д������ɫ
        ElseIf Range("$Q$3").Value = "��֤δ�ύ�ɼ�" Then
            '��ӡ��ѧ���̵ǼǱ�
            Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
            ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
            Call Excel2PDF("0-��ѧ���̵ǼǱ���д+��ӡ)", PDFFilePath, PDFFileName & "--��ѧ���̵ǼǱ�.pdf", isOpenAfterPublish)
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
            
            Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = True
            Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
            ActiveSheet.PageSetup.CenterFooter = ""
            
            Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = True
            Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
            ActiveSheet.PageSetup.CenterFooter = ""
            ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
            Rows("11:21").Select
            Selection.EntireRow.Hidden = False
            SumRow = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
            Rows(SumRow + 11 & ":20").Select
            Selection.EntireRow.Hidden = True
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
            
            Call Excel2PDF("1-�γ�Ŀ���ɶ����ۣ���ӡ��", PDFFilePath, PDFFileName & "--�γ�Ŀ�����������.pdf", isOpenAfterPublish)
            Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = False
    
            Select Case SchoolName
                Case "�������Ϣ�밲ȫѧԺ"
                    Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
                    Sheets("3-�ۺϷ�������ӡ��").Visible = True
                    Worksheets("3-�ۺϷ�������ӡ��").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    
                    
                    Call Excel2PDF("3-�ۺϷ�������ӡ��", PDFFilePath, PDFFileName & "--�γ��ۺϷ���.pdf", isOpenAfterPublish)
                    Sheets("3-�ۺϷ�������ӡ��").Visible = False
                    
                    Worksheets("4-�����������棨��д+��ӡ��").Activate
                    Call ȡ����������������д������ɫ
                    
                    Call Excel2PDF("4-�����������棨��д+��ӡ��", PDFFilePath, PDFFileName & "--������������.pdf", isOpenAfterPublish)
                    
                    Call ������������������д������ɫ
                Case "���ӹ������Զ���ѧԺ"
                    Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = True
                    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    
                    Call Excel2PDF("2-��ҵҪ���ɶ����ۣ���ӡ��", PDFFilePath, PDFFileName & "--��ҵҪ�����������.pdf", isOpenAfterPublish)
                    Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
                    Sheets("3-�ۺϷ�������ӡ��").Visible = True
                    Worksheets("3-�ۺϷ�������ӡ��").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    
                    
                    Call Excel2PDF("3-�ۺϷ�������ӡ��", PDFFilePath, PDFFileName & "--�γ��ۺϷ���.pdf", isOpenAfterPublish)
                    Sheets("3-�ۺϷ�������ӡ��").Visible = False
                    
                    Worksheets("4-�����������棨��д+��ӡ��").Activate
                    Call ȡ����������������д������ɫ
                    
                    Call Excel2PDF("4-�����������棨��д+��ӡ��", PDFFilePath, PDFFileName & "--������������.pdf", isOpenAfterPublish)
                    
                    Call ������������������д������ɫ
                End Select
        ElseIf Range("$Q$3").Value = "��֤���ύ�ɼ�" Then
            Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = True
            Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
            ActiveSheet.PageSetup.CenterFooter = ""
            
            Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = True
            Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
            ActiveSheet.PageSetup.CenterFooter = ""
            ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
            Rows("11:20").Select
            Selection.EntireRow.Hidden = False
            SumRow = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
            Rows(SumRow + 11 & ":20").Select
            Selection.EntireRow.Hidden = True
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
            
            Call Excel2PDF("1-�γ�Ŀ���ɶ����ۣ���ӡ��", PDFFilePath, PDFFileName & "--�γ�Ŀ�����������.pdf", isOpenAfterPublish)
            Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = False
            Select Case SchoolName
                Case "�������Ϣ�밲ȫѧԺ"
                    Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
                    Worksheets("3-�ۺϷ�������ӡ��").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    Call Excel2PDF("3-�ۺϷ�������ӡ��", PDFFilePath, PDFFileName & "--�γ��ۺϷ���.pdf", isOpenAfterPublish)
                    Sheets("3-�ۺϷ�������ӡ��").Visible = False
                Case "���ӹ������Զ���ѧԺ"
                    Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = True
                    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    Call Excel2PDF("2-��ҵҪ���ɶ����ۣ���ӡ��", PDFFilePath, PDFFileName & "--��ҵҪ�����������.pdf", isOpenAfterPublish)
                    Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
                    Sheets("3-�ۺϷ�������ӡ��").Visible = True
                    Worksheets("3-�ۺϷ�������ӡ��").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    Call Excel2PDF("3-�ۺϷ�������ӡ��", PDFFilePath, PDFFileName & "--�γ��ۺϷ���.pdf", isOpenAfterPublish)
                    Sheets("3-�ۺϷ�������ӡ��").Visible = False
            End Select
        Else
            Call MsgInfo(NoMsgBox, "2-�γ�Ŀ����ۺϷ�������д���������еġ��Ƿ���֤��δѡ��")
        End If
    End If
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    Worksheets(CurrentWorksheet).Activate
End Sub
Sub ��ҵҪ�����ݱ�ʽ()
    On Error Resume Next
    Dim SchoolName As String
    Dim AllowEditCount As Integer
    Dim i As Integer
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    SchoolName = Range("B2").Value
    Worksheets("3-��ҵҪ�����ݱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Select Case SchoolName
        Case "�������Ϣ�밲ȫѧԺ"
            '��Ժר��
            Range("A7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(ROW(RC1)-ROW(R7C1)+1=1,1,IF(ROW(RC1)-ROW(R7C1)+1<=MAX('��ҵҪ��-ָ������ݱ�'!C1),R[-1]C1+1,""""))"
            Range("B7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(ISNA(MATCH('3-��ҵҪ�����ݱ���д��'!RC[-1],'��ҵҪ��-ָ������ݱ�'!C[-1],0)),"""",""��ҵҪ��""&INDEX('��ҵҪ��-ָ������ݱ�'!C[1],MATCH('3-��ҵҪ�����ݱ���д��'!RC[-1],'��ҵҪ��-ָ������ݱ�'!C[-1],0)))"
            Range("C7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(OR(R3C2="""",ISERROR(VLOOKUP(MID(RC[-1],5,LEN(RC[-1])-4),'��ҵҪ��-ָ������ݱ�'!R6C3:R46C6,4,0))),"""",IF(VLOOKUP(MID(RC[-1],5,LEN(RC[-1])-4),'��ҵҪ��-ָ������ݱ�'!R6C3:R46C6,4,0)>0,""��"",""""))"
            Range("E7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(RC2="""","""",IF(ROW(RC[-3])-ROW(R7C[-3])=COLUMN(R[-3]C)-COLUMN(R[-3]C5),100,""""))"
            Range("E7").Select
            Selection.AutoFill Destination:=Range("E7:N7"), Type:=xlFillDefault
            Range("E7:N7").Select
            Range("E7:N7").Select
            Selection.AutoFill Destination:=Range("E7:N16"), Type:=xlFillDefault
            Range("E7:N16").Select
            'ͳ������༭�����������ȫ��ɾ��
            AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
            If (AllowEditCount <> 0) Then
                For i = 1 To AllowEditCount
                    Sheets("3-��ҵҪ�����ݱ���д��").Protection.AllowEditRanges(1).Delete
                Next i
            End If
            Range("E7:N18").Select
            Range("N18").Activate
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Case "���ӹ������Զ���ѧԺ"
            '��Ժר��
            Range("A7").Select
            ActiveCell.FormulaR1C1 = "=ROW(RC1)-ROW(R7C1)+1"
            Range("B7").Select
            ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",""��ҵҪ��""&RC1)"
            Range("C7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(OR(R3C2="""",ISERROR(VLOOKUP(RC[-1],'��ҵҪ��-ָ������ݱ�'!R6C2:R46C6,5,0))),IF(RC[1]<>"""",""��"",""""),IF(VLOOKUP(RC[-1],'��ҵҪ��-ָ������ݱ�'!R6C2:R46C6,5,0)>0,""��"",IF(RC[1]<>"""",""��"","""")))"
            'ͳ������༭�����������ȫ��ɾ��
            AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
            If (AllowEditCount <> 0) Then
                For i = 1 To AllowEditCount
                    Sheets("3-��ҵҪ�����ݱ���д��").Protection.AllowEditRanges(1).Delete
                Next i
            End If
            Range("E7:N18").Select
            Range("N18").Activate
            ActiveSheet.Protection.AllowEditRanges.Add Title:="����1", Range:=Range( _
                "E7:N18")
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.14996795556505
                .PatternTintAndShade = 0
            End With
    End Select
    
    Range("D7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R3C[-2]="""","""",IF(RC[11]=100,ROUND(SUM(RC[14]:RC[23]),0),""""))"
    Range("O7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SUM(RC[-10]:RC[-1])=0,"""",IF(SUM(RC[-10]:RC[-1])<>100,""�����ϼƲ�Ϊ100"",SUM(RC[-10]:RC[-1])))"

    Range("Q7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]="""","""",IF(RC[-2]<>100,1,IF(RC[-14]<>""��"",1,"""")))"
    Range("R7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]="""","""",IF(RC[-3]<>100,1,IF(RC[-15]<>""��"",1,"""")))"
    Range("S7").Select
    Selection.AutoFill Destination:=Range("R7:S7"), Type:=xlFillDefault
    Range("R7:S7").Select
    Range("R7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-13]="""",R5C[-13]=""""),0,RC[-13]*R5C[-13]/100)"
    Range("R7").Select
    Selection.AutoFill Destination:=Range("R7:AA7"), Type:=xlFillDefault
    Range("A7:D7").Select
    Selection.AutoFill Destination:=Range("A7:D18"), Type:=xlFillDefault
    Range("A7:D18").Select
    
    Range("O7:O7").Select
    Selection.AutoFill Destination:=Range("O7:O18"), Type:=xlFillDefault
    Range("O7:O18").Select
    Range("Q7:AA7").Select
    Selection.AutoFill Destination:=Range("Q7:AA18"), Type:=xlFillDefault
    Range("Q7:AA18").Select

    Range("S1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R7C17:R18C17)"
    Range("S3").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
End Sub
Sub ���ñ������()
    On Error Resume Next
    Application.ScreenUpdating = False
    Worksheets("3-��ҵҪ�����ݱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("E1:K1").Select
    ActiveCell.FormulaR1C1 = _
        "='2-�γ�Ŀ����ۺϷ�������д��'!R[6]C[-3]&""רҵ-��ҵҪ��������������ݱ�"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:R1").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[1]C[1]&""  �γ�Ŀ����������۱�"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:N1").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[1]C[1]&""  ��ҵҪ����������۱�"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.ScreenUpdating = True
End Sub
Sub ��ҵҪ���ɶ����۹�ʽ()
    On Error Resume Next
    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = True
    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "='3-��ҵҪ�����ݱ���д��'!R[1]C"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "='3-��ҵҪ�����ݱ���д��'!R[1]C"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]="""","""",VLOOKUP(RC[-1],'3-��ҵҪ�����ݱ���д��'!R7C[-1]:R18C[1],3,0))"
    Range("D4:M4").Select
    Selection.FormulaArray = "=TRANSPOSE('2-�γ�Ŀ����ۺϷ�������д��'!R[7]C[-1]:R[16]C[-1])"
    Range("B6:C6").Select
    Selection.AutoFill Destination:=Range("B6:C17"), Type:=xlFillDefault
    Range("B6:C17").Select
    Range("D6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('3-��ҵҪ�����ݱ���д��'!R[1]C[1]="""","""",'3-��ҵҪ�����ݱ���д��'!R[1]C[1])"
    Range("D6").Select
    Selection.AutoFill Destination:=Range("D6:M6"), Type:=xlFillDefault
    Range("D6:M6").Select
    Selection.AutoFill Destination:=Range("D6:M17"), Type:=xlFillDefault
    Range("D6:M17").Select
    Range("N6").Select
    ActiveCell.FormulaR1C1 = "='3-��ҵҪ�����ݱ���д��'!R[1]C[1]"
    Range("N6").Select
    Selection.AutoFill Destination:=Range("N6:N17"), Type:=xlFillDefault
    Range("A1:N1").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[1]C[1]&""  ��ҵҪ����������۱�"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
End Sub


Sub ָ������ݱ�ʽ()
    On Error Resume Next
    Dim SchoolName As String
    Dim Major As String
    Dim i As Integer
    Dim MajorCount As Integer
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Major = Range("B7").Value
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    SchoolName = Range("B2").Value
    MajorCount = Application.WorksheetFunction.CountA(Range("B4:B" & MajorLastRow))
    Sheets("��ҵҪ��-ָ������ݱ�").Visible = True
    Worksheets("��ҵҪ��-ָ������ݱ�").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("B2:C2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[2]C"
    Range("B3:C3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R2C2="""","""",IF(OR(MID(R[-1]C,LEN(R[-1]C),1)=""A"",MID(R[-1]C,LEN(R[-1]C),1)=""B"",MID(R[-1]C,LEN(R[-1]C),1)=""C""),MID(R[-1]C,1,LEN(R[-1]C)-1),IF(ISNUMBER(FIND(""��"",R[-1]C,1)),MID(R[-1]C,1,FIND(""��"",R[-1]C,1)-1),IF(ISNUMBER(FIND(""("",R[-1]C,1)),MID(R[-1]C,1,FIND(""("",R[-1]C,1)-1),R[-1]C))))"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC5=1,IF(R[-1]C=""���"",1,R[-1]C+1),R[-1]C)"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R4C3="""",ISNA(MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R7C2,רҵ����״̬!C2,0))),"""",IF(VLOOKUP(R4C3,INDIRECT(""'""&INDEX(רҵ����״̬!C3,MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R7C2,רҵ����״̬!C2,0))&""'!$B$3:$AS$3""),ROW(RC3)-2,0)="""","""",VLOOKUP(R4C3,INDIRECT(""'""&INDEX(רҵ����״̬!C3,MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R7C2,רҵ����״̬!C2,0))&""'!$B$3:$AS$3""),ROW(RC3)-2,0)))"
    Range("E6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R2C2="""",RC[-1]="""",ISNA(MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R7C2,רҵ����״̬!C2,0))),"""",IF(ISNA(VLOOKUP(R4C[-2]&""-""&R3C2&""*"",INDIRECT(""'""&INDEX(רҵ����״̬!C3,MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R7C2,רҵ����״̬!C2,0))&""'!$A:$AS""),ROW(RC1)-1,0)),0,VLOOKUP(R4C[-2]&""-""&R3C2&""*"",INDIRECT(""'""&INDEX(רҵ����״̬!C3,MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R7C2,רҵ����״̬!C2,0))&""'!$A:$AS""),ROW(RC1)-1,0)))"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C2,RC2,C[-1])"
    Range("C4:G4").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R7C2"
    Range("A6").Select
    Selection.AutoFill Destination:=Range("A6:A46"), Type:=xlFillDefault
    Range("D6:G6").Select
    Range("D6").Activate
    Selection.AutoFill Destination:=Range("D6:G46"), Type:=xlFillDefault
    Range("D6:G46").Select
    Columns("H:S").Select
    Selection.Delete Shift:=xlToLeft
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets("��ҵҪ��-ָ������ݱ�").Visible = False
    Worksheets("רҵ����״̬").Visible = False
End Sub
Sub MergePDF(PDFFile1Name As String, PDFFile2Name As String, MergePDFFileName As String)
    On Error Resume Next
    Dim ok As Boolean
    Dim PDFApp As Acrobat.AcroApp
    Dim pddoc As Acrobat.AcroPDDoc
    Dim tempPDDoc As Acrobat.AcroPDDoc
    Set PDFApp = CreateObject("AcroExch.App")
    Set pddoc = CreateObject("AcroExch.PDDoc")
    Set tempPDDoc = CreateObject("AcroExch.PDDoc")
    If Not tempPDDoc.Open(PDFFile2Name) Then
        Set tempddoc = Nothing
        Set pddoc = Nothing
        Set PDFApp = Nothing
        Exit Sub
    End If
    ok = pddoc.Open(PDFFile1Name)
    If ok <> -1 Then
        ok = pddoc.Create()
    End If
    
    Call pddoc.InsertPages(pddoc.GetNumPages() - 1, tempPDDoc, 0, tempPDDoc.GetNumPages(), False)
    PDFPages = pddoc.GetNumPages
    ok = pddoc.Save(1, MergePDFFileName)

    Call pddoc.Close
    tempPDDoc.Close
    PDFApp.Exit
    Set pddoc = Nothing
    Set tempPDDoc = Nothing
    Set PDFApp = Nothing
End Sub
Public Sub Excel2PDF(WorkSheetName As String, PathName As String, FileName As String, isOpenAfterPublish As Boolean)
    On Error Resume Next
    Dim PDFFileName As String
    Dim WorkSheetVisble As Boolean
    WorkSheetVisble = Worksheets(WorkSheetName).Visible
    If (WorkSheetVisble = False) Then
        Worksheets(WorkSheetName).Visible = True
    End If
    Worksheets(WorkSheetName).Activate
    PDFFileName = PathName & "\" & FileName
    If Not ActiveWorkbook.Saved Then
        ThisWorkbook.Saved = True
    End If
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=PDFFileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=isOpenAfterPublish
    If (WorkSheetVisble = False) Then
        Worksheets(WorkSheetName).Visible = False
    End If
End Sub
Sub �����ļ�()
' �����ļ� ��
    On Error Resume Next
    Dim CellHeight As Integer
    Dim SumCellHeight As Integer
    Dim AverageCellHeight As Integer
    Dim CountCell As Integer
    Dim i As Integer
    Dim FileName As String
    Dim NewFileName As String
    Dim CurrentWorksheet As String
    CurrentWorksheet = ActiveSheet.Name
    
    Application.ScreenUpdating = False
    Call ����������ɫ("1-�Ծ�ɼ��ǼǱ���д��", "E2:M" & MaxLineCout, xlThemeColorDark1)
    Call ����������ɫ("3-��ҵҪ�����ݱ���д��", "E7:N18", xlThemeColorDark1)

    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    CourseName = Range("$B$4").Value
    
    Worksheets("����ʵ��ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("�ɼ��˶�").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password

    Worksheets("�ɼ�¼��").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("3-��ҵҪ�����ݱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("4-�����������棨��д+��ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
    ActiveSheet.PageSetup.CenterFooter = "1-" & CourseName & "�γ�Ŀ���ɶ����۱�"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
    ActiveSheet.PageSetup.CenterFooter = "2-" & CourseName & "��ҵҪ���ɶ����۱�"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("3-�ۺϷ�������ӡ��").Activate
    ActiveSheet.PageSetup.CenterFooter = "3 -" & CourseName & "�γ��ۺϷ�����"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
   
    Application.ScreenUpdating = False
    Call MsgInfo(NoMsgBox, "���ڽ��й������ʽ�������������ļ��������ĵȴ���")
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Call ���ý�ѧ���̵ǼǱ�
    Call �����������������ʽ
    'Application.Calculation = xlManual
    Call ��������ʽ
    'Call �ĵ���д���
    FileName = ThisWorkbook.Name

    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Range("AE7").Value = FileName
    Range("AE5").Value = ThisWorkbook.Path
    If (FileName = Range("AE6").Value) Then
        ActiveWorkbook.Save
    Else
        ActiveWorkbook.Save
        ActiveWorkbook.SaveAs FileName:=Range("AE4").Value, FileFormat:=56
    End If

    Application.ScreenUpdating = True
    Worksheets(CurrentWorksheet).Activate
End Sub
Sub ��������ʽ()
    On Error Resume Next
    '"1-�γ�Ŀ���ɶ����ۣ���ӡ��" ��2ҳ�Զ����ÿγ�Ŀ���и�
    Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = True
    Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password

    CountCell = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
    Rows("11:" & CountCell + 11).Select
    Selection.Rows.AutoFit
    Rows(CountCell + 11 & ":20").Select
    Selection.RowHeight = 20
    SumCellHeight = Rows("11:20").Height
    AverageCellHeight = (650 - SumCellHeight) / 10
    For i = 11 To 20
      Rows(i & ":" & i).Select
      Selection.RowHeight = Selection.RowHeight + AverageCellHeight
    Next i
    
    Range("A11:R20").Select
    With Selection
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("11:20").Select
    Selection.EntireRow.Hidden = False
    SumRow = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
    Rows(SumRow + 11 & ":20").Select
    Selection.EntireRow.Hidden = True
    Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
    Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = True
    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Rows("3:17").Select
    Selection.RowHeight = 25
    Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password

    '"3-�ۺϷ�������ӡ��" ��1ҳ�Զ����и�
    Worksheets("3-�ۺϷ�������ӡ��").Visible = True
    Worksheets("3-�ۺϷ�������ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Rows("3:3").Select
    Selection.Rows.AutoFit
    Rows("5:5").Select
    Selection.Rows.AutoFit
    Rows("7:7").Select
    Selection.Rows.AutoFit
    Rows("9:9").Select
    Selection.Rows.AutoFit
    Rows("11:11").Select
    Selection.Rows.AutoFit
    SumCellHeight = Rows("3:9").Height
    AverageCellHeight = (620 - SumCellHeight) / 5
    For i = 1 To 5
      Rows(2 * i + 1).Select
      Selection.RowHeight = Selection.RowHeight + AverageCellHeight
    Next i
    Worksheets("3-�ۺϷ�������ӡ��").Activate
    Worksheets("3-�ۺϷ�������ӡ��").Visible = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub �������ù�ʽ��ť()
  On Error Resume Next
  Dim SumCount As Integer
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
  SumCount = Application.WorksheetFunction.Count(Range("A6:A185"))
  ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
  Call ���ø�����ɫ
  ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
  Call �γ�Ŀ����ۺϷ�����ʽ
  Call �������ù�ʽ(SumCount)
  Call �Ծ�ɼ��ǼǱ�ʽ
  Call �ɼ�¼�빫ʽ
  Call ʵ��ɼ���ʽ
  Call ��ҵҪ�����ݱ�ʽ
  Call ָ������ݱ�ʽ
  Call ƽʱ�ɼ���ʽ
  Call �ɼ��˶Ա�ʽ
  Call �ɼ���ʽ
  Call ��ҵҪ���ɶ����۹�ʽ
  Call ���ۻ��ڱ������ù�ʽ
  Call �����������湫ʽ
  Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

Sub ���ý�ѧ���̵ǼǱ�()
    On Error Resume Next
    Dim i As Integer
    Dim Count As Integer
    Dim DataRecord As String
    Dim DataAddr As Integer
    Dim HPageBreaksCount As Integer
    Dim ErrorLog As String
    Dim CourseNumber As String
    Dim t1 As Integer
    Dim t2 As Integer
    Dim SumCount As Integer
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
    CourseNumber = Range("AN1").Value
    SumCount = Range("AM1").Value
    If (SumCount > 0) Then
        Call �������ù�ʽ(SumCount)
    ElseIf (SumCount = 0) Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    If (CourseNumber <> "") Then
    '��ȡ����
        'Count = Application.WorksheetFunction.Count(Range("B6:B185"))
        Call ���ñ����("A6", "AG" & (SumCount + 5), 9)
        Call ɾ�������(SumCount + 6, MaxRecord)
        Call ���һ�б����(SumCount + 5)
        HPageBreaksCount = 0
        ActiveSheet.ResetAllPageBreaks
        'ÿ30�м�¼������ҳ��
        Pages = Int(SumCount / 30)
        If (SumCount Mod 30 = 0) Then
            Pages = Pages - 1
        End If
        For i = 1 To Pages
          Rows(30 * i + 6 & ":" & 30 * i + 6).Select
          ActiveWindow.SelectedSheets.HPageBreaks.Add before:=ActiveCell
          HPageBreaksCount = HPageBreaksCount + 1
        Next i
        '����ʵ��ɼ��Ϳ��˳ɼ�Ϊ���Դ�С����ʽ
        Range("AC6:AD" & (SumCount + 6)).Select
        Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
        Range("A" & (SumCount + 6) & ":AF" & (MaxRecord + 3)).Select
        Selection.ClearContents
        'ȡ�����е�Ԫ������У��
        Range("D6:AF" & MaxRecord).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
            :=xlBetween
            .IgnoreBlank = True
            .InCellDropdown = True
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        Range("B6:C" & (SumCount + 5)).Select
        Selection.FormulaHidden = False
        '������ҵ�Ǽ����������У��
        Range("D6:T" & (SumCount + 5)).Select
        Selection.FormulaHidden = False
        
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$AV$6:$AV$22"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = "��ҵ����͵����ǼǴ���"
            .InputMessage = ""
            .ErrorMessage = "�밴��ѧ���̵ǼǱ����Ϸ��ɼ��ȼ�������д���������ͳٵ����Ų�������ҵ�ȼ�������ͬ��"
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With

        '���óɼ�����е�����У��
        Range("AF6:AF" & (SumCount + 5)).Select
        Selection.FormulaHidden = False
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$AY$6:$AY$11"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        '���ÿ�������е�����У��
        Range("AG6:AG" & (SumCount + 5)).Select
        Selection.FormulaHidden = False
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$AZ$6:$AZ$8"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        '����δ������ҵ�ɼ����۷�ʽ
        Range("AX6").Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$AY$4:$AY$5"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
    
        Range("AE" & SumCount + 6 & ":AE" & MaxRecord + 5).Select
        Selection.FormatConditions.Delete
        Range("A" & (SumCount + 7) & ":AG" & (SumCount + 7)).Select
        '�������һ����Ϣ��ʽ
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
    
        'д�����һ�гɼ����ɱ���������ˣ���д���ڣ����������ε���Ϣ
        Range("A" & SumCount + 7).Select
        ActiveCell.FormulaR1C1 = "=R2C34"
        Range("AJ1").Value = Count + 7
        '�����ҳ��
        Rows(SumCount + 8 & ":" & SumCount + 8).Select
        Range("A" & SumCount + 8).Activate
        ActiveWindow.SelectedSheets.HPageBreaks.Add before:=ActiveCell
    
        Range("A1:AG" & SumCount + 7).Select
        ActiveSheet.PageSetup.PrintArea = "$A$1:$AG$" & SumCount + 7
    Else
        Call ���ñ����("A6", "AG185", 9)
    End If
        '���ý�ѧ���̵ǼǱ��ͷ��ʽ
    Call ���ñ����("A4", "AG5", 9)
    Range("A4:AG5").Select
    With Selection.Font
        .Name = "����"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("A6:AG" & (SumCount + 5)).Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "����"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Call ���ø�����ɫ
    Columns("AH:AR").Select
    Range("AH2").Activate
    Selection.EntireColumn.Hidden = True
    
    Columns("AY:AZ").Select
    Selection.EntireColumn.Hidden = True
    Columns("BE:BE").Select
    Selection.EntireColumn.Hidden = True
    
    Call ���ñ����("AT6", "AW21", 9)
    Range("AT6:AW21").Select
    Selection.Font.Bold = False
    With Selection.Font
        .Name = "����"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    '����������
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub �ɼ���ʽ()
    On Error Resume Next
    Worksheets("�ɼ���").Activate
    Sheets("�ɼ���").Visible = 1
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "�����ɼ�"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "���"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "ѧ��"
    Range("R1").Select
    Selection.FormulaR1C1 = "�ı�ѧ��"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "�����ɼ�"
    Range("U1").Select
    Selection.FormulaR1C1 = "ƽʱ�ɼ�"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "�ɼ����"
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "�������"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""���"",1,R[-1]C+1)"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[-1],OFFSET(R1C1,R3C33-1,R3C34-1,601,1),0)),INDEX(OFFSET(R1C1,R4C33-1,R4C34-1,601,1),MATCH(RC[-1],OFFSET(R1C1,R3C33-1,R3C34-1,601,1),0)),IF(ISNUMBER(MATCH(RC[-1],OFFSET(R1C1,R3C33-1,R3C37-1,601,1),0)),INDEX(OFFSET(R1C1,R4C33-1,R4C37-1,601,1),MATCH(RC[-1],OFFSET(R1C1,R3C33-1,R3C37-1,601,1),0)),""""))"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]=0,"""",IF(ISNUMBER(RC[-1]),TEXT(RC[-1],""0000000000""),RC[-1]))"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),""""))"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-2]="""",INDEX(C33,MATCH(R1C,C30,0))=0),"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),""""))"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),""""))"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),""""))"
    Range("W2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0))),""""))"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0))),""""))"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0))),""""))"
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0))),""""))"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0))),""""))"
    Range("AB2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),""""))"
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""ѧ��"",C30,0))-1,INDEX(C34,MATCH(""ѧ��"",C30,0))-1,601,1),0)),""""))"
    Range("P2:AC2").Select
    Selection.AutoFill Destination:=Range("P2:AC601"), Type:=xlFillDefault
    Range("P2:AC601").Select
    
    Range("AD6:AD15").Select
    Selection.FormulaArray = "=TRANSPOSE(R[-5]C[-10]:R[-5]C[-1])"
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "ƫ��"
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "�ؼ���������"
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "��1���ؼ���������"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[1]=2,2,1)"
    Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=COUNT(R[2]C[2]:R[2]C[5])"
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(OFFSET(R[-1]C1,0,R4C34-1,200))"
    Range("AD3").Select
    ActiveCell.FormulaR1C1 = "���"
    Range("AD4").Select
    ActiveCell.FormulaR1C1 = "ѧ��"
    Range("AD5").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("AE3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[3]<>"""",OFFSET(R1C1,RC33-1,RC34-1,1,1),"""")"
    Range("AE3").Select
    Selection.AutoFill Destination:=Range("AE3:AE15"), Type:=xlFillDefault
    Range("AE3:AE15").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets("�ɼ���").Visible = 0
End Sub
Sub �γ�Ŀ����ۺϷ�����ʽ()
    Dim Evaluation1 As String
    Dim Evaluation2 As String
    On Error Resume Next
    Call �����ѧ����
    Application.ScreenUpdating = False
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Application.EnableEvents = False
    Evaluation1 = Range("D2").Value
    Range("D2:E2").Select
    ActiveCell.FormulaR1C1 = Evaluation1
    Range("F2:G2").Select
    ActiveSheet.Unprotect
    ActiveCell.FormulaR1C1 = "��ҵ�ɼ�"
    Range("H2:I2").Select
    ActiveCell.FormulaR1C1 = "ʵ��ɼ�"
    Range("J2:K2").Select
    ActiveCell.FormulaR1C1 = "���ò���"
    Evaluation2 = Range("L2").Value
    Range("L2:M2").Select
    ActiveCell.FormulaR1C1 = Evaluation2
    Range("N2:O2").Select
    ActiveCell.FormulaR1C1 = "���˳ɼ�"
    Range("F2:K2").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("N2:O2").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("L2:M2,D2:E2").Select
    Range("D2").Activate
    With Selection.Font
        .Name = "����"
        .FontStyle = "�Ӵ�"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = 255
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
        .PatternTintAndShade = 0
    End With
    Range("B7").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=OFFSET(רҵ����״̬!$B$4,,,COUNTA(רҵ����״̬!$B$4:$B$12))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("Q3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$X$3:$X$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("D4:L4").Select
    ActiveCell.FormulaR1C1 = "=IF(R5C4=""һ"",""����"",""��ʵ��ɼ�"")"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2]="""","""",'1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2])"
    Range("N5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2]="""","""",'1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2])"
    Range("O5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2]="""","""",'1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2])"
    Range("P5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2]="""","""",'1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2])"
    Range("Q5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2]="""","""",'1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[2])"

    Range("V1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[2]C[-20]="""",0,IF(ISNUMBER(MATCH(TEXT(R[2]C[-20],""0000000""),��ѧ���̵ǼǱ�!C[-21],0)),COUNTIF(��ѧ���̵ǼǱ�!C[-20],R[2]C[-20]&""*""),IF(ISNUMBER(MATCH(R[2]C[-20],��ѧ���̵ǼǱ�!C[-21],0)),COUNTIF(��ѧ���̵ǼǱ�!C[-20],R[2]C[-20]&""*""),0)))"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R3C2="""","""",IF(R4C2=""��������Դ-��ѧ��������ӸÿκŵĿγ���Ϣ"","""",MID(VLOOKUP(R3C2,��ѧ����!C1:C13,2,0),1,9)&""ѧ��� ""&MID(VLOOKUP(R3C2,��ѧ����!C1:C13,2,0),11,1)&"" ѧ��""))"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SUM(RC[-12],RC[-10],RC[-8],RC[-6],RC[-4],RC[-2])=0,"""",SUM(RC[-12],RC[-10],RC[-8],RC[-6],RC[-4],RC[-2]))"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R3C2="""","""",IF(ISERROR(VLOOKUP(R3C2,��ѧ����!R1C1:R65536C13,5,0)),""��������Դ-��ѧ��������ӸÿκŵĿγ���Ϣ"",VLOOKUP(R3C2,��ѧ����!R1C1:R65536C13,5,0)))"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R3C2="""","""",IF(R4C2=""��������Դ-��ѧ��������ӸÿκŵĿγ���Ϣ"","""",VLOOKUP(R3C2,��ѧ����!R1C1:R65536C13,6,0)))"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C="""","""",COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C24,R9C4&""-""&R7C2&""-""&""��֤""))"
    Range("B9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-6]C="""","""",IF(R4C2=""��������Դ-��ѧ��������ӸÿκŵĿγ���Ϣ"","""",VLOOKUP(R3C2,��ѧ����!R1C1:R65536C13,4,0)))"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[1]="""","""",""�γ�Ŀ��""&(ROW(RC)-10))"
    Range("A11").Select
    ActiveWindow.SmallScroll Down:=3
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",ROUND(SUM(RC[16]:RC[29]),0))"
    Range("R11").Select
    ActiveCell.FormulaR1C1 = "=IF(SUM(RC[-14]:RC[-1])=0,"""",SUM(RC[-14]:RC[-1]))"
    Range("A11:A11").Select
    Selection.AutoFill Destination:=Range("A11:A20"), Type:=xlFillDefault
    Range("A11:A20").Select
    Range("C11:C11").Select
    Selection.AutoFill Destination:=Range("C11:C20"), Type:=xlFillDefault
    Range("C11:C11").Select
    Range("R11:R11").Select
    Selection.AutoFill Destination:=Range("R11:R20"), Type:=xlFillDefault
    Range("R11:R20").Select
    Range("D6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[1]="""","""",'1-�Ծ�ɼ��ǼǱ���д��'!R[-3]C[1])"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2="""",R[-1]C="""",COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C24,R9C4&""-""&R7C2&""-""&""��֤"")=0,COUNT(OFFSET('1-�Ծ�ɼ��ǼǱ���д��'!R4C1,,MATCH(R5C,'1-�Ծ�ɼ��ǼǱ���д��'!R2,0)-1,205))=0),"""",SUMIF('1-�Ծ�ɼ��ǼǱ���д��'!C24,R9C4&""-""&R7C2&""-""&""��֤"",OFFSET('1-�Ծ�ɼ��ǼǱ���д��'!R1C1,,MATCH(R5C,'1-�Ծ�ɼ��ǼǱ���д��'!R2,0)-1,183))/R6C2)"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C="""","""",ROUND(R[-1]C*100/R[-2]C,1))"
    Range("D6:D8").Select
    Selection.AutoFill Destination:=Range("D6:L8"), Type:=xlFillDefault
    Range("D6:L8").Select
    Range("M6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(OFFSET(R3C1,,MATCH(R5C,R2,0)-1,1)="""",OFFSET(R3C1,,MATCH(R5C,R2,0)-1,1)=0),"""",100)"
    Range("M7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2="""",R[-1]C="""",COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C24,R9C4&""-""&R7C2&""-""&""��֤"")=0,COUNT(OFFSET('1-�Ծ�ɼ��ǼǱ���д��'!R4C1,,MATCH(R5C,'1-�Ծ�ɼ��ǼǱ���д��'!R2,0)-1,205))=0),"""",SUMIF('1-�Ծ�ɼ��ǼǱ���д��'!C24,R9C4&""-""&R7C2&""-""&""��֤"",OFFSET('1-�Ծ�ɼ��ǼǱ���д��'!R1C1,,MATCH(R5C,'1-�Ծ�ɼ��ǼǱ���д��'!R2,0)-1,183))/R6C2)"
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C="""","""",ROUND(R[-1]C*100/R[-2]C,1))"
    Range("M6:M8").Select
    Selection.AutoFill Destination:=Range("M6:R8"), Type:=xlFillDefault
    Range("M6:R8").Select
    Range("R7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(COUNTBLANK(RC[-14]:RC[-6])=14,SUM(R6C4:R6C12)=0),"""",SUM('2-�γ�Ŀ����ۺϷ�������д��'!R7C4:R7C12)*100/SUM(R6C4:R6C12))"
    Range("R8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C18="""",R7C18="""",R7C18=0),"""",ROUND(R7C18*100/R6C18,1))"
    Range("D9:E9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-6]C[-2]="""","""",IF(R4C2=""��������Դ-��ѧ��������ӸÿκŵĿγ���Ϣ"","""",VLOOKUP(R3C2,��ѧ����!R1C1:R65536C13,8,0)))"
    Range("H9:I9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-6]C[-6]="""","""",IF(R4C2=""��������Դ-��ѧ��������ӸÿκŵĿγ���Ϣ"","""",VLOOKUP(R3C2,��ѧ����!R1C1:R65536C13,3,0)))"
    Range("N9:R9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SUM(R[2]C[-11]:R[11]C[-11])=0,"""",ROUND(AVERAGE(R[2]C[-11]:R[11]C[-11]),0))"
    Range("B21:R21").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2=""""),"""",IF(CONCATENATE(R[-14]C[17],R[-14]C[18],R[-14]C[19],R[-14]C[20],R[-14]C[21],R[-14]C[22],R[-14]C[23],R[-14]C[24],R[-14]C[25])="""","""",""�Ծ����ƽ���÷ֱ�Ϊ��""&MID(CONCATENATE(R[-14]C[17],R[-14]C[18],R[-14]C[19],R[-14]C[20],R[-14]C[21],R[-14]C[22],R[-14]C[23],R[-14]C[24],R[-14]C[25]),1,LEN(CONCATENATE(R[-14]C[17],R[-14]C[18],R[-14]C[19],R[-14]C[2" & _
        "0],R[-14]C[21],R[-14]C[22],R[-14]C[23],R[-14]C[24],R[-14]C[25]))-1)))" & _
        ""
    Range("B22:R22").Select
    Range("B24:R24").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2=""""),"""",IF(CONCATENATE(R[-13]C[31],R[-12]C[31],R[-11]C[31],R[-10]C[31],R[-9]C[31],R[-8]C[31],R[-7]C[31],R[-6]C[31],R[-5]C[31],R[-4]C[31])="""","""",""�γ�Ŀ���ɶȷֱ�Ϊ��""&MID(CONCATENATE(R[-13]C[31],R[-12]C[31],R[-11]C[31],R[-10]C[31],R[-9]C[31],R[-8]C[31],R[-7]C[31],R[-6]C[31],R[-5]C[31],R[-4]C[31]),1,LEN(CONCATENATE(R[-13]C[31],R[-12]C[31],R[-11]C[3" & _
        "1],R[-10]C[31],R[-9]C[31],R[-8]C[31],R[-7]C[31],R[-6]C[31],R[-5]C[31],R[-4]C[31]))-1)))" & _
        ""
    Range("B26:R26").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2=""""),"""",IF(CONCATENATE(RC[18],RC[19],RC[20],RC[21],RC[22],RC[23],RC[24],RC[25],RC[26],RC[27],RC[28],RC[29])="""","""",""��ҵҪ���ɶȷֱ�Ϊ��""&MID(CONCATENATE(RC[18],RC[19],RC[20],RC[21],RC[22],RC[23],RC[24],RC[25],RC[26],RC[27],RC[28],RC[29]),1,LEN(CONCATENATE(RC[18],RC[19],RC[20],RC[21],RC[22],RC[23],RC[24],RC[25],RC[26],RC[27],RC[28],RC[29]))-1)))"
    Range("B27:R27").Select
    Range("T1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(R[2]C[-18]&""-1"",��ѧ���̵ǼǱ�!C[-18],0)),""������"",IF(MATCH(R[2]C[-18]&""-1"",��ѧ���̵ǼǱ�!C[-18],0)-6<0,0,MATCH(R[2]C[-18]&""-1"",��ѧ���̵ǼǱ�!C[-18],0)-6))"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[2]C[-20]="""",0,IF(ISNUMBER(MATCH(TEXT(R[2]C[-20],""0000000""),��ѧ���̵ǼǱ�!C[-21],0)),COUNTIF(��ѧ���̵ǼǱ�!C[-20],R[2]C[-20]&""*""),IF(ISNUMBER(MATCH(R[2]C[-20],��ѧ���̵ǼǱ�!C[-21],0)),COUNTIF(��ѧ���̵ǼǱ�!C[-20],R[2]C[-20]&""*""),0)))"
    Range("W1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]<>""������"",MATCH(R[2]C[-21]&""-1"",��ѧ���̵ǼǱ�!C[-21],0)+COUNTIF(��ѧ���̵ǼǱ�!C[-22],R[2]C[-21])-1,"""")"
    Range("AG1").Select
    Selection.FormulaR1C1 = "=IF(R3C2="""","""",VLOOKUP(R3C2,��ѧ����!C1:C13,2,0))"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "=ISNUMBER(MATCH(R[1]C[-18],��ѧ���̵ǼǱ�!C[-19],0))"
    Range("U2:W2").Select
    ActiveCell.FormulaR1C1 = "='0-��ѧ���̵ǼǱ���д+��ӡ)'!R[2]C[9]"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "='4-�����������棨��д+��ӡ��'!R[4]C[-22]"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "='4-�����������棨��д+��ӡ��'!R[3]C[-17]"
    Range("T4").Select
    Selection.FormulaR1C1 = _
        "=VLOOKUP(""����"",���ۻ��ڱ�������!R3C[-19]:R3C[-12],MATCH(""�ۺ�ƽʱ"",���ۻ��ڱ�������!R2C[-19]:R2C[-12],0),0)"
    Range("U4").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(""����"",���ۻ��ڱ�������!R3C[-20]:R3C[-13],MATCH(""ʵ��ɼ�"",���ۻ��ڱ�������!R2C[-20]:R2C[-13],0),0)"
    Range("V4").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(""����"",���ۻ��ڱ�������!R3C[-21]:R3C[-14],MATCH(""���˳ɼ�"",���ۻ��ڱ�������!R2C[-21]:R2C[-14],0),0)"
    Range("W4").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(""����"",���ۻ��ڱ�������!R3C[-22]:R3C[-15],MATCH(""���гɼ�"",���ۻ��ڱ�������!R2C[-22]:R2C[-15],0),0)"
    Range("AE4").Select
    ActiveCell.FormulaR1C1 = "=R[1]C&""\""&R[2]C"
    Range("S6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-15]="""","""",ROUND(R[1]C[-15]/RC[-15],3))"
    Range("T6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-15]="""","""",ROUND(R[1]C[-15]/RC[-15],3))"
    Range("U6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-15]="""","""",ROUND(R[1]C[-15]/RC[-15],3))"
    Range("V6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-15]="""","""",ROUND(R[1]C[-15]/RC[-15],3))"
    Range("W6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-15]="""","""",ROUND(R[1]C[-15]/RC[-15],3))"
    Range("X6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-15]="""","""",ROUND(R[1]C[-15]/RC[-15],3))"
    Range("Y6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-15]="""","""",ROUND(R[1]C[-15]/RC[-15],3))"
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-15]="""","""",ROUND(R[1]C[-15]/RC[-15],3))"
    Range("AE6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-3]C[-29]="""",R[1]C,MID(R[-4]C[-29],3,2)&""-""&MID(R[-4]C[-29],8,2)&""-""&MID(R[-4]C[-29],14,1)&""-""&R[-3]C[-29]&""-""&R[1]C[-29]&""-""&R[-1]C[-29]&""-""&R[-2]C[-29])"
    
    Range("S7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-1]C="""","""",""��""&R[-2]C[-15]&""�⣺""&TEXT(R[-1]C,""00.0%"")&""��"")"
    Range("S7").Select
    Range("X7").Select
    Range("S7").Select
    Selection.AutoFill Destination:=Range("S7:AA7"), Type:=xlFillDefault
    Range("S7:AA7").Select
    Range("S10").Select
    ActiveWindow.SmallScroll Down:=3
    ActiveCell.FormulaR1C1 = "=R[-5]C[-15]"
    Range("S10").Select
    Selection.AutoFill Destination:=Range("S10:AF10"), Type:=xlFillDefault
    Range("S10:AF10").Select
    Range("S11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(OFFSET(R8C1,,MATCH(R10C,R5C1:R5C18,0)-1)="""",OFFSET(RC1,,MATCH(R10C,R5C1:R5C18,0)-1)=""""),0,OFFSET(R8C1,,MATCH(R10C,R5C1:R5C18,0)-1)*OFFSET(RC1,,MATCH(R10C,R5C1:R5C18,0)-1)/100)"
    Range("S11").Select
    Selection.AutoFill Destination:=Range("S11:AF11"), Type:=xlFillDefault
    Range("S11:AF11").Select
    Selection.AutoFill Destination:=Range("S11:AF20"), Type:=xlFillDefault
    Range("S11:AF20").Select
    Range("T25").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(R[-3]C&""*"",'3-��ҵҪ�����ݱ���д��'!C2,0)),"""",INDEX('3-��ҵҪ�����ݱ���д��'!C4,MATCH(R[-3]C&""*"",'3-��ҵҪ�����ݱ���д��'!C2,0)))"
    Range("T25").Select
    Selection.AutoFill Destination:=Range("T25:AE25"), Type:=xlFillDefault
    Application.EnableEvents = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub �����������湫ʽ()
    On Error Resume Next
    Worksheets("4-�����������棨��д+��ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H4:K5").Select
    Selection.NumberFormatLocal = "yyyy""��""m""��""d""��"";@"
    Range("A6:P16").Select
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Range("P10,F13:P13,F16:P16").Select
    Selection.NumberFormatLocal = "0.00_ "
    Range("D8:E9,J8:K9,P8:P10").Select
    Range("P8").Activate
    Selection.NumberFormatLocal = "0%"
    Range("P10").Select
    Selection.NumberFormatLocal = "0.00%"
    Range("A2:P2").Select
    ActiveCell.FormulaR1C1 = _
        "=""������ڣ� ""&TEXT('2-�γ�Ŀ����ۺϷ�������д��'!R[6]C[1],""YYYY��MM��DD��"")&""    """
    Range("C3:E3").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[2]C[-1]"
    Range("K3:P3").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[1]C[-9]"
    Range("C4:E5").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[-1]C[-1]"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTA('1-�Ծ�ɼ��ǼǱ���д��'!R4C2:R183C2)-COUNTBLANK('1-�Ծ�ɼ��ǼǱ���д��'!R4C2:R183C2)"
    Range("F6:G6").Select
    ActiveCell.FormulaR1C1 = _
        "=RC[-4]-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[14],""����"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[14],""ȡ��"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[14],""����"")"
    Range("H6:P6").Select
    ActiveCell.FormulaR1C1 = _
        "=""�ˣ�������""&COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[12],""����"")&""��;������""&COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[13],""����"")&""�ˣ�ȡ�������ʸ�""&COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[13],""ȡ��"")&""��"""
    Range("D8:E8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R8C1,���ۻ��ڱ�������!R2,0)-1,1,1)=0,"""",OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R8C1,���ۻ��ڱ�������!R2,0)-1,1,1))"
    Range("J8:K8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R8C6,���ۻ��ڱ�������!R2,0)-1,1,1)=0,"""",OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R8C6,���ۻ��ڱ�������!R2,0)-1,1,1))"
    Range("P8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R8C12,���ۻ��ڱ�������!R2,0)-1,1,1)=0,"""",OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R8C12,���ۻ��ڱ�������!R2,0)-1,1,1))"
    Range("D9:E9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R9C1,���ۻ��ڱ�������!R2,0)-1,1,1)=0,"""",OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R9C1,���ۻ��ڱ�������!R2,0)-1,1,1))"
    Range("J9:K9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R9C6,���ۻ��ڱ�������!R2,0)-1,1,1)=0,"""",OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R9C6,���ۻ��ڱ�������!R2,0)-1,1,1))"
    Range("P9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R9C12,���ۻ��ڱ�������!R2,0)-1,1,1)=0,"""",OFFSET(���ۻ��ڱ�������!R3C1,,MATCH(R9C12,���ۻ��ڱ�������!R2,0)-1,1,1))"
    Range("F10:I10").Select
    ActiveCell.FormulaR1C1 = "=""���""&MAX('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14)&""��"""
    Range("J10:L10").Select
    ActiveCell.FormulaR1C1 = "=""���""&MIN('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14)&""��"""
    Range("F12:H12").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,"">=90"")"
    Range("I12:J12").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,"">=80"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,"">=90"")"
    Range("K12:L12").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,"">=70"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,"">=80"")"
    Range("M12:O12").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,"">=60"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,"">=70"")"
    Range("P12").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[4],""ȡ��"")=R6C2,COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,""<60"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[4],""����"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!C[4],""����""),COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C14:R400C14,""<60""))"
    Range("F15:H15").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,"">=90"")"
    Range("I15:J15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,"">=80"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,"">=90"")"
    Range("K15:L15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,"">=70"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,"">=80"")"
    Range("M15:O15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,"">=60"")-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,"">=70"")"
    Range("P15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""<60"")+COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")+COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""ȡ��"")"
  
    
    Range("A9:C9").Select
    ActiveCell.FormulaR1C1 = "='0-��ѧ���̵ǼǱ���д+��ӡ)'!R[-5]C[26]"
    Range("L9:O9").Select
    ActiveCell.FormulaR1C1 = "='0-��ѧ���̵ǼǱ���д+��ӡ)'!R[-5]C[13]"
    
    Range("P10").Select
    ActiveCell.FormulaR1C1 = "=IF(R[3]C="""","""",(100-R[3]C)/100)"
    
    Range("F13:H13").Select
    ActiveCell.FormulaR1C1 = "=IF(R6C6=0,"""",R[-1]C*100/R6C6)"
    Range("I13:J13").Select
    ActiveCell.FormulaR1C1 = "=IF(R6C6=0,"""",R[-1]C*100/R6C6)"
    Range("K13:L13").Select
    Selection.FormulaR1C1 = "=IF(R6C6=0,"""",R[-1]C*100/R6C6)"
    Range("M13:O13").Select
    ActiveCell.FormulaR1C1 = "=IF(R6C6=0,"""",R[-1]C*100/R6C6)"
    Range("P13").Select
    ActiveCell.FormulaR1C1 = "=IF(R6C6=0,"""",R[-1]C*100/R6C6)"
    Range("F16:H16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("I16:J16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("K16:L16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("M16:O16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("P16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-�Ծ�ɼ��ǼǱ���д��'!R4C20:R400C20,""����"")))"
    Range("D17:P17").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[5]C[-2]"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub �ɼ��˶Ա�ʽ()
    On Error Resume Next
    Worksheets("�ɼ��˶�").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "ѧ��"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "ƽʱ�ɼ�"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "���˳ɼ�"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "�����ɼ�"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "�ɼ����"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "�������"
    Range("P1:X1").Select
    Selection.UnMerge
    Columns("P:P").Select
    Selection.ColumnWidth = 12
    Columns("Q:Q").Select
    Selection.ColumnWidth = 8
    Columns("T:X").Select
    Selection.ColumnWidth = 6
    Range("P1:X1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(RC[-1],'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C1:R179C1,0)),"""",IF(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C2:R179C2,MATCH(RC[-1],'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C1:R179C1,0))=0,"""",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C2:R179C2,MATCH(RC[-1],'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C1:R179C1,0))))"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(RC[-2],'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C1:R179C1,0)),"""",IF(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C2:R179C2,MATCH(RC[-2],'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C1:R179C1,0))=0,"""",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C3:R179C3,MATCH(RC[-2],'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C1:R179C1,0))))"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]=""�뵼������"","""",INDEX(�ɼ�¼��!R3C[1]:R183C[1],MATCH(RC2,�ɼ�¼��!R3C1:R183C1,0)))"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]=""�뵼������"","""",INDEX(�ɼ�¼��!R3C[1]:R183C[1],MATCH(RC2,�ɼ�¼��!R3C1:R183C1,0)))"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]=""�뵼������"","""",INDEX(�ɼ�¼��!R3C[1]:R183C[1],MATCH(RC2,�ɼ�¼��!R3C1:R183C1,0)))"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-5]=""�뵼������"","""",INDEX(�ɼ�¼��!R3C[1]:R183C[1],MATCH(RC2,�ɼ�¼��!R3C1:R183C1,0)))"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]=""�뵼������"","""",INDEX(�ɼ�¼��!R3C[1]:R183C[1],MATCH(RC2,�ɼ�¼��!R3C1:R183C1,0)))"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""�뵼������""),"""",IF(COUNTBLANK(R3C16:R182C16)=180,""δ����"",IF(ISNA(MATCH(RC2,C16,0)),""�˿�"",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))))"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""�뵼������"",RC10=""δ����"",RC10=""�˿�""),"""",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""�뵼������"",RC10=""δ����"",RC10=""�˿�""),"""",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""�뵼������"",RC10=""δ����"",RC10=""�˿�""),"""",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""�뵼������"",RC10=""δ����"",RC10=""�˿�""),"""",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(�ɼ��˶�!RC2,�ɼ���!C18,0)),"""",VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(""�ı�""&R2C16,�ɼ���!R1C18:R1C29,0),0))"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(�ɼ��˶�!RC2,�ɼ���!C18,0)),"""",VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(R2C,�ɼ���!R1C18:R1C29,0),0))"
    Range("T3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(�ɼ��˶�!RC2,�ɼ���!C18,0)),"""",VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(R2C,�ɼ���!R1C18:R1C29,0),0))"
    Range("U3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(�ɼ��˶�!RC2,�ɼ���!C18,0)),"""",VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(R2C,�ɼ���!R1C18:R1C29,0),0))"
    Range("V3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(�ɼ��˶�!RC2,�ɼ���!C18,0)),"""",IF(ISNUMBER(VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(R2C,�ɼ���!R1C18:R1C29,0),0)),ROUND(VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(R2C,�ɼ���!R1C18:R1C29,0),0),0),VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(R2C,�ɼ���!R1C18:R1C29,0),0)))"
    Range("W3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(�ɼ��˶�!RC2,�ɼ���!C18,0)),"""",VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(R2C,�ɼ���!R1C18:R1C29,0),0))"
    Range("X3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(�ɼ��˶�!RC2,�ɼ���!C18,0)),"""",VLOOKUP(RC2,�ɼ���!C18:C29,MATCH(R2C,�ɼ���!R1C18:R1C29,0),0))"
    Range("B3:X3").Select
    Selection.AutoFill Destination:=Range("B3:X182"), Type:=xlFillDefault
    Range("B3").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub ƽʱ�ɼ���ʽ()
    On Error Resume Next
    Worksheets("ƽʱ�ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password

    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1,0)),"""",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1,0)))"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1,0)),"""",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1,0)))"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC=INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH(""�ѽ�δ����"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)),ISNUMBER(MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R18C48:R23C48,0)),'0-��ѧ���̵ǼǱ���д+��ӡ)'!RC=""X""),'0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,IF(ISNA(MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C48:R23C48,0)),0,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C49:R23C49,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0" & _
        "-��ѧ���̵ǼǱ���д+��ӡ)'!R6C48:R23C48,0))))" & _
        ""
    Range("D6").Select
    Selection.AutoFill Destination:=Range("D6:U6"), Type:=xlFillDefault
    Range("D6:U6").Select
    Range("V6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC),'0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,IF(ISNA(MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C48:R23C48,0)),0,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C49:R23C49,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C48:R23C48,0))))"
    Range("W6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC),'0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,IF(ISNA(MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C48:R23C48,0)),0,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C49:R23C49,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C48:R23C48,0))))"
    Range("X6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC),'0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,IF(ISNA(MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C48:R23C48,0)),0,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C49:R23C49,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C48:R23C48,0))))"
    Range("Y6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[5]:RC[6])"
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC25+RC29=0,0,ROUND((RC32)/(R5C25+R5C29),0))"
    Range("AA6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R5C28=0,RC[-25]=0),0,ROUND(SUMIF(R4C[-5]:R4C[-3],""����"",RC[-5]:RC[-3])/R5C28,0))"
    Range("AB6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(OFFSET(RC1,,R3C[6]-1,,R3C38),"">0"")"
    Columns("AC:AC").Select
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Range("AC6").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC4:RC21,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH(""������"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)))+COUNTIF('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC4:RC21,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH(""�����ٵ�"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)))+COUNTIF('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC4:RC21,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH(""�������"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)))"
    Range("AD6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(RC4:RC21,"">0"")"
    Range("AE6").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF(RC4:RC21,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH(""�ѽ�δ����"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)))"
    Range("AF6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[3])"
    Range("AG6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC4:RC21)"
    Range("AH6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C50=""Ĭ�ϳɼ�"",RC31*VLOOKUP(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH(""�ѽ�δ����"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)),'0-��ѧ���̵ǼǱ���д+��ӡ)'!C48:C49,2,0),RC31*RC36)"
    Columns("AC:AJ").Select
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Range("AI6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R5C29=0,0,COUNTIF('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC4:RC21,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!R18C46,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)))*VLOOKUP(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!R18C46,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)),'0-��ѧ���̵ǼǱ���д+��ӡ)'!C48:C49,2,0)+COUNTIF('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC4:RC21,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!R19C46,'0-��" & _
        "ѧ���̵ǼǱ���д+��ӡ)'!C46,0)))*VLOOKUP(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!R19C46,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)),'0-��ѧ���̵ǼǱ���д+��ӡ)'!C48:C49,2,0)+COUNTIF('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC4:RC21,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!R20C46,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)))*VLOOKUP(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!R20C46,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46" & _
        ",0)),'0-��ѧ���̵ǼǱ���д+��ӡ)'!C48:C49,2,0)+COUNTIF('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC4:RC21,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!R21C46,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)))*VLOOKUP(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!R21C46,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)),'0-��ѧ���̵ǼǱ���д+��ӡ)'!C48:C49,2,0))" & _
        ""
    Range("AJ6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC25=0,0,IF(RC[-6]=0,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C[13]:R19C[13],MATCH(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C48,MATCH(""�ѽ�δ����"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!C46,0)),'0-��ѧ���̵ǼǱ���д+��ӡ)'!R6C[12]:R19C[12],0)),ROUND(RC33/RC30,1)))"
    Range("B6:AJ6").Select
    
    Range("AJ6").Activate
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    Selection.AutoFill Destination:=Range("B6:AJ185"), Type:=xlFillDefault
    Range("B6:AJ185").Select
    Range("Y5").Select
    ActiveCell.FormulaR1C1 = "=MAX(R6C25:R185C25)"
    Range("AB5").Select
    ActiveCell.FormulaR1C1 = "=MAX(R6C28:R185C28)"
    Range("AC5").Select
    ActiveCell.FormulaR1C1 = "=MAX(R6C:R185C)"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub

Sub �Ծ�ɼ��ǼǱ�ʽ()
On Error Resume Next
    Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("AB3").Select
    ActiveCell.FormulaR1C1 = "�޳�������ѧ��"
    Columns("AB:AB").Select
    Selection.EntireColumn.Hidden = True
    Range("Z2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$AB$2:$AB$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
        .PatternTintAndShade = 0
    End With
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "ѧ��"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "='0-��ѧ���̵ǼǱ���д+��ӡ)'!R[2]C[16]"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[-11]"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[-10]"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[-9]"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[-8]"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[-7]"
    If (Range("T2").Value <> "�����ɼ�") Then
        Columns("T:T").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    If (Range("U2").Value <> "�ɼ����") Then
        Columns("U:U").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("T2").Select
        ActiveCell.FormulaR1C1 = "�����ɼ�"
        Range("U2").Select
        ActiveCell.FormulaR1C1 = "�ɼ����"
        Range("V2").Select
        ActiveCell.FormulaR1C1 = "רҵ"
        Range("W2").Select
        ActiveCell.FormulaR1C1 = "�꼶"
        Range("X2").Select
        ActiveCell.FormulaR1C1 = "�꼶-רҵ-��֤״̬"
        Range("Y2").Select
        ActiveCell.FormulaR1C1 = "��֤״̬"
    End If
    Range("R1:X1").Select
    Selection.UnMerge
    Columns("R:U").Select
    Selection.ColumnWidth = 4.5
    Columns("V:V").Select
    Selection.ColumnWidth = 16.81
    Columns("W:W").Select
    Selection.ColumnWidth = 6
    Columns("X:X").Select
    Selection.ColumnWidth = 23.31
    Columns("Y:Y").Select
    Selection.ColumnWidth = 9.06
    Range("R1:X1").Select
    Selection.Merge
    Columns("B:B").Select
    Selection.ColumnWidth = 10
    Columns("C:C").Select
    Selection.ColumnWidth = 8
    Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Call �Ծ�ɼ��ǼǱ���Ĺ�ʽ
    
    If Sheets("ѧ������") Is Nothing Then
        Worksheets("ʹ�ð���").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "ѧ������"
    End If
    Worksheets("ѧ������").Activate
    If (Application.WorksheetFunction.CountIf(Range("E:E"), Grade) < 1) Then
        Call ����ѧ������
    End If

End Sub
Sub �Ծ�ɼ��ǼǱ���Ĺ�ʽ()
On Error Resume Next
    Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("B4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C2="""","""",IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(VLOOKUP(RC1,�ɼ���!C16:C19,MATCH(""�ı�""&R2C2,�ɼ���!R1C16:R1C19,0),0)="""",IF(ISNA(VLOOKUP(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1:C3,MATCH(R2C2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0)),"""",IF(VLOOKUP(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1:C3,MATCH(R2C2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0)=""�뵼������"","""",VLOOKUP(RC1,'0-��ѧ���̵ǼǱ�" & _
        "����д+��ӡ)'!C1:C3,MATCH(R2C2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0))),VLOOKUP(RC1,�ɼ���!C16:C19,MATCH(""�ı�""&R2C2,�ɼ���!R1C16:R1C19,0),0)),IF(ISNA(VLOOKUP(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1:C3,MATCH(R2C2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0)),"""",VLOOKUP(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1:C3,MATCH(R2C2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0))))" & _
        ""
    Range("C4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C2="""","""",IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(VLOOKUP(RC1,�ɼ���!C16:C19,MATCH(""�ı�""&R2C2,�ɼ���!R1C16:R1C19,0),0)="""",IF(ISNA(VLOOKUP(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1:C3,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0)),"""",VLOOKUP(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1:C3,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0)),VLOOKUP(RC2,�ɼ���!C18:C19,MATCH(R2C3,�ɼ���!R1" & _
        "C18:R1C19,0),0)),IF(ISNA(VLOOKUP(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1:C3,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0)),"""",VLOOKUP(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C1:C3,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C1:R4C3,0),0))))" & _
        ""
    Range("N4").Select
    ActiveCell.FormulaR1C1 = _
                "=IF(OR(RC2="""",COUNT(RC[-9]:RC[-1])=0),"""",SUM(RC[-9]:RC[-1]))"

    Range("O4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(OR(RC[-13]="""",ISNA(VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0)))"
    Range("P4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(OR(RC[-14]="""",ISNA(VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0)))"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(OR(RC[-15]="""",ISNA(VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0)))"
    Range("R4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(OR(RC[-16]="""",ISNA(VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C&""ƽ����"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0)))"
    Range("S4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(OR(RC[-17]="""",ISNA(VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),IF(OR(RC2="""",R3C19=""""),"""",VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C&""�÷�"",'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0)/INDEX(���ۻ��ڱ�������!R3,MATCH(R2C19,���ۻ��ڱ�������!R2,0))))"
    Range("T4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(ISNA(VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),"""",VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0)))"
    Range("U4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(ISNA(VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),"""",VLOOKUP(RC2,�ɼ���!C18:C27,MATCH(R2C,�ɼ���!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",IF(VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0)=0,"""",VLOOKUP(RC2,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2:C44,MATCH(R2C,'0-��ѧ���̵ǼǱ���д+��ӡ)'!R4C2:R4C44,0),0))))"
    Range("V4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-20]=""""),"""",IF(ISNA(MATCH(RC2&""-""&'2-�γ�Ŀ����ۺϷ�������д��'!R7C2,ѧ������!C7,0)),IF(ISNUMBER(MATCH(VALUE(RC2),ѧ������!C2,0)),VLOOKUP(VALUE(RC2),ѧ������!C2:C6,5,0),IF(ISNUMBER(MATCH(RC2,ѧ������!C2,0)),VLOOKUP(RC2,ѧ������!C2:C6,5,0),"""")),INDEX(ѧ������!C6,MATCH(RC2&""-""&'2-�γ�Ŀ����ۺϷ�������д��'!R7C2,ѧ������!C7,0))))"
    Range("W4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-21]=""""),"""",IF(ISNA(VLOOKUP(VALUE(RC[-21]),ѧ������!C2:C6,4,0)),IF(ISNA(VLOOKUP(RC[-21],ѧ������!C2:C6,4,0)),"""",VLOOKUP(RC[-21],ѧ������!C2:C6,4,0)),VLOOKUP(VALUE(RC[-21]),ѧ������!C2:C6,4,0)))"
    Range("X4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-22]="""",RC[1]=""""),"""",RC[-1]&""-""&RC[-2]&""-""&RC[1])"
    Range("Y4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(AND(RC[-3]<>"""",R2C26=""""),AND(RC[-3]<>"""",R2C26=""�޳�������ѧ��"",RC[-5]>=60)),IF(ISNUMBER(RC[-2]),IF(AND(RC[-2]='2-�γ�Ŀ����ۺϷ�������д��'!R9C4,RC[-3]='2-�γ�Ŀ����ۺϷ�������д��'!R7C2),""��֤"",""""),IF(AND(VALUE(RC[-2])='2-�γ�Ŀ����ۺϷ�������д��'!R9C4,RC[-3]='2-�γ�Ŀ����ۺϷ�������д��'!R7C2),""��֤"","""")),"""")"
    
    Range("A4:C4").Select
    Selection.AutoFill Destination:=Range("A4:C" & MaxLineCout), Type:=xlFillDefault
    Range("A4:C" & MaxLineCout).Select
    Range("N4:Y4").Select
    Range("W4").Activate
    Selection.AutoFill Destination:=Range("N4:Y" & MaxLineCout), Type:=xlFillDefault
    Range("N4:Y" & MaxLineCout).Select
    Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub �ɼ�¼�빫ʽ()
On Error Resume Next
    Worksheets("�ɼ�¼��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=MID('2-�γ�Ŀ����ۺϷ�������д��'!R[1]C[1],1,9)&""-""&MID('2-�γ�Ŀ����ۺϷ�������д��'!R[1]C[1],14,1)&""ѧ�� ��ĩ�ɼ�"""
    Range("A2:H2").Select
    ActiveCell.FormulaR1C1 = _
        "='2-�γ�Ŀ����ۺϷ�������д��'!R[2]C[1]&""���ɼ�����=""&CONCATENATE(���ۻ��ڱ�������!R12C2,���ۻ��ڱ�������!R12C3,���ۻ��ڱ�������!R12C4,���ۻ��ڱ�������!R12C5)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=""�κţ�""&'2-�γ�Ŀ����ۺϷ�������д��'!R[1]C[-7]"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('0-��ѧ���̵ǼǱ���д+��ӡ)'!R[2]C[1]<>"""",'0-��ѧ���̵ǼǱ���д+��ӡ)'!R[2]C[1],"""")"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",'0-��ѧ���̵ǼǱ���д+��ӡ)'!R[2]C[1])"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]="""","""",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C39,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0)))"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-5]="""","""",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C30,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0)))"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC1="""","""",IF(OR(INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C31,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0))=""ȡ��"",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C31,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0))=""����"",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C31,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0))=""�˿�"",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C31,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0))=""����""),0,INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C31,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0))))"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC1="""","""",IF('0-��ѧ���̵ǼǱ���д+��ӡ)'!R[2]C[24]="""",""����"",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C32,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0))))"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC1="""","""",IF('0-��ѧ���̵ǼǱ���д+��ӡ)'!R[2]C[24]="""",""����"",INDEX('0-��ѧ���̵ǼǱ���д+��ӡ)'!C33,MATCH(RC1,'0-��ѧ���̵ǼǱ���д+��ӡ)'!C2,0))))"
    Range("A4:I4").Select
    Range("I4").Activate
    Selection.AutoFill Destination:=Range("A4:I" & MaxLineCout), Type:=xlFillDefault
    Range("A4:I" & MaxLineCout).Select
    Columns("J:T").Select
    Selection.ClearContents
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub ʵ��ɼ���ʽ()
On Error Resume Next
    Worksheets("����ʵ��ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("U3").Select
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "���"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "ѧ��"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "�ı�ѧ��"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "ʵ��ɼ�"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "ѧ��"
    Range("U4").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("U5").Select
    ActiveCell.FormulaR1C1 = "ʵ��ɼ�"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""���"",1,R[-1]C+1)"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(HLOOKUP(INDEX(C22,MATCH(R1C,C21,0)),OFFSET(R1C1,R3C24-1,R3C25-1,601,7),RC16+R2C23,FALSE)=0,IF(R2C23=1,"""",IF(HLOOKUP(INDEX(C26,MATCH(R1C,C21,0)),OFFSET(R1C1,R3C24-1,R3C28-1,601,7),RC16+R2C23-R2C25,FALSE)=0,"""",HLOOKUP(INDEX(C26,MATCH(R1C,C21,0)),OFFSET(R1C1,R3C24-1,R3C28-1,601,7),RC16+R2C23-R2C25,FALSE))),HLOOKUP(INDEX(C22,MATCH(R1C,C21,0)),OFFSET(R1C1,R3C24-1" & _
        ",R3C25-1,601,7),RC16+R2C23,FALSE))" & _
        ""
    Range("R2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(RC[-1]),TEXT(RC[-1],""0000000000""),RC[-1])"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[-2],OFFSET(R1C1,R3C24-1,R3C25-1,601,1),0)),INDEX(OFFSET(R1C1,R3C24-1,R4C25-1,601,1),MATCH(RC[-2],OFFSET(R1C1,R3C24-1,R3C25-1,601,1),0)),IF(ISNUMBER(MATCH(RC[-2],OFFSET(R1C1,R3C24-1,R3C28-1,601,1),0)),INDEX(OFFSET(R1C1,R3C24-1,R4C28-1,601,1),MATCH(RC[-2],OFFSET(R1C1,R3C24-1,R3C28-1,601,1),0)),""""))"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[-3],OFFSET(R1C1,R3C24-1,R3C25-1,601,1),0)),INDEX(OFFSET(R1C1,R3C24-1,R5C25-1,601,1),MATCH(RC[-3],OFFSET(R1C1,R3C24-1,R3C25-1,601,1),0)),IF(ISNUMBER(MATCH(RC[-3],OFFSET(R1C1,R3C24-1,R3C28-1,601,1),0)),INDEX(OFFSET(R1C1,R3C24-1,R5C28-1,601,1),MATCH(RC[-3],OFFSET(R1C1,R3C24-1,R3C28-1,601,1),0)),""""))"
    Range("P2:T2").Select
    Selection.AutoFill Destination:=Range("P2:T601"), Type:=xlFillDefault
    Range("P2:T601").Select
    Range("V3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[3]<>"""",OFFSET(R1C1,RC24-1,RC25-1,1,1),"""")"
    Range("V3").Select
    Selection.AutoFill Destination:=Range("V3:V5"), Type:=xlFillDefault
    Range("V3:V5").Select
    Range("Z3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[2]<>"""",OFFSET(R1C1,RC24-1,RC28-1,1,1),"""")"
    Range("Z3").Select
    Selection.AutoFill Destination:=Range("Z3:Z5"), Type:=xlFillDefault
    Range("Z3:Z5").Select
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[6]=2,2,1)"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=COUNT(R[1]C[2]:R[1]C[5])"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(OFFSET(RC1,0,R3C25-1,200))"
    Range("AB2").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(OFFSET(RC1,0,R3C28-1,200))"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub

Sub ���ۻ��ڱ������ù�ʽ()
On Error Resume Next
    Worksheets("���ۻ��ڱ�������").Activate
    Sheets("���ۻ��ڱ�������").Visible = 1
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-8]C2:R3C7,MATCH(R[-1]C&""*"",R2C2:R2C7,0))"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-8]C2:R3C7,MATCH(R[-1]C&""*"",R2C2:R2C7,0))"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "���ۻ���"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[4]"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[5]"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[5]"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[6]"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!RC[7]"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "�ۺ�ƽʱ"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "ȡ�������ʸ�����ƽʱ�ɼ�"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-�γ�Ŀ����ۺϷ�������д��'!R3C1,,MATCH(R2C,'2-�γ�Ŀ����ۺϷ�������д��'!R2,0)-1,1,1)/100"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-�γ�Ŀ����ۺϷ�������д��'!R3C1,,MATCH(R2C,'2-�γ�Ŀ����ۺϷ�������д��'!R2,0)-1,1,1)/100"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-�γ�Ŀ����ۺϷ�������д��'!R3C1,,MATCH(R2C,'2-�γ�Ŀ����ۺϷ�������д��'!R2,0)-1,1,1)/100"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-�γ�Ŀ����ۺϷ�������д��'!R3C1,,MATCH(R2C,'2-�γ�Ŀ����ۺϷ�������д��'!R2,0)-1,1,1)/100"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-�γ�Ŀ����ۺϷ�������д��'!R3C1,,MATCH(R2C,'2-�γ�Ŀ����ۺϷ�������д��'!R2,0)-1,1,1)/100"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-�γ�Ŀ����ۺϷ�������д��'!R3C1,,MATCH(R2C,'2-�γ�Ŀ����ۺϷ�������д��'!R2,0)-1,1,1)/100"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-1]C[-4]<>""���гɼ�"",INDEX(RC[-6]:RC[-1],MATCH(R2C2,R[-1]C[-6]:R[-1]C[-1],0))+INDEX(RC[-6]:RC[-1],MATCH(R2C4,R[-1]C[-6]:R[-1]C[-1],0))+INDEX(RC[-6]:RC[-1],MATCH(R2C5,R[-1]C[-6]:R[-1]C[-1],0))+INDEX(RC[-6]:RC[-1],MATCH(R2C6,R[-1]C[-6]:R[-1]C[-1],0)),INDEX(RC[-6]:RC[-1],MATCH(R2C2,R[-1]C[-6]:R[-1]C[-1],0))+INDEX(RC[-6]:RC[-1],MATCH(R2C5,R[-1]C[-6]:R[-1]C[-1],0))+I" & _
        "NDEX(RC[-6]:RC[-1],MATCH(R2C6,R[-1]C[-6]:R[-1]C[-1],0)))" & _
        ""
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "30"
    Range("B7:E7").Select
    ActiveCell.FormulaR1C1 = "��ѧ���̵ǼǱ����һ�гɼ�������ǩ����"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "���˳ɼ�"
    Range("B8:N8").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(R[-4]C,R[-3]C,R[-3]C[3],R[-3]C[4],R[-3]C[1],R[-3]C[2],R[-3]C[5],R[-2]C)"
    Range("B9:E9").Select
    ActiveCell.FormulaR1C1 = "�ɼ�¼����ͷ�ɼ�����"
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "ʵ��"
    Range("B11").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-8]C2:R3C7,MATCH(R[-1]C&""*"",R2C2:R2C7,0))"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-8]C2:R3C7,MATCH(R[-1]C&""*"",R2C2:R2C7,0))"
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "ƽʱ"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "=���ۻ��ڱ�������!R[-8]C[4]*100"
    Range("E10").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "=���ۻ��ڱ�������!R[-8]C[2]*100"
    Range("B12").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=0,"""",R[-2]C&""*""&R[-1]C&""%+"")"
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=0,"""",R[-2]C&""*""&R[-1]C&""%+"")"
    Range("D12").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=0,"""",R[-2]C&""*""&R[-1]C&""%+"")"
    Range("E12").Select
    ActiveCell.FormulaR1C1 = "=R[-2]C&""*""&R[-1]C&""%��"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets("���ۻ��ڱ�������").Visible = 0
End Sub
Sub �����������������ʽ()
On Error Resume Next
    Dim i As Integer
    Dim Count As Integer
    Dim DataRecord As String
    Dim DataAddr As Integer
    Dim LenCell As Integer
    Dim LenD18 As Integer
    Dim LenD19 As Integer
    Dim LenD20 As Integer
    Dim LenArray1(0 To 3) As Integer
    Dim LenArray2(0 To 3) As Integer
    Dim ZihaoArray(0 To 3) As Double
    LenArray1(0) = 250
    LenArray1(1) = 300
    LenArray1(2) = 320
    LenArray1(3) = 380
    LenArray2(0) = 100
    LenArray2(1) = 150
    LenArray2(2) = 160
    LenArray2(3) = 210
    ZihaoArray(0) = 10.5
    ZihaoArray(1) = 10
    ZihaoArray(2) = 9.5
    ZihaoArray(3) = 9
    
    Application.ScreenUpdating = False
    Worksheets("4-�����������棨��д+��ӡ��").Activate
     'ȡ������
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Rows("3:16").Select
    Range("P16").Activate
    Selection.RowHeight = 22
    i = 0
    Do
        Call �����ֺ�("D17:P17", ZihaoArray(i))
        LenCell = Len(Range("D17").Value)
        If (i < 4) Then
            i = i + 1
        End If
        If (LenCell <= LenArray2(i)) Or (i = 3) Then
            Exit Do
        End If
    Loop
    i = 0
    Do
        Call �����ֺ�("D18:P18", ZihaoArray(i))
        LenCell = Len(Range("D18").Value)
        If (i < 4) Then
            i = i + 1
        End If
        If (LenCell <= LenArray2(i)) Or (i = 3) Then
            Exit Do
        End If
    Loop
    i = 0
    Do
        Call �����ֺ�("D19:P19", ZihaoArray(i))
        LenCell = Len(Range("D19").Value)
        If (i < 4) Then
            i = i + 1
        End If
        If (LenCell <= LenArray2(i)) Or (i = 3) Then
            Exit Do
        End If
    Loop
    i = 0
    Do
        Call �����ֺ�("D20:P20", ZihaoArray(i))
        LenCell = Len(Range("D20").Value)
        If (i < 4) Then
            i = i + 1
        End If
        If (LenCell <= LenArray2(i)) Or (i = 3) Then
            Exit Do
        End If
    Loop
    
    '����������
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.ScreenUpdating = True
End Sub
Sub �����ֺ�(Cell As String, Zihao As Double)
  Range(Cell).Select
    With Selection.Font
        .Name = "����"
        .Size = Zihao
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlJustify
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
End Sub
Sub ȡ��������ɫ()
Dim SumCount As Integer
    SumCount = Application.WorksheetFunction.Count(Range("A6:A185")) + 5
    Range("A6:AG185").FormatConditions.Delete
End Sub
Sub ���ø�����ɫ()
Dim SumCount As Integer
    SumCount = Application.WorksheetFunction.Count(Range("A6:A185")) + 5
    Range("A6:AG185").FormatConditions.Delete
    
    Cells.FormatConditions.Delete
    Range("A6:AG" & SumCount).Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Bold = False
        .Italic = False
    End With
  
    Range("AE6:AE" & SumCount).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD($A6,2)=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("AE6:AE" & SumCount).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AND(MOD($A6,2)=1,OR($AE6=""ȡ��"",$AE6=""����"",$AE6<60))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("AE6:AE" & SumCount).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AND(MOD($A6,2)=0,OR($AE6=""ȡ��"",$AE6=""����"",$AE6<60))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    Range("A6:AD" & SumCount).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD($A6,2)=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Range("AF6:AG" & SumCount).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD($A6,2)=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
End Sub

Sub ���ñ����(StartCell As String, EndCell As String, FontSize As Integer)
    Range(StartCell & ":" & EndCell).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "����"
        .Size = FontSize
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
End Sub
Sub ɾ�������(StartRow As String, EndRow As String)

   'ɾ�������
    Rows(StartRow & ":" & EndRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub ���һ�б����(EndRow As String)
    '�������һ�м�¼���±߸���
    Range("A" & EndRow & ":AG" & EndRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub
Sub ������������������д������ɫ()
    Worksheets("4-�����������棨��д+��ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H4:K5,D18:P18,D17:P17").Select
    Range("D17").Activate
    Range("H4:K5,D18:P18,D17:P17,D19:P19,D20:P20").Select
    Range("D20").Activate
    ActiveWindow.SmallScroll Down:=3
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub ȡ����������������д������ɫ()
    Worksheets("4-�����������棨��д+��ӡ��").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H4:K5,D17:P17").Select
    Range("D17").Activate
    Range("H4:K5,D17:P17,D18:P18").Select
    Range("D18").Activate
    ActiveWindow.SmallScroll Down:=6
    Range("H4:K5,D17:P17,D18:P18,D19:P19,D20:P20").Select
    Range("D20").Activate
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub �������ù�ʽ(EndCell As Integer)
    Dim Count As Integer
    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
     'ȡ������
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A2:AG2").Select
    ActiveCell.FormulaR1C1 = _
        "=רҵ����״̬!RC[1]&"" ""&'2-�γ�Ŀ����ۺϷ�������д��'!R[2]C[1]&"" �γ�(����/����/ѡ��)��ѧ���̵ǼǱ�"""
    Range("AP4:AP5").Select
    ActiveCell.FormulaR1C1 = "=RC[-17]&""�÷�"""
    Range("B4:B5").Select
    ActiveCell.FormulaR1C1 = "ѧ��"
    Range("C4:C5").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("AA4:AA5").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[-2]C[-23]"
    Range("Y4:Y5").Select
    ActiveCell.FormulaR1C1 = "='2-�γ�Ŀ����ۺϷ�������д��'!R[-2]C[-13]"
    Range("Y4:AA5").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Range("A6:AG" & EndCell + 5).Select
    Selection.UnMerge
    Range("A6:C" & EndCell + 5).Select
    Selection.ClearContents
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "2"
    '������ѧ���̵ǼǱ�������������ɼ���
    Range("AE4:AE5").Select
    ActiveCell.FormulaR1C1 = "�����ɼ�"
    'Count = Application.WorksheetFunction.Count(Range("A6:A300"))
    'ѧ��
    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(ISNA(MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R3C2&""-""&RC1,��ѧ���̵ǼǱ�!C2,0)),""�뵼������"",TEXT(INDEX(��ѧ���̵ǼǱ�!C[2],MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R3C2&""-""&RC1,��ѧ���̵ǼǱ�!C2,0)),""0000000000"")))"
    '����
    Range("C6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]="""",ISNA(MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R3C2&""-""&RC1,��ѧ���̵ǼǱ�!C2,0))),"""",INDEX(��ѧ���̵ǼǱ�!C[2],MATCH('2-�γ�Ŀ����ۺϷ�������д��'!R3C2&""-""&RC1,��ѧ���̵ǼǱ�!C2,0)))"
    '��ҵ�ɼ�
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC32=""�˿�"","""",RC44)"
    
    'ƽʱ�ɼ�
    Range("AB6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(ISNA(VLOOKUP(RC[-26],�ɼ���!C[-9]:C[-4],3,0)),"""",VLOOKUP(RC[-26],�ɼ���!C[-9]:C[-4],3,0)),IF(RC32=""�˿�"","""",RC[11]))"
    'ʵ��ɼ�
    Range("AC6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC32=""�˿�"",RC2="""",INDEX(���ۻ��ڱ�������!R3,MATCH(R4C,���ۻ��ڱ�������!R2,0))=0),"""",IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(ISNA(VLOOKUP(RC[-27],�ɼ���!C[-10]:C[-5],4,0)),"""",VLOOKUP(RC[-27],�ɼ���!C[-10]:C[-5],4,0)),IF(ISNA(VLOOKUP(RC2,����ʵ��ɼ���!C18:C20,MATCH(""ʵ��ɼ�"",����ʵ��ɼ���!R1C18:R1C20,0),0)),""ȱ"",VLOOKUP(RC2,����ʵ��ɼ���!C18:C20,MATCH(""ʵ��ɼ�"",����ʵ��ɼ���!R1C18:R1C20,0),0))))"
    '���˳ɼ�
    Range("AD6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(OR(RC32=""�˿�""),"""",IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(ISNA(VLOOKUP(RC[-28],�ɼ���!C[-11]:C[-6],5,0)),"""",VLOOKUP(RC[-28],�ɼ���!C[-11]:C[-6],5,0)),IF(ISNA(MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC2,'1-�Ծ�ɼ��ǼǱ���д��'!C2,0)),"""",INDEX('1-�Ծ�ɼ��ǼǱ���д��'!C14,MATCH('0-��ѧ���̵ǼǱ���д+��ӡ)'!RC2,'1-�Ծ�ɼ��ǼǱ���д��'!C2,0))))))"
    '�����ɼ�
    Range("AE6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF('2-�γ�Ŀ����ۺϷ�������д��'!R3C17=""��֤���ύ�ɼ�"",IF(ISNA(VLOOKUP(RC[-29],�ɼ���!C[-12]:C[-7],6,0)),"""",VLOOKUP(RC[-29],�ɼ���!C[-12]:C[-7],6,0)),IF(RC34="""",IF(RC[2]=""����"",RC[2],ROUND(SUM(RC[4]:RC[7]),0)),RC34)))"
    '���
    Range("AH6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(RC32<>"""",RC32,IF(RC[-1]=""����"","""",IF(OR(RC29=""����"",RC29=""ȡ��"",RC29=""ȱ"",RC26=""ȡ��""),""ȡ��"",""""))))"    '���гɼ�
    Range("AI6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",RC[-8]*INDEX(���ۻ��ڱ�������!R3,MATCH(R4C27,���ۻ��ڱ�������!R2,0)))"
    'ʵ��ɼ�
    Range("AJ6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(OR(RC[-7]="""",RC[-7]=""ȱ"",RC[-7]=""����"",RC[-7]=""ȡ��""),0,RC[-7]*INDEX(���ۻ��ڱ�������!R3,MATCH(R4C36,���ۻ��ڱ�������!R2,0))))"
    '���˳ɼ�
    Range("AK6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(RC[-7]="""",0,RC[-7]*INDEX(���ۻ��ڱ�������!R3,MATCH(R4C37,���ۻ��ڱ�������!R2,0))))"
    '�ۺ�ƽʱ
    Range("AL6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[1]="""","""",RC[1]*INDEX(���ۻ��ڱ�������!R3,MATCH(R4C38,���ۻ��ڱ�������!R2,0)))"
    '�ۺ�ƽʱ�ɼ�
    Range("AM6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2=FALSE,���ۻ��ڱ�������!R3C8=0),"""",IF(OR(RC32=""ȡ��""),0,IF(R4C27=""���гɼ�"",ROUND(SUM(RC[2]:RC[4])/INDEX(���ۻ��ڱ�������!R3,MATCH(R4C38,���ۻ��ڱ�������!R2,0)),0),ROUND((SUM(RC[2]:RC[4])+RC[-12]*INDEX(���ۻ��ڱ�������!R3,MATCH(R4C27,���ۻ��ڱ�������!R2,0)))/INDEX(���ۻ��ڱ�������!R3,MATCH(R4C38,���ۻ��ڱ�������!R2,0)),0))))"
    '���ò���ƽ����
    Range("AN6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(OR(RC32=""�˿�"",RC32=""ȡ��""),0,INDEX(ƽʱ�ɼ���!C27,MATCH(RC2,ƽʱ�ɼ���!C2,0))))"
    '���ò���
    Range("AO6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]="""","""",RC[-1]*INDEX(���ۻ��ڱ�������!R3,MATCH(R4C41,���ۻ��ڱ�������!R2,0)))"
    '�γ̱���
    Range("AP6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(RC[-17]="""",0,IF(ISNUMBER(RC[-17]),RC[-17]*INDEX(���ۻ��ڱ�������!R3,MATCH(MID(R4C42,1,4),���ۻ��ڱ�������!R2,0)),INDEX(C[7],MATCH(RC[-17],C48,0))*INDEX(���ۻ��ڱ�������!R3,MATCH(MID(R4C42,1,4),���ۻ��ڱ�������!R2,0)))))"
    '��ҵ�ɼ�
    Range("AQ6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[1]="""","""",IF(RC[1]=""ȡ��"",0,RC[1]*INDEX(���ۻ��ڱ�������!R3,MATCH(R4C43,���ۻ��ڱ�������!R2,0))))"
    '��ҵ�ɼ�ƽ����
    Range("AR6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""�뵼������""),"""",IF(OR(RC32=""�˿�"",RC32=""ȡ��""),0,IF(VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0)=0,""ȡ��"",IF('2-�γ�Ŀ����ۺϷ�������д��'!R5C[-42]<>""��ʿ��"",VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0),IF(AND(RC[-23]<>"""",VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0)<>""ȡ��""),IF(VL" & _
        "OOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0)+RC[-23]/0.2>100,100,VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0)+RC[-23]/0.2),VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0))))))" & _
        ""
    Range("AH6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""�뵼������""),"""",IF(RC32<>"""",RC32,IF(RC[-1]=""����"","""",IF(OR(RC29=""����"",RC29=""ȡ��"",RC29=""ȱ"",RC26=""ȡ��""),""ȡ��"",""""))))"
    Range("AR6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(OR(RC32=""�˿�"",RC32=""ȡ��""),0,IF(VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0)=0,""ȡ��"",IF('2-�γ�Ŀ����ۺϷ�������д��'!R5C[-42]<>""��ʿ��"",VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0),IF(AND(RC[-23]<>"""",VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0)<>""ȡ��""),IF(VLOOKUP(RC[-42],ƽʱ�ɼ�" & _
        "��!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0)+RC[-23]/0.2>100,100,VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0)+RC[-23]/0.2),VLOOKUP(RC[-42],ƽʱ�ɼ���!C[-42]:C[-18],MATCH(R4C44,ƽʱ�ɼ���!R4C2:R4C26,0),0))))))" & _
        ""
    Range("B6:C6").Select
    Selection.AutoFill Destination:=Range("B6:C7"), Type:=xlFillDefault
    Range("Z6:Z6").Select
    Selection.AutoFill Destination:=Range("Z6:Z" & EndCell + 5), Type:=xlFillDefault
    Range("A6:C6").Select
    Selection.AutoFill Destination:=Range("A6:C" & EndCell + 5), Type:=xlFillDefault
    Range("AB6:AE6").Select
    Selection.AutoFill Destination:=Range("AB6:AE" & EndCell + 5), Type:=xlFillDefault
    Range("AH6:AR6").Select
    Selection.AutoFill Destination:=Range("AH6:AR" & EndCell + 5), Type:=xlFillDefault
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub �����ѧ����()
    Dim FileName As Variant
     '���ļ��Ի��򷵻ص��ļ�������һ��ȫ·���ļ�������ֵҲ������False���������ΪVariant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim temp As String
    Dim JuzhenSheets As String
    Dim strCurPath As String    ' Ӧ�ó���ĵ�ǰ·��.
    Dim strWbkPath As String    ' ���������ڵ�·��.
    Application.ScreenUpdating = False
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    
    ThisWorkBookName = ThisWorkbook.Name
    FileName = ThisWorkbook.Path
    SourceWorkBook = "����Դ-��ѧ����.xls"
    FileName = FileName & "\" & SourceWorkBook
    If MyFile.FileExists(FileName) = False Then
      Call MsgInfo(NoMsgBox, "����Դ-��ѧ����.xls�����ڣ��뽫����Դ-��ѧ����.xls���Ƶ���ǰ�ļ���")
      FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
    '����Windows���ļ��Ի���
    If (FileName = False) Then Exit Sub
    SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    On Error Resume Next
    If Sheets("��ѧ����") Is Nothing Then
        Worksheets("��ҵҪ��-ָ������ݱ�").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "��ѧ����"
    End If
    Worksheets("��ѧ����").Activate
    Worksheets("��ѧ����").Visible = True
    Call CopySheet(FileName, SourceWorkBook, "��ѧ����", "A:M", ThisWorkBookName, "��ѧ����", "A:M")
    Application.ScreenUpdating = False
End Sub 'ѡ����ļ���û����ʵ�İ�����
Sub �����ҵҪ�����(Major As String)
    Dim ThisWorksheet As String
    Dim FileName As Variant
     '���ļ��Ի��򷵻ص��ļ�������һ��ȫ·���ļ�������ֵҲ������False���������ΪVariant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim temp As String
    Dim MatrixSheet As String
    Dim strCurPath As String    ' Ӧ�ó���ĵ�ǰ·��.
    Dim strWbkPath As String    ' ���������ڵ�·��.
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    ThisWorksheet = ActiveSheet.Name
    On Error Resume Next
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    
    ThisWorkBookName = ThisWorkbook.Name
    FileName = ThisWorkbook.Path
    SourceWorkBook = "����Դ-" & Major & "-ָ������ݾ���.xls"
    FileName = FileName & "\" & SourceWorkBook
    If MyFile.FileExists(FileName) = False Then
      Call MsgInfo(NoMsgBox, "����Դ-" & Major & "-ָ������ݾ���.xls�����ڣ��뽫����Դ-" & Major & "-ָ������ݾ���.xls���Ƶ���ǰ�ļ���")
      FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
    '����Windows���ļ��Ի���
    If (FileName = False) Then Exit Sub
    SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    On Error Resume Next
    Worksheets("רҵ����״̬").Activate
    MatrixSheet = Application.Index(Range("C4:C" & MajorLastRow), Application.Match(Major, Range("B4:B" & MajorLastRow), 0))
    If Sheets(MatrixSheet) Is Nothing Then
        Worksheets("��ҵҪ��-ָ������ݱ�").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = MatrixSheet
    End If
    Worksheets(MatrixSheet).Activate
    Worksheets(MatrixSheet).Visible = True
    Call CopySheet(FileName, SourceWorkBook, "ָ���γ̾���", "A:AS", ThisWorkBookName, MatrixSheet, "A:AS")
    Application.ScreenUpdating = False
    Worksheets(ThisWorksheet).Activate
End Sub 'ѡ����ļ���û����ʵ�İ�����
Sub ��ȡ��ҵҪ�������Ϣ(Major As String)
    Dim ThisWorksheet As String
    Dim FileName As Variant
     '���ļ��Ի��򷵻ص��ļ�������һ��ȫ·���ļ�������ֵҲ������False���������ΪVariant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim strCurPath As String    ' Ӧ�ó���ĵ�ǰ·��.
    Dim strWbkPath As String    ' ���������ڵ�·��.
    Dim CourseCount As Integer
    Dim PointCount As Integer
    
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    ThisWorksheet = ActiveSheet.Name
    On Error Resume Next
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    ThisWorkBookName = ThisWorkbook.Name
    FileName = ThisWorkbook.Path
    SourceWorkBook = "����Դ-" & Major & "-ָ������ݾ���.xls"
    FileName = FileName & "\" & SourceWorkBook
    Sheets("רҵ����״̬").Visible = True
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Worksheets("רҵ����״̬").Activate
    If MyFile.FileExists(FileName) = False Then
        'ָ������ݾ����ļ�������
        Range("D" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = "������"
    Else
        Range("D" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = "����"
    End If
    Workbooks.Open FileName
    Workbooks(SourceWorkBook).Activate
    Application.ScreenUpdating = False

    Worksheets("ָ���γ̾���").Activate
    CourseCount = Application.CountA(Range("D4:D200"))
    PointCount = Application.CountA(Range("E4:AS200"))
    Workbooks(ThisWorkBookName).Activate
    Worksheets("רҵ����״̬").Activate
    Range("E" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = CourseCount
    Range("F" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = PointCount
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Sheets("רҵ����״̬").Visible = False
    Workbooks(SourceWorkBook).Activate
    If ActiveWorkbook.Name = SourceWorkBook Then ActiveWorkbook.Close True
End Sub
Sub ����ѧ������()
    Dim FileName As Variant
     '���ļ��Ի��򷵻ص��ļ�������һ��ȫ·���ļ�������ֵҲ������False���������ΪVariant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim temp As String
    Dim strCurPath As String    ' Ӧ�ó���ĵ�ǰ·��.
    Dim strWbkPath As String    ' ���������ڵ�·��.
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    
    ThisWorkBookName = ThisWorkbook.Name
    CourseNumber = Range("$B$3").Value
    FileName = ThisWorkbook.Path
    SourceWorkBook = "����Դ-ѧ������.xls"
    FileName = FileName & "\" & SourceWorkBook
    If MyFile.FileExists(FileName) = False Then
      Call MsgInfo(NoMsgBox, "����Դ-ѧ������.xls�����ڣ����ֶ�ָ��")
      FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
    '����Windows���ļ��Ի���
    If (FileName = False) Then Exit Sub
    SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:H", ThisWorkBookName, "ѧ������", "A:H")
    Worksheets("ѧ������").Visible = True
    Worksheets("ѧ������").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Count = Application.CountA(Range("C:C"))
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=RC2&""-""&RC6"
    Range("G1:G1").Select
    Selection.AutoFill Destination:=Range("G1:G" & Count), Type:=xlFillDefault
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("ѧ������").Visible = False
    Application.ScreenUpdating = False
End Sub 'ѡ����ļ���û����ʵ�İ�����
Sub ����ɼ���()
     '���ļ��Ի��򷵻ص��ļ�������һ��ȫ·���ļ�������ֵҲ������False���������ΪVariant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Dim i As Integer
    Dim j As Integer
    Const Num As String = "AD"
    Const KeyRow As String = "AG"
    Const Key1Col As String = "AH"
    Const Key2Col As String = "AK"
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim temp As String
    Dim strCurPath As String    ' Ӧ�ó���ĵ�ǰ·��.
    Dim strWbkPath As String    ' ���������ڵ�·��.
    Dim CountNum As Integer
    Dim NumExitFlag As Boolean
    Dim NumRow As Integer
    Dim Msg As String
    Dim Num2Col As Integer
    Dim ThisSheetName As String
    Dim ScoreType(6) As String
    Dim CountItem As Integer
    Dim Table(8) As String
    Application.ScreenUpdating = False
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    Application.EnableEvents = False
    ThisSheetName = ActiveSheet.Name
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    CourseNumber = Range("$B$3").Value
    FileName = ThisWorkbook.Path
    SourceWorkBook = "�ɼ���-" & CourseNumber & ".xls"
    FileName = FileName & "\" & SourceWorkBook
    '�ɼ���-1720835
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    If (Range("Q3").Value = "��֤δ�ύ�ɼ�") Then
        Call MsgInfo(NoMsgBox, "���γ�Ŀ����ۺϷ�����������֤״̬Ϊδ�ύ�ɼ�������Ҫ����ɼ���")
        Worksheets(ThisSheetName).Activate
        Exit Sub
    End If
    For i = 0 To 5
        ScoreType(i) = Cells(2, 2 * i + 4).Value
    Next i
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    ThisWorkBookName = ThisWorkbook.Name
    If (ThisSheetName = "1-�Ծ�ɼ��ǼǱ���д��") Then
        Call �Ծ�ɼ��ǼǱ�ʽ
    ElseIf (ThisSheetName = "�ɼ��˶�") Then
        Call �Ծ�ɼ��ǼǱ�ʽ
        Call �ɼ��˶Ա�ʽ
    End If
    If MyFile.FileExists(FileName) = False Then
        Call MsgInfo(NoMsgBox, "�ɼ���-" & CourseNumber & ".xls" & "�����ڣ����ֶ�ָ��")
        Msg = "�Ѿ��ڽ���ϵͳ�ύ�ɼ����ɼ�����Ҫ��" & vbCr
        Msg = Msg & "�ɼ����������Ҫ������ѧ�š�������ʵ��ɼ���ƽʱ�ɼ������˳ɼ��������ɼ����ɼ������ҵ�ɼ������ò��顢�γ̱��桿�ȹؼ��ʣ�" & vbCr
        Call MsgInfo(NoMsgBox, Msg)
        FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
        '����Windows���ļ��Ի���
        If (FileName = False) Then Exit Sub
        SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:N", ThisWorkBookName, "�ɼ���", "B:O")

    Application.ScreenUpdating = False
    Worksheets("�ɼ���").Visible = True
    Worksheets("�ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Columns("A:A").Select
    Selection.ClearContents
    Worksheets("�ɼ���").Activate
    For i = 0 To 5
         Cells(1, i + 23).Value = ScoreType(i)
    Next i
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Call �ɼ���ʽ
    Worksheets("�ɼ���").Visible = True
    Worksheets("�ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:O600").Select
    Selection.UnMerge

    Range(KeyRow & "3:" & Key1Col & "14").Select
    Selection.ClearContents
    Range(Key2Col & "3:" & Key2Col & "14").Select
    Selection.ClearContents
    NumExitFlag = False
    'ȷ���ɼ��������

    For j = 1 To 10
        If Not (isError(Application.Match("*ѧ*��*", Range("A" & j & ":" & "O" & j), 0))) Then
            NumRow = j
            CountNum = Application.WorksheetFunction.CountIf(Range("A" & j & ":" & "O" & j), "*ѧ*��*")
            Num2Col = Application.Match("*��*��*", Range("A" & j & ":" & "O" & j), 0) + 2
        End If
        If Not (isError(Application.Match("*��*��*", Range("A" & j & ":" & "O" & j), 0))) Then
            NumExitFlag = True
        End If
    Next j
    '�ɼ���Ϊѧ���ƹ���ϵͳ�ɼ����ʽ��û������У���Ҫ���������
    If (Not NumExitFlag) Then
        Range("A" & NumRow).Value = "���"
        Range("A" & NumRow + 1).Select
        ActiveCell.FormulaR1C1 = "=IF(RC[1]="""","""",IF(R[-1]C[1]=""ѧ��"",1,R[-1]C+1))"
        Range("A" & NumRow + 1).Select
        Selection.AutoFill Destination:=Range("A" & NumRow + 1 & ":A600"), Type:=xlFillDefault
        Range("A" & NumRow + 1 & ":A600").Select
    End If
    For i = 3 To 15
        For j = 1 To 10
            If (CountNum = 1) Then
                If Not (isError(Application.Match("*" & Range(Num & i).Value & "*", Range("A" & j & ":" & "O" & j), 0))) Then
                    Range(KeyRow & i).Value = j
                    Range(Key1Col & i).Value = Application.Match("*" & Range(Num & i).Value & "*", Range("A" & j & ":" & "O" & j), 0)
                End If
            ElseIf (CountNum = 2) Then
                If Not (isError(Application.Match("*" & Range(Num & i).Value & "*", Range("A" & j & ":" & "I" & j), 0))) Then
                    Range(KeyRow & i).Value = j
                    Range(Key1Col & i).Value = Application.Match("*" & Range(Num & i).Value & "*", Range("A" & j & ":" & "I" & j), 0)
                End If
                If Not (isError(Application.Match("*" & Range(Num & i).Value & "*", Range(Cells(j, Num2Col), Cells(j, 15)), 0))) Then
                    Range(Key2Col & i).Value = Application.Match("*" & Range(Num & i).Value & "*", Range(Cells(j, Num2Col), Cells(j, 15)), 0) + Range(Cells(1, 1), Cells(1, Num2Col)).Cells.Count - 1
                End If
            End If
         Next j
    Next i
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("�ɼ���").Visible = False
    Worksheets("1 - �Ծ�ɼ��ǼǱ�(��д)").Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub 'ѡ����ļ���û����ʵ�İ����򿪡�

Sub �������ϵͳ�ɼ���()
    Dim FileName As Variant
     '���ļ��Ի��򷵻ص��ļ�������һ��ȫ·���ļ�������ֵҲ������False���������ΪVariant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Dim i As Integer
    Dim j As Integer
    Const Num As String = "AD"
    Const KeyRow As String = "AG"
    Const Key1Col As String = "AH"
    Const Key2Col As String = "AK"
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim temp As String
    Dim strCurPath As String    ' Ӧ�ó���ĵ�ǰ·��.
    Dim strWbkPath As String    ' ���������ڵ�·��.
    Dim CountNum As Integer
    Dim NumExitFlag As Boolean
    Dim NumRow As Integer
    Dim Msg As String
    Dim Num2Col As Integer
    Dim ThisSheetName As String
    Dim ScoreType(6) As String
    Application.ScreenUpdating = False
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ThisSheetName = ActiveSheet.Name
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    CourseNumber = Range("$B$3").Value
    FileName = ThisWorkbook.Path
    SourceWorkBook = CourseNumber & "��ĩ�ɼ�.xls"
    FileName = FileName & "\" & SourceWorkBook

    For i = 0 To 5
        ScoreType(i) = Cells(2, 2 * i + 4).Value
    Next i
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    ThisWorkBookName = ThisWorkbook.Name
    If (ThisSheetName = "1-�Ծ�ɼ��ǼǱ���д��") Then
        Call �Ծ�ɼ��ǼǱ�ʽ
    ElseIf (ThisSheetName = "�ɼ��˶�") Then
        Call �Ծ�ɼ��ǼǱ�ʽ
        Call �ɼ��˶Ա�ʽ
    End If
    If MyFile.FileExists(FileName) = False Then
        Call MsgInfo(NoMsgBox, CourseNumber & "��ĩ�ɼ�.xls" & "�����ڣ����ֶ�ָ��")
        Msg = "��ѧ���ƹ���ϵͳ�ύ�ɼ�ǰ�������سɼ����ڳɼ��˶Թ������ֱ�ӵ��룬�ɼ����������Ҫ������ѧ�š�������ʵ��ɼ���ƽʱ�ɼ������˳ɼ��������ɼ����ɼ���𡿵ȹؼ��ʣ�" & vbCr
        Call MsgInfo(NoMsgBox, Msg)
        FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
        '����Windows���ļ��Ի���
        If (FileName = False) Then Exit Sub
        SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    Worksheets("�ɼ���").Visible = True
    Worksheets("�ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:O600").Select
    Selection.UnMerge
    Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:N", ThisWorkBookName, "�ɼ���", "B:O")
    Application.ScreenUpdating = False
    Worksheets("�ɼ���").Visible = True
    Worksheets("�ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    For i = 0 To 5
         Cells(1, i + 23).Value = ScoreType(i)
    Next i
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Call �ɼ���ʽ
    Worksheets("�ɼ���").Visible = True
    Worksheets("�ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:O600").Select
    Selection.UnMerge
    Range("A1:A600").Select
    Selection.ClearContents
    Range(KeyRow & "3:" & Key1Col & "15").Select
    Selection.ClearContents
    Range(Key2Col & "3:" & Key2Col & "15").Select
    Selection.ClearContents
    NumExitFlag = False
    'ȷ���ɼ��������

    For j = 1 To 10
        If Not (isError(Application.Match("ѧ��", Range("A" & j & ":" & "O" & j), 0))) Then
            NumRow = j
            CountNum = Application.WorksheetFunction.CountIf(Range("A" & j & ":" & "O" & j), "ѧ��")
            Num2Col = Application.Match("����*", Range("A" & j & ":" & "O" & j), 0) + 2
        End If
        If Not (isError(Application.Match("���", Range("A" & j & ":" & "O" & j), 0))) Then
            NumExitFlag = True
        End If
    Next j
    '�ɼ���Ϊѧ���ƹ���ϵͳ�ɼ����ʽ��û������У���Ҫ���������
    If (Not NumExitFlag) Then
        Range("A" & NumRow).Value = "���"
        Range("A" & NumRow + 1).Select
        ActiveCell.FormulaR1C1 = "=IF(RC[1]="""","""",IF(R[-1]C[1]=""ѧ��"",1,R[-1]C+1))"
        Range("A" & NumRow + 1).Select
        Selection.AutoFill Destination:=Range("A" & NumRow + 1 & ":A600"), Type:=xlFillDefault
        Range("A" & NumRow + 1 & ":A600").Select
    End If
    For i = 3 To 15
        For j = 1 To 10
            If (CountNum = 1) Then
                If Not (isError(Application.Match("*" & Range(Num & i).Value & "*", Range("A" & j & ":" & "O" & j), 0))) Then
                    Range(KeyRow & i).Value = j
                    Range(Key1Col & i).Value = Application.Match("*" & Range(Num & i).Value & "*", Range("A" & j & ":" & "O" & j), 0)
                End If
            ElseIf (CountNum = 2) Then
                If Not (isError(Application.Match("*" & Range(Num & i).Value & "*", Range("A" & j & ":" & "I" & j), 0))) Then
                    Range(KeyRow & i).Value = j
                    Range(Key1Col & i).Value = Application.Match("*" & Range(Num & i).Value & "*", Range("A" & j & ":" & "I" & j), 0)
                End If
                If Not (isError(Application.Match("*" & Range(Num & i).Value & "*", Range(Cells(j, Num2Col), Cells(j, 15)), 0))) Then
                    Range(Key2Col & i).Value = Application.Match("*" & Range(Num & i).Value & "*", Range(Cells(j, Num2Col), Cells(j, 15)), 0) + Range(Cells(1, 1), Cells(1, Num2Col)).Cells.Count - 1
                End If
            End If
         Next j
    Next i
    Worksheets("�ɼ���").Visible = False
    Worksheets(ThisSheetName).Activate
    Application.ScreenUpdating = True
End Sub 'ѡ����ļ���û����ʵ�İ����򿪡�
Sub ����ʵ��ɼ���()
    Dim FileName As Variant
     '���ļ��Ի��򷵻ص��ļ�������һ��ȫ·���ļ�������ֵҲ������False���������ΪVariant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Dim i As Integer
    Dim j As Integer
    Const Num As String = "U"
    Const KeyRow As String = "X"
    Const Key1Col As String = "Y"
    Const Key2Col As String = "AB"
    Dim strCurPath As String    ' Ӧ�ó���ĵ�ǰ·��.
    Dim strWbkPath As String    ' ���������ڵ�·��.

    Application.ScreenUpdating = False
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Application.EnableEvents = False
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    
    ThisWorkBookName = ThisWorkbook.Name
    CourseNumber = Range("$B$3").Value
    FileName = ThisWorkbook.Path
    SourceWorkBook = "ʵ��ɼ���-" & CourseNumber & ".xls"
    FileName = FileName & "\" & SourceWorkBook
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    
    If MyFile.FileExists(FileName) = False Then
        Call MsgInfo(NoMsgBox, "ʵ��ɼ���-" & CourseNumber & ".xls" & "�����ڣ����ֶ�ָ��")
        FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
        '����Windows���ļ��Ի���
        If (FileName = False) Then Exit Sub
        SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    
    Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:N", ThisWorkBookName, "����ʵ��ɼ���", "A:N")
    Application.ScreenUpdating = False
    Worksheets("����ʵ��ɼ���").Visible = True
    Worksheets("����ʵ��ɼ���").Activate
    Call ʵ��ɼ���ʽ
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("X3:Y5").Select
    Selection.ClearContents
    Range("AB3:AB5").Select
    Selection.ClearContents
    For i = 3 To 5
        For j = 1 To 5
            If Not (isError(Application.Match(Mid(Range(Num & i).Value, 1, 1) & "*", Range("A" & j & ":" & "G" & j), 0))) Then
                Range(KeyRow & i).Value = j
                Range(Key1Col & i).Value = Application.Match(Mid(Range(Num & i).Value, 1, 1) & "*", Range("A" & j & ":" & "G" & j), 0)
            End If
            If Not (isError(Application.Match(Mid(Range(Num & i).Value, 1, 1) & "*", Range("H" & j & ":" & "N" & j), 0))) Then
                Range(Key2Col & i).Value = Application.Match(Mid(Range(Num & i).Value, 1, 1) & "*", Range("H" & j & ":" & "N" & j), 0) + Range("A1:H1").Cells.Count - 1
            End If
         Next j
    Next i
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("����ʵ��ɼ���").Visible = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub 'ѡ����ļ���û����ʵ�İ����򿪡�
Function �����ѧ���̵ǼǱ�() As Boolean
    Dim FileName As Variant
     '���ļ��Ի��򷵻ص��ļ�������һ��ȫ·���ļ�������ֵҲ������False���������ΪVariant
    Dim ThisWorkBookName As String
    Dim Xueqi As String
    Dim Count As Integer
    Dim SourceWorkBook As String
    Dim FirstRow As Integer
    Dim LaseRow As Integer
    Dim CourseNum As String
    Dim CourseNumFlag As Boolean
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim strCurPath As String    ' Ӧ�ó���ĵ�ǰ·��.
    Dim strWbkPath As String    ' ���������ڵ�·��.
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* Ӧ�ó���ĵ�ǰ·�����ǹ��������ڵ�·��. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ThisWorkBookName = ThisWorkbook.Name
    Xueqi = Range("$AG$1").Value
    If Range("B4").Value = "��������Դ-��ѧ��������ӸÿκŵĿγ���Ϣ" Then
        Application.EnableEvents = True
        �����ѧ���̵ǼǱ� = False
        Exit Function
    Else
        SourceWorkBook = "����Դ-" & Xueqi & "ѧ�ڽ�ѧ���̵ǼǱ�.xls"
        FileName = ThisWorkbook.Path & "\" & SourceWorkBook
        If MyFile.FileExists(FileName) = False Then
            Call MsgInfo(NoMsgBox, FileName & "�����ڣ����ֶ�ָ��" & Xueqi & "ѧ�ڵĽ�ѧ���̵ǼǱ�֧�ֵ���ѧ�ֹ���ϵͳ���سɼ���������Ҳ���Դӽ���ϵͳ���ƽ�ѧ���̵ǼǱ���������EXCEL�ĵ�")
            FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
            If FileName = False Then
                Sheets("��ѧ���̵ǼǱ�").Visible = False
                Application.EnableEvents = True
                �����ѧ���̵ǼǱ� = False
                Exit Function
            End If
            '����Windows���ļ��Ի���
            SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
        End If
        Sheets("��ѧ���̵ǼǱ�").Visible = 1
        Worksheets("��ѧ���̵ǼǱ�").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Range("A1:Z40000").Select
        Selection.ClearContents
        Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:C", ThisWorkBookName, "��ѧ���̵ǼǱ�", "C:E")
        Application.ScreenUpdating = False
        Sheets("��ѧ���̵ǼǱ�").Visible = 1
        Worksheets("��ѧ���̵ǼǱ�").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Range("A4").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(MID(R[-3]C[2],1,4)=""�ο���ʦ"",MID(R[-3]C3,FIND(""�γ����:"",R[-3]C[2],1)+5,LEN(R[-3]C3)-FIND(""�γ����:"",R[-3]C[2],1)),R[-1]C)"
        Range("B4").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(OR(RC[1]=""���"",RC[1]="""",RC[3]=""""),"""",RC[-1]&""-""&RC[1])"
        Range("A4:B4").Select
        Selection.AutoFill Destination:=Range("A4:B50000"), Type:=xlFillDefault
        Range("A4:B50000").Select
        
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        If (Range("$T$1").Value = "������") Then
            Call MsgInfo(NoMsgBox, "����Ľ�ѧ���̵ǼǱ���û�иÿκŵ����������ֶ�ָ����ѧ���̵ǼǱ��ļ���֧�ֵ���ѧ�ֹ���ϵͳ���سɼ���������Ҳ���Դӽ���ϵͳ���ƽ�ѧ���̵ǼǱ���������EXCEL�ĵ�")
            FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
            '����Windows���ļ��Ի���
            If (FileName = False) Then
                Application.EnableEvents = True
                �����ѧ���̵ǼǱ� = False
                Exit Function
            End If
          SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
          Sheets("��ѧ���̵ǼǱ�").Visible = 1
          Worksheets("��ѧ���̵ǼǱ�").Activate
          Range("C1:Z40000").Select
          Selection.ClearContents
          Selection.UnMerge
          Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:C", ThisWorkBookName, "��ѧ���̵ǼǱ�", "C:E")
          Application.ScreenUpdating = False
          Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        End If
        CourseNum = Range("B3").Value

        Sheets("��ѧ���̵ǼǱ�").Visible = 1
        Worksheets("��ѧ���̵ǼǱ�").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        
        '���������Ϊ����ϵͳ���صĳɼ���������ʽ
        If Not (isError(Application.Match("*" & CourseNum & "*", Range("K:K"), 0))) Then
            Count = Application.WorksheetFunction.CountA(Range("C4:C200"))
            Range("C1:K200").Select
            Selection.UnMerge
            Columns("C:C").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("C3").Select
            ActiveCell.FormulaR1C1 = "���"
            Range("A4").Select
            ActiveCell.FormulaR1C1 = CourseNum
            Range("C4").Select
            ActiveCell.FormulaR1C1 = "1"
            Range("B4").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]&""-""&RC[1]"
            Range("A5").Select
            ActiveCell.FormulaR1C1 = CourseNum
            Range("C5").Select
            ActiveCell.FormulaR1C1 = "2"
            Range("B5").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]&""-""&RC[1]"
            Range("A4:C5").Select
            Selection.AutoFill Destination:=Range("A4:C" & Count + 3), Type:=xlFillDefault
            Call �������ù�ʽ(Count)
            
        End If
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        
        FirstRow = Range("$T$1").Value
        LastRow = Range("$W$1").Value
        Worksheets("��ѧ���̵ǼǱ�").Activate
        If (LastRow <> "������") Then
          Range("A" & LastRow & ":Z40000").Select
          Selection.ClearContents
        End If
        If (FirstRow <> 0) Then
            Range("A1:Z" & FirstRow).Select
            Selection.ClearContents
            Selection.Delete Shift:=xlUp
        End If
        Sheets("��ѧ���̵ǼǱ�").Visible = 0
    End If
    Worksheets("��ѧ���̵ǼǱ�").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Call ���ý�ѧ���̵ǼǱ�
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
    �����ѧ���̵ǼǱ� = True
End Function 'ѡ����ļ���û����ʵ�İ����򿪡�
Sub �½�רҵ����״̬������()
    Dim AllowEditCount As Integer
    On Error Resume Next
    Worksheets("רҵ����״̬").Activate
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ѧԺ��רҵ��ָ������ݾ������ü�״̬"
    Range("A2").Select
    Selection.FormulaR1C1 = "ѧԺ����"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "���"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "רҵ����"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "����������"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "�ļ�״̬"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "�γ�����"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "ָ�����"
    Range("A1:F1").Select
    Selection.UnMerge
    Range("B2:F2").Select
    Selection.UnMerge
    Rows("1:12").Select
    Selection.RowHeight = 30
    Columns("A:A").Select
    Selection.ColumnWidth = 10
    Columns("B:B").Select
    Selection.ColumnWidth = 25
    Columns("C:C").Select
    Selection.ColumnWidth = 20
    Columns("D:F").Select
    Selection.ColumnWidth = 10
    Range("A2:F12").Select
    With Selection.Font
        .Name = "����"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Bold = True
    Range("A1").Select
    With Selection.Font
        .Name = "����"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Bold = True
    Range("A1:F1").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("B2:F2").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("A1:F12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[1]="""","""",IF(R[-1]C=""���"",1,R[-1]C+1))"
    Range("A4").Select
    Selection.AutoFill Destination:=Range("A4:A12"), Type:=xlFillDefault
    Range("A4:A12").Select
    
    AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
    If (AllowEditCount <> 0) Then
        For i = 1 To AllowEditCount
            Sheets("רҵ����״̬").Protection.AllowEditRanges(1).Delete
        Next i
    End If
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "�汾��"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "�޶�����"
    Range("H1:H2").Select
    ActiveSheet.Protection.AllowEditRanges.Add Title:="רҵ", Range:=Range("B4:C12")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="ѧԺ", Range:=Range("B2:D2")
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub CopySheet(SourceWorkBookFileName As Variant, SourceWorkBook As String, SourceSheet As String, SourceCol As String, TargetWorkBook As String, TargetSheet As String, _
TargetCol As String)
    Workbooks(TargetWorkBook).Activate
    Worksheets(TargetSheet).Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Columns(TargetCol).Select
    Selection.ClearContents
    Workbooks.Open SourceWorkBookFileName
    Workbooks(SourceWorkBook).Activate
    Worksheets(SourceSheet).Activate
    Columns(SourceCol).Select
    Selection.Copy
    
    Workbooks(TargetWorkBook).Activate
    Worksheets(TargetSheet).Activate
    Range(TargetCol).Select
    ActiveSheet.Paste
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets(TargetSheet).Visible = 0
    Workbooks(SourceWorkBook).Activate
    If ActiveWorkbook.Name = SourceWorkBook Then ActiveWorkbook.Close True
End Sub
Sub CopySheet1(SourceWorkBookFileName As Variant, SourceWorkBook As String, SourceSheet As String, SourceStartRow As Integer, SourceStartCol As Integer, SourceEndRow As Integer, SourceEndCol As Integer, TargetWorkBook As String, TargetSheet As String, _
TargetStartRow As Integer, TargetStartCol As Integer, TargetEndRow As Integer, TargetEndCol As Integer, PasteSpecial As String)
    Workbooks(TargetWorkBook).Activate
    Worksheets(TargetSheet).Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range(Cells(TargetStartRow, TargetStartCol), Cells(600, TargetEndCol)).Select
    Selection.ClearContents
    'Workbooks.Open SourceWorkBookFileName
    Workbooks(SourceWorkBook).Activate
    Worksheets(SourceSheet).Activate
    Range(Cells(SourceStartRow, SourceStartCol), Cells(SourceEndRow, SourceEndCol)).Select
    Selection.Copy
    Workbooks(TargetWorkBook).Activate
    Worksheets(TargetSheet).Activate
    Range(Cells(TargetStartRow, TargetStartCol), Cells(TargetEndRow, TargetEndCol)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=PasteSpecial, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    'Sheets(TargetSheet).Visible = 0
    'Workbooks(SourceWorkBook).Activate
    'If ActiveWorkbook.Name = SourceWorkBook Then ActiveWorkbook.Close True
End Sub

'�޶����ڣ�2019��4��15��
''�ɼ������  �޶����ڣ�2019��1��25��
  Sub CreateRecordWorkBook()

    Dim i As Long
    Dim FolderName  As String
    Dim CourseNumber As String
    Dim Term As String
    Dim CourseName As String
    Dim ProportionNomal As Double
    Dim ProportionExperiment As Double
    Dim ProportionMidterm As Double
    Dim ProportionExamine As Double
    Dim StudentNumber As String
    Dim StudentName As String
    Dim ExperimentScore As String
    Dim NomalScore As String
    Dim MidtermScore As String
    Dim ExamineScore As String
    Dim SumScore As String
    Dim ScoreCategory As String
    Dim ExamineCategory As String
    Dim Count As Integer
    Dim TuikeCount As Integer
    
    Dim CourseFileName As String
    Dim MobanWorkbookName As String
    Dim CourseWorkbookName As String
    Dim CurrentWorksheet As String
    CurrentWorksheet = ActiveSheet.Name
    Application.ScreenUpdating = False
    Sheets("�ɼ�¼��").Visible = True
    Sheets("�ɼ�¼��").Select
    FolderName = ThisWorkbook.Path
    MobanWorkbookName = ThisWorkbook.Name
    
    Sheets("2-�γ�Ŀ����ۺϷ�������д��").Select
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    CourseNumber = Range("B3").Value
    Term = Range("AG2").Value
    CourseName = Range("B4").Value
    ProportionNomal = Range("T4").Value
    ProportionExperiment = Range("U4").Value
    ProportionMidterm = Range("W4").Value
    ProportionExamine = Range("V4").Value
    Count = Range("S3").Value
    Sheets("2-�γ�Ŀ����ۺϷ�������д��").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    CourseFileName = CourseNumber & "�ɼ�¼���.xls"
    
    Call CreateNewWorkbook(FolderName, CourseFileName)
    CourseWorkbookName = ThisWorkbook.Name
    Range("1:200").Select
    Selection.UnMerge
    
    Workbooks(MobanWorkbookName).Activate
    Sheets("�ɼ�¼��").Select
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    
    Rows("1:200").Select
    Selection.Copy

    
    Workbooks(CourseFileName).Activate
    Sheets("Sheet1").Select
    Rows("1:200").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("A1:I1").Select
    Selection.Merge
    Range("A2:H2").Select
    Selection.Merge
    
    If (ProportionExperiment = 0) Then
        Columns("C:C").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    If (ProportionMidterm = 0) Then
        Columns("D:D").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    Workbooks(CourseFileName).Activate
    Sheets("Sheet1").Select
    TuikeCount = 0
    For i = 1 To 200
    ScoreCategory = Range("H" & i).Value
    If (ScoreCategory = "�˿�") Then
      Rows(i & ":" & i).Select
      Selection.Delete
    End If
    ExamineCategory = Range("I" & i).Value
    If (ScoreCategory = "����") Then
      Rows(i & ":" & i).Select
      Selection.Delete
    End If
    Next i
    
    Range("1:200").Select
    Selection.UnMerge
    Rows("2:" & Count + 3).Select
    Selection.RowHeight = 20
    
    Call SetColumnWidth("A", 10)
    Call SetColumnWidth("B", 10)
    Call SetColumnWidth("C", 10)
    Call SetColumnWidth("D", 10)
    Call SetColumnWidth("E", 10)
    Call SetColumnWidth("F", 10)
    Call SetColumnWidth("G", 10)
    Call SetColumnWidth("H", 15)
    Call SetColumnWidth("I", 15)
    
    Range("A1:I1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Merge
    With Selection.Font
        .Name = "����"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Bold = True
    Range("A2:H2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Merge
    Range("H3:I3").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWorkbook.Save
    ActiveWindow.ScrollColumn = 1
    Workbooks(MobanWorkbookName).Activate
    Sheets("�ɼ�¼��").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets("�ɼ�¼��").Visible = False
    
    Workbooks(MobanWorkbookName).Activate
    Application.ScreenUpdating = True
    Worksheets(CurrentWorksheet).Activate
End Sub
   
 
 
 Sub SetSheetFormat(SetMultiCol1 As String, SetMultiCol2 As String)
    Columns(SetMultiCol1 & ":" & SetMultiCol2).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Public Sub SetRowHeight(SetRow As String, SetRowHeight As Long)
  Rows(SetRow & ":" & SetRow).Select
  Selection.RowHeight = SetRowHeight
End Sub
Public Sub SetMultiRowHeight(SetRow1 As String, SetRow2 As String, SetRowHeight As Long)
  Rows(SetRow1 & ":" & SetRow2).Select
  Selection.RowHeight = SetRowHeight
End Sub
Public Sub SetColumnWidth(SetCol As String, SetColWidth As Long)
  Columns(SetCol & ":" & SetCol).Select
  Selection.ColumnWidth = SetColWidth
End Sub
Public Sub SetAutoLineFeed(SetCol As String)
    Columns(SetCol & ":" & SetCol).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Public Sub SetColCenter(SetMultiCol As String)
    Range(SetMultiCol & ":" & SetMultiCol).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub CreateNewWorkbook(FolderName As String, FileName As String)
    Dim Wk As Workbook
    Dim WorkBookName As String
    On Error Resume Next
    WorkBookName = FileName
    Application.DisplayAlerts = False
    Workbooks(WorkBookName).Activate
    If Err.Number <> 0 Then
      If FileExists(FolderName & "/" & FileName) Then
        Workbooks.Open FolderName & "/" & FileName
      Else
        Set Wk = Workbooks.Add
        Wk.SaveAs FileName:=FolderName & "/" & FileName
      End If
    End If
End Sub
Function FileExists(FullFileName As String) As Boolean
  '�������������,�򷵻�True
  FileExists = Len(Dir(FullFileName)) > 0
End Function
Sub CreateNewSheet(SheetName As String)
  On Error Resume Next
  '��SheetName�Ĺ���������,���½�һ��������
  Worksheets(SheetName).Activate
  If Err.Number <> 0 Then
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = SheetName
  Else
    ActiveSheet.Clear
  End If
End Sub
Sub CreateTXTfile(FileName As String, Content As String, AfterOpen As Boolean)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set MyFileobj = fso.CreateTextFile(FileName, True, True)
    MyFileobj.Write (Content)
    MyFileobj.Close
    If AfterOpen Then
        For Each Process In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='notepad.exe'")
            Process.Terminate (0)
        Next
        Shell "notepad.exe " & FileName
    End If
End Sub
Sub �γ�Ŀ����ۺϷ���(Mode As String, TargetValue As String, TargetRow As String, TargetColumn As String)
    Dim ImportStatus As Boolean
    Dim TempMsgBox As Boolean
    Select Case Mode
        Case "��֤״̬"
            If TargetValue = "����֤" Then
                Call ����֤
            ElseIf TargetValue = "��֤δ�ύ�ɼ�" Then
                Call ��֤δ�ύ�ɼ�
            ElseIf TargetValue = "��֤���ύ�ɼ�" Then
                Call ��֤���ύ�ɼ�
            End If
        Case "�γ����"
            If TargetValue <> "" Then
                If Range("B4").Value = "��������Դ-��ѧ��������ӸÿκŵĿγ���Ϣ" Then
                    Call �����ѧ����
                    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
                End If
                If Range("Q3").Value = "��֤δ�ύ�ɼ�" Or Range("Q3").Value = "����֤" Then
                    ImportStatus = �����ѧ���̵ǼǱ�
                    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
                    SumCount = Range("AM1").Value
                    If (Application.WorksheetFunction.Count(Range("A6:A200")) <> SumCount) And SumCount <> 0 Then
                        Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
                        Call ���ý�ѧ���̵ǼǱ�
                    End If
                ElseIf Range("Q3").Value = "��֤���ύ�ɼ�" Then
                    TempMsgBox = NoMsgBox
                    NoMsgBox = True
                    ImportStatus = �����ѧ���̵ǼǱ�
                    NoMsgBox = TempMsgBox
                End If
                Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
            End If
        Case "��֤רҵ"
            If TargetValue <> "" Then
                Sheets("רҵ����״̬").Visible = True
                Worksheets("רҵ����״̬").Activate
                If Not isError(Application.Match(TargetValue, Range("B4:B" & MajorLastRow), 0)) Then
                    Sheets("רҵ����״̬").Visible = False
                    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
                    If (Range("B6").Value = 0) Then
                        Application.ScreenUpdating = False
                        Application.EnableEvents = False
                        Call ����ѧ������
                        Application.ScreenUpdating = True
                        Application.EnableEvents = True
                        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
                    End If
                    Application.ScreenUpdating = False
                    Call �������
                    Call ָ������ݱ�ʽ
                    Call ��ҵҪ�����ݱ�ʽ
                    Application.ScreenUpdating = True
                    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
                End If
            End If
    End Select
End Sub
Sub ����֤()
    Dim WorkBookName As String
      '����Ҫ���д�ɶ����۵Ŀγ̣�ֻ��Ҫ��д��ѧ���̵ǼǱ��2-�γ�Ŀ����ۺϷ�������д�����еĿκţ����ۻ��ڱ�������Ϣ����ӡ��ѧ���̵ǼǱ�������������档
        Application.ScreenUpdating = False
        Call ���ñ������
        WorkBookName = ThisWorkbook.Name
        Workbooks(WorkBookName).Activate
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        '����"2-�γ�Ŀ����ۺϷ�������д��"��ɶ����۲��ֱ��
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("2-�γ�Ŀ����ۺϷ�������д��").Rows("10:28").Select
        Selection.EntireRow.Hidden = True
        Call �γ�Ŀ������༭����
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = True
        Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Rows("11:20").Select
        Selection.EntireRow.Hidden = False
        SumRow = Application.WorksheetFunction.CountA(Range("B11:B20"))
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Rows(SumRow + 11 & ":20").Select
        Selection.EntireRow.Hidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        '����3-��ҵҪ�����ݱ���д����1-�γ�Ŀ���ɶ����ۣ���ӡ����2-��ҵҪ���ɶ����ۣ���ӡ����3-�ۺϷ�������ӡ��������
        Sheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Visible = True
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Visible = True
        Sheets("4-�����������棨��д+��ӡ��").Visible = True
        Sheets("�ɼ��˶�").Visible = True
        Sheets("3-��ҵҪ�����ݱ���д��").Visible = False
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = False
        Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
        Sheets("3-�ۺϷ�������ӡ��").Visible = False
        
        Sheets("�ɼ�¼��").Visible = False
        Sheets("��ѧ���̵ǼǱ�").Visible = False
        Sheets("ѧ������").Visible = False
        Sheets("����ʵ��ɼ���").Visible = False
        Sheets("ƽʱ�ɼ���").Visible = False
        Sheets("�ɼ���").Visible = False
        Sheets("���ۻ��ڱ�������").Visible = False
        Sheets("��ҵҪ��-ָ������ݱ�").Visible = False
        Sheets("�γ�Ŀ���ɶȻ���������").Visible = False
        Sheets("��ҵҪ���ɶȻ���������").Visible = False
        '1-�Ծ�ɼ��ǼǱ���д���������е����гɼ���ƽʱ�ɼ���ʵ��ɼ������ò��飬�γ̱���������
        '1-�Ծ�ɼ��ǼǱ���д��������Ŀ��˳ɼ�������༭������ɾ����ʽ
        
        Workbooks(WorkBookName).Activate
        Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("R1:X1").Select
        Selection.UnMerge
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Columns("V:Y").Select
        Selection.EntireColumn.Hidden = True
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("R1:X1").Select
        Selection.Merge
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("N4:N" & MaxLineCout).Select
        Selection.Locked = True
        Selection.FormulaHidden = False
        With Selection.Font
          .Name = "Calibri"
          .FontStyle = "�Ӵ�"
          .Size = 11
          .Strikethrough = False
          .Superscript = False
          .Subscript = False
          .OutlineFont = False
          .Shadow = False
          .Underline = xlUnderlineStyleNone
          .Color = 255
          .TintAndShade = 0
          .ThemeFont = xlThemeFontNone
        End With
        With Selection.Interior
          .Pattern = xlSolid
          .PatternColorIndex = xlAutomatic
          .ThemeColor = xlThemeColorDark1
          .TintAndShade = -0.14996795556505
          .PatternTintAndShade = 0
        End With
        'ͳ������༭�����������ȫ��ɾ��
        AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
        If (AllowEditCount <> 0) Then
            For i = 1 To AllowEditCount
                Sheets("1-�Ծ�ɼ��ǼǱ���д��").Protection.AllowEditRanges(1).Delete
            Next i
        End If

        On Error Resume Next
        ActiveSheet.Protection.AllowEditRanges.Add Title:="��������", Range:=Range("E2:M2")
        ActiveSheet.Protection.AllowEditRanges.Add Title:="���Դ������", Range:=Range("E3:M" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="���˳ɼ�", Range:=Range("N4:N" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="�����ɼ�", Range:=Range("O4:R" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="��֤��ʽ", Range:=Range("Z2")
        Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("N4").Select
        ActiveCell.FormulaR1C1 = _
          "=IF(OR(SUM(RC[-9]:RC[-1])=0,RC2=""""),"""",SUM(RC[-9]:RC[-1]))"
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("N4").Select
        Selection.AutoFill Destination:=Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("N4:N" & MaxLineCout), Type:=xlFillDefault
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("N4:N" & MaxLineCout).Select
        Call �Ծ�ɼ��ǼǱ���Ĺ�ʽ
        
        Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Application.ScreenUpdating = True
End Sub

Sub ��֤δ�ύ�ɼ�()
        Dim WorkBookName As String
        Dim Msg As String
        Dim SchoolName As String
        Application.ScreenUpdating = False
        Worksheets("רҵ����״̬").Visible = True
        Worksheets("רҵ����״̬").Activate
        SchoolName = Range("B2").Value
        Worksheets("רҵ����״̬").Visible = False
        Call ���ñ������
        Sheets("2-�γ�Ŀ����ۺϷ�������д��").Select
        '�ָ�"2-�γ�Ŀ����ۺϷ�������д��"��ɶ����۲��ֱ��
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("2-�γ�Ŀ����ۺϷ�������д��").Rows("10:28").Select
        Selection.EntireRow.Hidden = False
        Call �γ�Ŀ������༭����
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = True
        Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Rows("11:20").Select
        Selection.EntireRow.Hidden = False
        SumRow = Application.WorksheetFunction.CountA(Range("B11:B20"))
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Rows(SumRow + 11 & ":20").Select
        Selection.EntireRow.Hidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        '�ָ�3-��ҵҪ�����ݱ���д����1-�γ�Ŀ���ɶ����ۣ���ӡ����2-��ҵҪ���ɶ����ۣ���ӡ����3-�ۺϷ�������ӡ��������

        Sheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Visible = True
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Visible = True
        Sheets("4-�����������棨��д+��ӡ��").Visible = True
        Sheets("�ɼ��˶�").Visible = True
        If SchoolName = "�������Ϣ�밲ȫѧԺ" Then
            Sheets("3-��ҵҪ�����ݱ���д��").Visible = False
        Else
            Sheets("3-��ҵҪ�����ݱ���д��").Visible = True
        End If
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = False
        Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
        Sheets("3-�ۺϷ�������ӡ��").Visible = False
        
        Sheets("�ɼ�¼��").Visible = False
        Sheets("��ѧ���̵ǼǱ�").Visible = False
        Sheets("ѧ������").Visible = False
        Sheets("����ʵ��ɼ���").Visible = False
        Sheets("ƽʱ�ɼ���").Visible = False
        Sheets("�ɼ���").Visible = False
        Sheets("���ۻ��ڱ�������").Visible = False
        Sheets("��ҵҪ��-ָ������ݱ�").Visible = False
        Sheets("�γ�Ŀ���ɶȻ���������").Visible = False
        Sheets("��ҵҪ���ɶȻ���������").Visible = False
        
        Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password

        '1-�Ծ�ɼ��ǼǱ���д��������Ŀ��˳ɼ��в�����༭
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("N4:N" & MaxLineCout).Select
        Selection.Locked = True
        Selection.FormulaHidden = True
        'ͳ������༭�����������ȫ��ɾ��
        AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
        If (AllowEditCount <> 0) Then
            For i = 1 To AllowEditCount
                Sheets("1-�Ծ�ɼ��ǼǱ���д��").Protection.AllowEditRanges(1).Delete
            Next i
        End If

        On Error Resume Next
        ActiveSheet.Protection.AllowEditRanges.Add Title:="��������", Range:=Range("E2:M2")
        ActiveSheet.Protection.AllowEditRanges.Add Title:="���Դ������", Range:=Range("E3:M" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="��֤��ʽ", Range:=Range("Z2")
        With Selection.Font
          .Name = "Calibri"
          .FontStyle = "�Ӵ�"
          .Size = 11
          .Strikethrough = False
          .Superscript = False
          .Subscript = False
          .OutlineFont = False
          .Shadow = False
          .Underline = xlUnderlineStyleNone
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .ThemeFont = xlThemeFontNone
        End With
        With Selection.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Columns("O:S").Select
        Selection.EntireColumn.Hidden = False
        Call �Ծ�ɼ��ǼǱ���Ĺ�ʽ
        Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Application.ScreenUpdating = True
End Sub

Sub ��֤���ύ�ɼ�()
    Dim WorkBookName As String
    Dim Msg As String
    Dim SchoolName As String
        Application.ScreenUpdating = False
        Worksheets("רҵ����״̬").Visible = True
        Worksheets("רҵ����״̬").Activate
        SchoolName = Range("B2").Value
        Worksheets("רҵ����״̬").Visible = False
        Call ���ñ������
        Sheets("2-�γ�Ŀ����ۺϷ�������д��").Select
        '�ָ�"2-�γ�Ŀ����ۺϷ�������д��"��ɶ����۲��ֱ��
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("2-�γ�Ŀ����ۺϷ�������д��").Rows("10:28").Select
        Selection.EntireRow.Hidden = False
        Call �γ�Ŀ������༭����
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = True
        Worksheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Rows("11:20").Select
        Selection.EntireRow.Hidden = False
        SumRow = Application.WorksheetFunction.CountA(Range("B11:B20"))
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Rows(SumRow + 11 & ":20").Select
        Selection.EntireRow.Hidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        '�ָ�3-��ҵҪ�����ݱ���д����1-�γ�Ŀ���ɶ����ۣ���ӡ����2-��ҵҪ���ɶ����ۣ���ӡ����3-�ۺϷ�������ӡ��������
        Sheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Visible = False
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Visible = True
        Sheets("4-�����������棨��д+��ӡ��").Visible = False
        Sheets("�ɼ��˶�").Visible = False
        If SchoolName = "�������Ϣ�밲ȫѧԺ" Then
            Sheets("3-��ҵҪ�����ݱ���д��").Visible = False
        Else
            Sheets("3-��ҵҪ�����ݱ���д��").Visible = True
        End If
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = False
        Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
        Sheets("3-�ۺϷ�������ӡ��").Visible = False
        
        Sheets("�ɼ�¼��").Visible = False
        Sheets("��ѧ���̵ǼǱ�").Visible = False
        Sheets("ѧ������").Visible = False
        Sheets("����ʵ��ɼ���").Visible = False
        Sheets("ƽʱ�ɼ���").Visible = False
        Sheets("�ɼ���").Visible = False
        Sheets("���ۻ��ڱ�������").Visible = False
        Sheets("��ҵҪ��-ָ������ݱ�").Visible = False
        Sheets("�γ�Ŀ���ɶȻ���������").Visible = False
        Sheets("��ҵҪ���ɶȻ���������").Visible = False
        Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password

        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Columns("O:S").Select
        Selection.EntireColumn.Hidden = False
        
        '1-�Ծ�ɼ��ǼǱ���д��������Ŀ��˳ɼ��в�����༭
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Range("N4:N" & MaxLineCout).Select
        Selection.Locked = True
        Selection.FormulaHidden = True
        'ͳ������༭�����������ȫ��ɾ��
        AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
        If (AllowEditCount <> 0) Then
            For i = 1 To AllowEditCount
                Sheets("1-�Ծ�ɼ��ǼǱ���д��").Protection.AllowEditRanges(1).Delete
            Next i
        End If

        On Error Resume Next
        ActiveSheet.Protection.AllowEditRanges.Add Title:="��������", Range:=Range("E2:M2")
        ActiveSheet.Protection.AllowEditRanges.Add Title:="���Դ������", Range:=Range("E3:M" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="��֤��ʽ", Range:=Range("Z2")
        With Selection.Font
          .Name = "Calibri"
          .FontStyle = "�Ӵ�"
          .Size = 11
          .Strikethrough = False
          .Superscript = False
          .Subscript = False
          .OutlineFont = False
          .Shadow = False
          .Underline = xlUnderlineStyleNone
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .ThemeFont = xlThemeFontNone
        End With
        With Selection.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        Sheets("1-�Ծ�ɼ��ǼǱ���д��").Columns("O:S").Select
        Selection.EntireColumn.Hidden = False
        Call �Ծ�ɼ��ǼǱ���Ĺ�ʽ
        
        Worksheets("1-�Ծ�ɼ��ǼǱ���д��").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Application.ScreenUpdating = True
End Sub

Sub ��ѧ���̵ǼǱ�()
Dim SumCount As Integer
Dim Xueqi As String
Dim AllowEditCount As Integer
Dim ImportStatus As Boolean
    Application.ScreenUpdating = False
    Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
    Xueqi = Range("AN1").Value
    SumCount = Range("AM1").Value
    If (Application.WorksheetFunction.Count(Range("A6:A200")) <> SumCount) And SumCount <> 0 Then
        Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
        Call ���ý�ѧ���̵ǼǱ�
    ElseIf (SumCount = 0) Then
        If (Xueqi = "") Then
           Exit Sub
        Else
            ImportStatus = �����ѧ���̵ǼǱ�
            If ImportStatus = False Then
                Application.ScreenUpdating = True
                Exit Sub
            End If
            Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
            If (Range("AM1").Value <> 0) Then
                Call ���ý�ѧ���̵ǼǱ�
            End If
        End If
    End If
    Application.ScreenUpdating = True
End
Sub ����������ɫ(SetSheetName As String, SetRange As String, SetColor As String)
    Worksheets(SetSheetName).Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range(SetRange).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = SetColor
        .TintAndShade = -0.14996795556505
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
'[�汾��]V5.06.34


