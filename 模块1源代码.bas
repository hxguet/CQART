Attribute VB_Name = "ģ��1"
    Public Const MaxRecord As String = "185"
    Public Const OldPassword As String = "dpt8hx"
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
    Public School As String
    Public MajorCount As Integer
    Public MajorList(4) As String
    Public ModuleLastRivise() As String
''רҵ���޶�
Sub �޶���ʽ()
    Application.MacroOptions Macro:="�޶���ʽ", Description:="", ShortcutKey:="q"
    Application.EnableEvents = False
    Call ���ñ������
    Call ���ĵ�
    Call �������ù�ʽ��ť
    Call �޶�רҵ����״̬
    Call �޶��γ�Ŀ����ۺϷ�����ʽ
    Call �޶�ƽʱ�ɼ���
    Call �޶���ѧ���̵ǼǱ�ʽ
    Call �޶���ҵҪ���ɶ����۱�
    Call �����¼�����
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Application.EnableEvents = True
    ActiveWorkbook.Save
End Sub
Sub ���������()
    Dim temp As Boolean
    For Each sht In Sheets
        temp = Worksheets(sht.Name).Visible
        Worksheets(sht.Name).Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=OldPassword
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets(sht.Name).Visible = temp
    Next
End Sub
Sub ���ɰ汾��()
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
    BatFile = "��������.bat"
    Commit = Format(Now, "yyyy-mm-dd hh:mm:ss  Commit")
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
    If (RiviseVer < 40) Then
        RiviseVer = RiviseVer + 1
    Else
        Range("H5").Value = 1
        If (SubVer < 20) Then
            SubVer = SubVer + 1
        Else
            Range("H4").Value = 1
            If (MainVer < 10) Then
                MainVer = MainVer + 1
             End If
       End If
    End If
    Version = "V" & MainVer & "." & Format(SubVer, "00") & "." & Format(RiviseVer, "00")
    RiviseDate = Format(Now, "yyyy-mm-dd")
    Commit = Format(Now, "yyyy-mm-dd hh:mm:ss") & "Commit"
    ModuleFileName = BackupFilePath & "\ģ��1Դ����-" & Version & "-" & Format(RiviseDate, "YYYYMMDD") & ".bas"
    If Dir(ModuleFileName) <> "" Then
        Kill ModuleFileName
    End If
    If Dir(ReleaseFilePath & "\") = "" Then
        MkDir ReleaseFilePath
    End If
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
        Application.VBE.ActiveVBProject.VBComponents("ģ��1").Export (ModuleFileName)
        If Dir(ReleaseFilePath & "\" & ReleaseFile) <> "" Then
            Kill ReleaseFilePath & "\" & ReleaseFile
        End If
        Application.VBE.ActiveVBProject.VBComponents("ģ��1").Export (ReleaseFilePath & "\" & ReleaseFile)
        Range("H1").Select
        ActiveCell.FormulaR1C1 = _
            "=""V""&R[2]C&"".""&TEXT(R[3]C,""00"")&"".""&TEXT(R[4]C,""00"")"
        Range("H2").Value = RiviseDate
        Range("H3").Value = MainVer
        Range("H4").Value = SubVer
        Range("H5").Value = RiviseVer
        Call ����Readme(ReleaseFilePath, "Readme.txt", BackupFilePath, "Readme.txt", "ģ��1", ReleaseFile, Version, RiviseDate)
        '����Git�����������ļ�
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set MyTxtObj = fso.CreateTextFile(ReleaseFilePath & "\" & BatFile, True, False)
        MyTxtObj.WriteLine (Mid(ReleaseFilePath, 1, 2))
        MyTxtObj.WriteLine ("cd " & Mid(ReleaseFilePath, 3, Len(ReleaseFilePath) - 2))
        MyTxtObj.WriteLine ("git add .")
        TempStr = "git commit -m """
        TempStr = TempStr & Commit
        TempStr = TempStr & """"
        MyTxtObj.WriteLine (TempStr)
        MyTxtObj.WriteLine ("git push -u origin master")
        MyTxtObj.WriteLine ("exit")
        MyTxtObj.Close
        Shell (ReleaseFilePath & "\" & BatFile)
    End If
    Set MyTxtObj = Nothing
    Set fso = Nothing
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
End Sub
Sub ����Readme(ReleaseFilePath As String, ReleaseReadmeFile As String, BackupFilePath As String, BackupReadmeFile As String, ModuleName As String, ReleaseFile As String, Version As String, RiviseDate As String)
    Dim ModuleCount As Integer
    Dim i As Integer
    Dim k As Integer
    Dim LineCount As Integer
    Dim UpdateInfo As String
    Dim Status As String
    On Error Resume Next
    Status = DownFile(ReleaseFilePath, ReleaseReadmeFile)
    If Status = False Then
        Exit Sub
    End If
    Call GetVersionFromFile(ReleaseFilePath & "\" & ReleaseReadmeFile)
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
End Sub
Sub Զ�̸��´���()
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
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    LastFilePath = ThisWorkbook.Path
    LastReadme = "Readme.txt"
    If Dir((LastFilePath & "\" & LastReadme)) <> "" Then
        Kill LastFilePath & "\" & LastReadme
    End If
    MsgBox ("��������Զ�̷����������������°汾��")
    Status = DownFile(LastFilePath, LastReadme)
    If Status = False Or Dir(LastFilePath & "\" & LastReadme) <> "" Then
        GoTo Error
    End If
    Call GetVersionFromFile(LastFilePath & "\" & LastReadme)
    ModuleFile = ModuleLastRivise(1, CFileName)
    Status = DownFile(ThisWorkbook.Path, ModuleFile)
    If Status = False Or Dir(ThisWorkbook.Path & "\" & ModuleFile) <> "" Then
        GoTo Error
    End If
    CurrentVersion = Range("H1").Value
    CurrentRiviseDate = Range("H2").Value
    CtrResult = StrComp(CurrentVersion, ModuleLastRivise(1, CVersion), vbTextCompare)
    'Զ�̴���汾�űȵ�ǰ����汾����
    If CtrResult = -1 Then
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
        MsgBox ("�Ѹ��´���Ϊ��" & LastVersion & "�޶����ڣ�" & LastRiviseDate)
    Else
        MsgBox ("��ģ�����汾�Ѿ�Ϊ���°汾!")
    End If
    If Dir(LastFilePath & "\" & LastReadme) <> "" Then
        Kill LastFilePath & "\" & LastReadme
    End If
    If Dir(ThisWorkbook.Path & "\" & ModuleFile) <> "" Then
        Kill ThisWorkbook.Path & "\" & ModuleFile
    End If
    Call �޶���ʽ
Error:
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
End Sub
Function ShellAndWait(cmdStr As String) As String
    On Error Resume Next
    Dim oShell As Object, oExec As Object
    Set oShell = CreateObject("WScript.Shell")
    Set oExec = oShell.exec(cmdStr)
    If Err.Description <> "" Then
        ShellAndWait = Err.Description
    Else
        ShellAndWait = oExec.StdOut.ReadAll
    End If
    Set oShell = Nothing
    Set oExec = Nothing
End Function
Sub WriteLine()
    Dim fso As Object, sFile As Object
    Const ForReading = 1, ForWriting = 2, ForAppending = 8, TristateFalse = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sFile = fso.OpenTextFile("C:\FSOTest\testfile.txt", ForAppending, TristateFalse)
    sFile.WriteLine "WriteLine Test"
    sFile.Close
    Set fso = Nothing
    Set sFile = Nothing
End Sub
Sub GetVersionFromLocal()
    Dim ModuleRivise() As String
    Dim ModuleCount As Integer
    Dim ModuleFile As String
    Dim CurrentVersion As String
    Dim CurrentRiviseDate As String
    Dim wbList() As String
    Dim FileName As String
    Dim FileType As String
    Dim LastVersion As String
    Dim LastRiviseDate As String
    Dim CtrResult As String
    Dim Vbc As Object
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    CurrentVersion = Range("H1").Value
    CurrentRiviseDate = Range("H2").Value
    FolderName = ThisWorkbook.Path
    wbName = Dir(FolderName & "\*.bas")
    ModuleCount = 0
    While wbName <> ""
        Info = Split(Mid(wbName, 1, Len(wbName) - 4), "-")
        '�ļ����������ģ�������汾�ź��޶�����
        If UBound(Info) - LBound(Info) = 2 Then
            If Len(Info(0)) = 6 And Len(Info(1)) = 8 And Len(Info(2)) = 8 Then
                ModuleCount = ModuleCount + 1
                ReDim Preserve wbList(1 To ModuleCount)
                wbList(ModuleCount) = wbName
            End If
        End If
        wbName = Dir
    Wend
    If ModuleCount = 0 Then
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("רҵ����״̬").Visible = False
        Exit Sub
    End If
    LastVersion = ""
    LastRiviseDate = ""
    ReDim Preserve ModuleLastRivise(0 To ModuleCount, 0 To 4)
    ModuleLastRivise(0, CModuleCount) = ModuleCount
    For i = 1 To ModuleCount
        ModuleLastRivise(i, CFileName) = wbList(i)
        Info = Split(Mid(wbList(i), 1, Len(wbList(i)) - 4), "-")
        ModuleLastRivise(i, CModuleName) = Mid(Info(0), 1, 3)
        ModuleLastRivise(i, CVersion) = Info(1)
        ModuleLastRivise(i, CRiviseDate) = Info(2)
    Next i
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
End Sub
Function DownFile(FilePath As String, FileName As String)
    Dim TempFileName As String
    Dim Result As String
    Dim VersionFilePath As String
    RemoteFile = "https://raw.githubusercontent.com/hxguet/CQART/master/" & FileName
    TempFileName = ThisWorkbook.Path & "\wget.exe -O " & FilePath & "\" & FileName & " " & RemoteFile
    Result = ShellAndWait(TempFileName)
    If Result <> "" Then
        MsgBox ("����" & ThisWorkbook.Path & "\wget.exe �ļ��Ƿ���ڣ�")
        DownFile = False
    Else
        DownFile = True
    End If
End Function
Sub GetVersionFromFile(LocalFileName As String)
    Dim StrTxt() As String
    Dim n As Integer
    Dim StrTemp As String
    Dim i As Integer, x As Integer, y As Integer
    Dim Module As String
    Dim ModuleCount As Integer
    Dim UpdateInfo As String
    ModuleCount = 0
    UpdateInfo = ""
    Open LocalFileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, StrTemp
        n = n + 1
    Loop
    x = InStr(1, StrTemp, vbCrLf)
    If x = 0 Then
        StrTxt = Split(StrTemp, vbLf)
        n = UBound(StrTxt) - LBound(StrTxt)
    Else
        ReDim Preserve StrTxt(0 To n - 1)
        Do While Not EOF(1)
            Line Input #1, StrTxt(i)
        Loop
    End If
    Close #1
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
End Sub

Sub �޶��γ�Ŀ����ۺϷ�����ʽ()
    Dim MyShapes As Shapes
    Dim Shp As Shape
    Dim SchoolName As String
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    SchoolName = Range("B2").Value
    Worksheets("רҵ����״̬").Visible = False
    '�޶�"2-�γ�Ŀ����ۺϷ�������д��"���������ۻ��ڿγ̱������ҵ�ɼ���ʽ
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    Application.EnableEvents = False
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Set MyShapes = Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Shapes
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Set NewShp = ActiveSheet.Buttons.Add(866, 66, 77, 28).Select
    NewShp.Characters.Text = "���´���"
    NewShp.OnAction = "���ɰ汾��"
    NewShp.Font.Name = "΢���ź�"
    NewShp.Font.Size = 14
    NewShp.Font.ColorIndex = 3
    '2019.5.3�޶������ʵ��γ�ʵ��1��ʵ��2��ʵ��3������Ϊ100�֣��ϼƿ��˷ֳ���100�ֵ����
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
End Sub
Sub �޶���ѧ���̵ǼǱ�ʽ()
    Dim temp As Boolean
    '"0-��ѧ���̵ǼǱ���д+��ӡ)"�������޶����⣬ѧ�ţ������ؼ��ʣ��������壬��������༭����
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
End Sub
Sub �޶�ƽʱ�ɼ���()
    Worksheets("ƽʱ�ɼ���").Visible = True
    Worksheets("ƽʱ�ɼ���").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC25+RC29=0,0,ROUND((RC32)/(R5C25+R5C29),0))"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("ƽʱ�ɼ���").Visible = False
End Sub
Sub �޶�רҵ����״̬()
     '�޶�רҵ����״̬������
    Dim MyShapes As Shapes
    Dim Shp As Shape
    Set MyShapes = Worksheets("רҵ����״̬").Shapes
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
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
    Range("G1:G5").Select
    Selection.Font.Bold = True
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Set NewShp = ActiveSheet.Buttons.Add(695, 2.25, 102, 27) '��λ�ø߶ȣ�λ�ÿ�ȣ���ť�߶ȣ���ť��ȣ�
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
    Range("B4:C12,B2:F2").Select
    Range("B2").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
        .PatternTintAndShade = 0
    End With
    Range("A7:F7").Select
    Selection.Copy
    Range("A8:F12").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Rows("1:12").Select
    Selection.RowHeight = 30
    Call ���ñ����("A2", "AF12", 12)
    Range("A1:I12").Select
    With Selection.Validation
        .Delete
    End With
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "���뷢��·��"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "���뱸��·��"
    Range("G5").Select
    Selection.Copy
    Range("G6:G7").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Columns("H:H").Select
    Selection.ColumnWidth = 20
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("רҵ����״̬").Visible = False
End Sub
Sub �޶���ҵҪ���ɶ����۱�()
    Dim SchoolName As String
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
End Sub
Sub �γ�Ŀ������༭����()
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
Dim Grade As String
Dim Major As String
Dim i As Integer
Dim CourseCount As Integer
Dim PointCount As Integer
Dim MatrixSheet As String
On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    MsgBox ("���µ����ѧ����ѧ��������רҵ�������Ϣ�����Եȡ�����")
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
    Call ����ѧԺרҵ
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
    Application.ScreenUpdating = True
End Sub
Sub �������()
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
        If (Range("E" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value <> CourseCount) Or (Range("F" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value <> PointCount) Then
            Call �����ҵҪ�����(Major)
        End If
    End If
    Sheets("רҵ����״̬").Visible = False
End Sub
Sub �����¼�����()
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
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub ��ӡ()
' ���ø�ʽ ��
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
    CurrentWorksheet = ActiveSheet.Name
    Application.ScreenUpdating = False
    Worksheets("רҵ����״̬").Visible = True
    Worksheets("רҵ����״̬").Activate
    SchoolName = Range("B2").Value
    ErrNum = 0
    ErrorMsg = ""
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    CourseTargetCount = Application.CountA(Range("B11:B20"))
    If CourseTargetCount = 0 Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д��������δ��д���γ�Ŀ�꡿" & vbCr & vbLf
    End If
    '���ĵ����
    DateCompleted = Range("$B$8").Value
    Term = Range("$B$2").Value
    Term = Mid(Term, 3, 2) & "-" & Mid(Term, 8, 2) & "-" & Mid(Term, 14, 1)
    CourseNum = Range("$B$3").Value
    CourseName = Range("$B$4").Value
    Teacher = Range("$B$5").Value
    Major = Range("$B$7").Value
    PDFFileName = Term & "-" & CourseNum & "-" & Major & "-" & Teacher & "-" & CourseName
    If DateCompleted = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д��������δ��д����д���ڡ�" & vbCr & vbLf
    End If
    Worksheets("�γ�Ŀ���ɶȻ���������").Visible = True
    Worksheets("�γ�Ŀ���ɶȻ���������").Activate
    If Range("D2").Value = 0 Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "��2-�γ�Ŀ����ۺϷ�������д��������δ��д���γ���š�" & vbCr & vbLf
    ElseIf (Range("B2").Value = "") Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "���γ���Ϣ��ȡ����������������Դ-��ѧ����.xls�в��䡾ѧ�ڡ�" & vbCr & vbLf
    ElseIf Range("C2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "���γ���Ϣ��ȡ����������������Դ-��ѧ����.xls�в��䡾�γ����ơ�" & vbCr & vbLf
    ElseIf Range("E2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "���γ���Ϣ��ȡ����������������Դ-��ѧ����.xls�в��䡾������ʦ��" & vbCr & vbLf
    ElseIf Range("G2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "���γ���Ϣ��ȡ����������������Դ-��ѧ����.xls�в��䡾ѧ�֡�" & vbCr & vbLf
    End If
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    LinkCount = Application.CountA(Range("D7:Q7")) - Application.CountBlank(Range("D7:Q7"))
    For i = 4 To LinkCount + 4
        If Cells(7, i).Value <> "" Then
            If Application.CountBlank(Cells(11, i).Resize(10, 1)) = 10 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "�����ۻ���" & Cells(5, i).Value & "��ƽ���ɼ����������ۻ���δ֧�ſγ�Ŀ�ꡣ" & vbCr & vbLf
            End If
        Else
            If Application.CountBlank(Cells(11, i).Resize(10, 1)) <> 10 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "�����ۻ���" & Cells(5, i).Value & "û��ƽ���ɼ����������ۻ���֧���˿γ�Ŀ�ꡣ" & vbCr & vbLf
            End If
        End If
    Next i
    
    For i = 0 To CourseTargetCount - 1
        Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
        If Range("R" & i + 11).Value <> 100 Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "���γ�Ŀ��" & i + 1 & "֧�ű����ϼƲ�Ϊ100��" & vbCr & vbLf
        End If
        Worksheets("�γ�Ŀ���ɶȻ���������").Activate
        If (Not IsNumeric(Cells(2, 2 * i + 9).Value) Or Cells(2, 2 * i + 9).Value = "" Or Cells(2, 2 * i + 9).Value = 0) Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "���γ�Ŀ��" & i + 1 & "����������ȷ���顣" & vbCr & vbLf
        End If
    Next i
    Worksheets("�γ�Ŀ���ɶȻ���������").Visible = False
    Worksheets("3-��ҵҪ�����ݱ���д��").Activate
    RequirementCount = Application.CountA(Range("C7:C18")) - Application.CountBlank(Range("C7:C18"))
    RequirementReachCount = Application.Count(Range("D7:D18"))

    For i = 0 To Application.CountA(Range("B7:B18")) - 1
        If Range("O" & i + 11).Value <> "" And Range("O" & i + 11).Value <> 100 Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "��" & Cells(i + 7, 2).Value & "֧�ű����ϼƲ�Ϊ100��" & vbCr & vbLf
        End If
        If (Cells(i + 7, 3).Value = "��") And (Not IsNumeric(Cells(i + 7, 4).Value) Or Cells(i + 7, 4).Value = "" Or Cells(i + 7, 4).Value = 0) Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "��" & Cells(i + 7, 2).Value & "����������ȷ" & vbCr & vbLf
        End If
    Next i
    If ErrorMsg <> "" Then
        MsgBox (ErrorMsg)
        Call CreateTXTfile(ErrorMsg)
        Worksheets("�γ�Ŀ���ɶȻ���������").Visible = False
        Worksheets("��ҵҪ���ɶȻ���������").Visible = False
        Exit Sub
    End If
    Call ��������ʽ
    Worksheets("2-�γ�Ŀ����ۺϷ�������д��").Activate
    If Range("$Q$3").Value = "����֤" Then
        Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Call ȡ��������ɫ
        
        
        Call Excel2PDF("0-��ѧ���̵ǼǱ���д+��ӡ)", ThisWorkbook.Path, PDFFileName & "--��ѧ���̵ǼǱ�.pdf")
        Call ���ø�����ɫ
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        Worksheets("4-�����������棨��д+��ӡ��").Activate
        Call ȡ����������������д������ɫ
        
        Call Excel2PDF("4-�����������棨��д+��ӡ��", ThisWorkbook.Path, PDFFileName & "--������������.pdf")
        Call ������������������д������ɫ
    ElseIf Range("$Q$3").Value = "��֤δ�ύ�ɼ�" Then
        '��ӡ��ѧ���̵ǼǱ�
        Worksheets("0-��ѧ���̵ǼǱ���д+��ӡ)").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Call ȡ��������ɫ
        
        
        Call Excel2PDF("0-��ѧ���̵ǼǱ���д+��ӡ)", ThisWorkbook.Path, PDFFileName & "--��ѧ���̵ǼǱ�.pdf")
        Call ���ø�����ɫ
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
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
        
        Call Excel2PDF("1-�γ�Ŀ���ɶ����ۣ���ӡ��", ThisWorkbook.Path, PDFFileName & "--�γ�Ŀ�����������.pdf")
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = False

        Select Case SchoolName
            Case "�������Ϣ�밲ȫѧԺ"
                Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
                Sheets("3-�ۺϷ�������ӡ��").Visible = True
                Worksheets("3-�ۺϷ�������ӡ��").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                
                
                Call Excel2PDF("3-�ۺϷ�������ӡ��", ThisWorkbook.Path, PDFFileName & "--�γ��ۺϷ���.pdf")
                Sheets("3-�ۺϷ�������ӡ��").Visible = False
                
                Worksheets("4-�����������棨��д+��ӡ��").Activate
                Call ȡ����������������д������ɫ
                
                Call Excel2PDF("4-�����������棨��д+��ӡ��", ThisWorkbook.Path, PDFFileName & "--������������.pdf")
                
                Call ������������������д������ɫ
            Case "���ӹ������Զ���ѧԺ"
                Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = True
                Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                
                Call Excel2PDF("2-��ҵҪ���ɶ����ۣ���ӡ��", ThisWorkbook.Path, PDFFileName & "--��ҵҪ�����������.pdf")
                Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
                Sheets("3-�ۺϷ�������ӡ��").Visible = True
                Worksheets("3-�ۺϷ�������ӡ��").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                
                
                Call Excel2PDF("3-�ۺϷ�������ӡ��", ThisWorkbook.Path, PDFFileName & "--�γ��ۺϷ���.pdf")
                Sheets("3-�ۺϷ�������ӡ��").Visible = False
                
                Worksheets("4-�����������棨��д+��ӡ��").Activate
                Call ȡ����������������д������ɫ
                
                Call Excel2PDF("4-�����������棨��д+��ӡ��", ThisWorkbook.Path, PDFFileName & "--������������.pdf")
                
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
        
        Call Excel2PDF("1-�γ�Ŀ���ɶ����ۣ���ӡ��", ThisWorkbook.Path, PDFFileName & "--�γ�Ŀ�����������.pdf")
        Sheets("1-�γ�Ŀ���ɶ����ۣ���ӡ��").Visible = False
        Select Case SchoolName
            Case "�������Ϣ�밲ȫѧԺ"
                Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
                Worksheets("3-�ۺϷ�������ӡ��").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                Call Excel2PDF("3-�ۺϷ�������ӡ��", ThisWorkbook.Path, PDFFileName & "--�γ��ۺϷ���.pdf")
                Sheets("3-�ۺϷ�������ӡ��").Visible = False
            Case "���ӹ������Զ���ѧԺ"
                Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = True
                Worksheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                Call Excel2PDF("2-��ҵҪ���ɶ����ۣ���ӡ��", ThisWorkbook.Path, PDFFileName & "--��ҵҪ�����������.pdf")
                Sheets("2-��ҵҪ���ɶ����ۣ���ӡ��").Visible = False
                Sheets("3-�ۺϷ�������ӡ��").Visible = True
                Worksheets("3-�ۺϷ�������ӡ��").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                Call Excel2PDF("3-�ۺϷ�������ӡ��", ThisWorkbook.Path, PDFFileName & "--�γ��ۺϷ���.pdf")
                Sheets("3-�ۺϷ�������ӡ��").Visible = False
            End Select
    Else
        MsgBox ("2-�γ�Ŀ����ۺϷ�������д���������еġ��Ƿ���֤��δѡ��")
    End If
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    Worksheets(CurrentWorksheet).Activate
End Sub
Sub ��ҵҪ�����ݱ�ʽ()
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
                F
