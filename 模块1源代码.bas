Attribute VB_Name = "模块1"
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
''专业及修订
Sub 修订公式()
    Application.MacroOptions Macro:="修订公式", Description:="", ShortcutKey:="q"
    Application.EnableEvents = False
    Call 设置表格主题
    Call 打开文档
    Call 重新设置公式按钮
    Call 修订专业矩阵状态
    Call 修订课程目标和综合分析公式
    Call 修订平时成绩表
    Call 修订教学过程登记表公式
    Call 修订毕业要求达成度评价表
    Call 允许事件触发
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Application.EnableEvents = True
    ActiveWorkbook.Save
End Sub
Sub 工作表加密()
    Dim temp As Boolean
    For Each sht In Sheets
        temp = Worksheets(sht.Name).Visible
        Worksheets(sht.Name).Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=OldPassword
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets(sht.Name).Visible = temp
    Next
End Sub
Sub 生成版本号()
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
    BatFile = "发布代码.bat"
    Commit = Format(Now, "yyyy-mm-dd hh:mm:ss  Commit")
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    MainVer = Range("H3").Value
    SubVer = Range("H4").Value
    RiviseVer = Range("H5").Value
    ReleaseFilePath = Range("H6").Value
    BackupFilePath = Range("H7").Value
    ReleaseFile = "模块1源代码.bas"
    BatFile = "发布代码.bat"
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
    ModuleFileName = BackupFilePath & "\模块1源代码-" & Version & "-" & Format(RiviseDate, "YYYYMMDD") & ".bas"
    If Dir(ModuleFileName) <> "" Then
        Kill ModuleFileName
    End If
    If Dir(ReleaseFilePath & "\") = "" Then
        MkDir ReleaseFilePath
    End If
    ModuleCount = 0
    For Each Vbc In ThisWorkbook.VBProject.VBComponents
        If Vbc.Type = 1 And Mid(Vbc.Name, 1, 2) = "模块" Then
            ModuleCount = ModuleCount + 1
            If (Vbc.Name <> "模块1") Then
                Vbc.Name = "模块1"
            End If
        End If
    Next Vbc
    If ModuleCount = 1 Then
        Application.VBE.ActiveVBProject.VBComponents("模块1").Export (ModuleFileName)
        If Dir(ReleaseFilePath & "\" & ReleaseFile) <> "" Then
            Kill ReleaseFilePath & "\" & ReleaseFile
        End If
        Application.VBE.ActiveVBProject.VBComponents("模块1").Export (ReleaseFilePath & "\" & ReleaseFile)
        Range("H1").Select
        ActiveCell.FormulaR1C1 = _
            "=""V""&R[2]C&"".""&TEXT(R[3]C,""00"")&"".""&TEXT(R[4]C,""00"")"
        Range("H2").Value = RiviseDate
        Range("H3").Value = MainVer
        Range("H4").Value = SubVer
        Range("H5").Value = RiviseVer
        Call 生成Readme(ReleaseFilePath, "Readme.txt", BackupFilePath, "Readme.txt", "模块1", ReleaseFile, Version, RiviseDate)
        '生成Git发布批处理文件
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
    Worksheets("专业矩阵状态").Visible = False
End Sub
Sub 生成Readme(ReleaseFilePath As String, ReleaseReadmeFile As String, BackupFilePath As String, BackupReadmeFile As String, ModuleName As String, ReleaseFile As String, Version As String, RiviseDate As String)
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
    ModuleLastRivise(1, CModuleName) = "模块1"
    ModuleLastRivise(1, CFileName) = ReleaseFile
    ModuleLastRivise(1, CVersion) = Version
    ModuleLastRivise(1, CRiviseDate) = RiviseDate
    For i = 1 To ModuleCount
        MyTxtObj.WriteLine ("[模块名称]" & ModuleLastRivise(i, CModuleName))
        MyTxtObj.WriteLine ("[文件名称]" & ModuleLastRivise(i, CFileName))
        MyTxtObj.WriteLine ("[修订版本]" & ModuleLastRivise(i, CVersion))
        MyTxtObj.WriteLine ("[修订日期]" & ModuleLastRivise(i, CRiviseDate))
        UpdateInfo = InputBox("请输入" & ModuleLastRivise(i, CModuleName) & "此次更新说明" & vbCrLf & ModuleLastRivise(i, CUpdateInfo))
        MyTxtObj.WriteLine ("[更新说明]" & vbCrLf & ModuleLastRivise(i, CUpdateInfo)) & UpdateInfo
    Next i
    fso.CopyFile ReleaseFilePath & "\" & ReleaseReadmeFile, BackupFilePath & "\" & BackupReadmeFile
    MyTxtObj.Close
    Set fso = Nothing
    Set MyTxtObj = Nothing
End Sub
Sub 远程更新代码()
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
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    LastFilePath = ThisWorkbook.Path
    LastReadme = "Readme.txt"
    If Dir((LastFilePath & "\" & LastReadme)) <> "" Then
        Kill LastFilePath & "\" & LastReadme
    End If
    MsgBox ("正在连接远程服务器，检查代码最新版本！")
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
    '远程代码版本号比当前代码版本号新
    If CtrResult = -1 Then
        ModuleName = ModuleLastRivise(1, CModuleName)
        ModuleFile = ModuleLastRivise(1, CFileName)
        LastVersion = ModuleLastRivise(1, CVersion)
        LastRiviseDate = ModuleLastRivise(1, CRiviseDate)
        UpdateInfo = ModuleLastRivise(1, CUpdateInfo)
        For Each Vbc In ThisWorkbook.VBProject.VBComponents
            If Vbc.Type = 1 And Mid(Vbc.Name, 1, 2) = "模块" Then
                ThisWorkbook.VBProject.VBComponents.Remove Vbc
            End If
        Next Vbc
        If Dir(ThisWorkbook.Path & "\" & ModuleFile) <> "" Then
            ActiveWorkbook.VBProject.VBComponents.Import ThisWorkbook.Path & "\" & ModuleFile
            ModuleCount = 0
            For Each Vbc In ThisWorkbook.VBProject.VBComponents
                If Vbc.Type = 1 And Mid(Vbc.Name, 1, 2) = "模块" Then
                    ModuleCount = ModuleCount + 1
                End If
                If ModuleCount = 1 Then
                    Vbc.Name = "模块1"
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
        MsgBox ("已更新代码为：" & LastVersion & "修订日期：" & LastRiviseDate)
    Else
        MsgBox ("该模版代码版本已经为最新版本!")
    End If
    If Dir(LastFilePath & "\" & LastReadme) <> "" Then
        Kill LastFilePath & "\" & LastReadme
    End If
    If Dir(ThisWorkbook.Path & "\" & ModuleFile) <> "" Then
        Kill ThisWorkbook.Path & "\" & ModuleFile
    End If
    Call 修订公式
Error:
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("专业矩阵状态").Visible = False
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
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    CurrentVersion = Range("H1").Value
    CurrentRiviseDate = Range("H2").Value
    FolderName = ThisWorkbook.Path
    wbName = Dir(FolderName & "\*.bas")
    ModuleCount = 0
    While wbName <> ""
        Info = Split(Mid(wbName, 1, Len(wbName) - 4), "-")
        '文件名必须包含模块名，版本号和修订日期
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
        Worksheets("专业矩阵状态").Visible = False
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
    Worksheets("专业矩阵状态").Visible = False
End Sub
Function DownFile(FilePath As String, FileName As String)
    Dim TempFileName As String
    Dim Result As String
    Dim VersionFilePath As String
    RemoteFile = "https://raw.githubusercontent.com/hxguet/CQART/master/" & FileName
    TempFileName = ThisWorkbook.Path & "\wget.exe -O " & FilePath & "\" & FileName & " " & RemoteFile
    Result = ShellAndWait(TempFileName)
    If Result <> "" Then
        MsgBox ("请检查" & ThisWorkbook.Path & "\wget.exe 文件是否存在！")
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
        x = InStr(1, StrTxt(i), "[模块名称]")
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
        If InStr(1, StrTxt(i), "[模块名称]") > 0 Then
            ModuleLastRivise(k, CStartLine) = i + 1
            ModuleLastRivise(k, CModuleName) = Replace(StrTxt(i), "[模块名称]", "")
        ElseIf InStr(1, StrTxt(i), "[文件名称]") > 0 Then
            ModuleLastRivise(k, CFileName) = Replace(StrTxt(i), "[文件名称]", "")
        ElseIf InStr(1, StrTxt(i), "[修订版本]") > 0 Then
            ModuleLastRivise(k, CVersion) = Replace(StrTxt(i), "[修订版本]", "")
        ElseIf InStr(1, StrTxt(i), "[修订日期]") > 0 Then
            ModuleLastRivise(k, CRiviseDate) = Replace(StrTxt(i), "[修订日期]", "")
        ElseIf InStr(1, StrTxt(i), "[更新说明]") > 0 Then
            i = i + 1
            Do While i <= n - 1
                If InStr(1, StrTxt(i), "[文件名称]") = 0 Then
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

Sub 修订课程目标和综合分析公式()
    Dim MyShapes As Shapes
    Dim Shp As Shape
    Dim SchoolName As String
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    SchoolName = Range("B2").Value
    Worksheets("专业矩阵状态").Visible = False
    '修订"2-课程目标和综合分析（填写）"工作表评价环节课程报告和作业成绩公式
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Application.EnableEvents = False
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Set MyShapes = Worksheets("2-课程目标和综合分析（填写）").Shapes
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Set NewShp = ActiveSheet.Buttons.Add(866, 66, 77, 28).Select
    NewShp.Characters.Text = "更新代码"
    NewShp.OnAction = "生成版本号"
    NewShp.Font.Name = "微软雅黑"
    NewShp.Font.Size = 14
    NewShp.Font.ColorIndex = 3
    '2019.5.3修订，解决实验课程实验1，实验2，实验3等满分为100分，合计考核分超过100分的情况
    Range("R7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(COUNTBLANK(RC[-14]:RC[-1])=14,SUM(R6C4:R6C17)=0),"""",SUM('2-课程目标和综合分析（填写）'!R7C4:R7C17)*100/SUM(R6C4:R6C17))"
    Range("R8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R7C18="""",R7C18=0),"""",ROUND(R7C18*100/R6C18,1))"
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Call 课程目标允许编辑区域
    Range("N5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-试卷成绩登记表（填写）'!R[-3]C[2]="""","""",'1-试卷成绩登记表（填写）'!R[-3]C[2])"
    Range("Q5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-试卷成绩登记表（填写）'!R[-3]C[2]="""","""",'1-试卷成绩登记表（填写）'!R[-3]C[2])"
    If SchoolName = "计算机信息与安全学院" Then
        If Range("A26").Value = "（4）毕业要求达成度评价" Then
            Rows("26:27").Select
            Selection.EntireRow.Hidden = True
            Range("A28").Select
            ActiveCell.FormulaR1C1 = "（4）改进措施"
        End If
    End If
    Range("AH4:AQ4").Select
    Selection.Merge
    Range("AH5:AQ5").Select
    Selection.Merge
    Range("AH4:AQ4").Select
    ActiveCell.FormulaR1C1 = "=专业矩阵状态!R[-3]C[-27]&""：""&专业矩阵状态!R[-3]C[-26]"
    Range("AH5:AQ5").Select
    ActiveCell.FormulaR1C1 = _
        "=专业矩阵状态!R[-3]C[-27]&""：""&TEXT(专业矩阵状态!R[-3]C[-26],""YYYY年MM月DD日"")"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.EnableEvents = True
End Sub
Sub 修订教学过程登记表公式()
    Dim temp As Boolean
    '"0-教学过程登记表（填写+打印)"工作表修订标题，学号，姓名关键词，设置字体，设置允许编辑区域
    temp = Worksheets("0-教学过程登记表（填写+打印)").Visible
    Worksheets("0-教学过程登记表（填写+打印)").Visible = True
    Worksheets("0-教学过程登记表（填写+打印)").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A2:AG2").Select
    ActiveCell.FormulaR1C1 = _
        "=专业矩阵状态!RC[1]&"" ""&'2-课程目标和综合分析（填写）'!R[2]C[1]&"" 课程(考试/考查/选修)教学过程登记表"""
    Range("B4:B5").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "学号"
    Range("C4:C5").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "姓名"
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
            .Name = "宋体"
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
    Call 设置各行颜色
    '教学过程登记表允许编辑区域设置
    AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
    If (AllowEditCount <> 0) Then
        For i = 1 To AllowEditCount
            Sheets("0-教学过程登记表（填写+打印)").Protection.AllowEditRanges(1).Delete
        Next i
    End If
    ActiveSheet.Protection.AllowEditRanges.Add Title:="记录区", Range:=Range("$D$6:$Y$185")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="期中成绩", Range:=Range("$AA$6:$AA$185")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="成绩等级区", Range:=Range("$AV$6:$AW$21")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="成绩类别区", Range:=Range("$AF$6:$AF$185")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="考试类别区", Range:=Range("$AG$6:$AG$185")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="最后行数", Range:=Range("$AJ$1")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="评价方式", Range:=Range("$AX$6")
    Range("$D$6:$Y$185").Select
    Selection.FormulaHidden = False
    Range("$AA$6:$AA$185").Select
    Selection.FormulaHidden = False
    
    Range("AT20:AU20").Select
    ActiveCell.FormulaR1C1 = "点名迟到"
    Range("AV20").Select
    ActiveCell.FormulaR1C1 = "□"
    Range("AW20").Value = 70
    Range("AT21:AU21").Select
    ActiveCell.FormulaR1C1 = "点名请假"
    Range("AV21").Select
    ActiveCell.FormulaR1C1 = "☆"
    Range("AW21").Value = 80
    Range("AT20:AU20").Select
    Selection.Merge
    Range("AT21:AU21").Select
    Selection.Merge
    Call 设置表格线("AT6", "AW21", 9)
    Range("AT6:AW21").Select
    Selection.Font.Bold = False
    With Selection.Font
        .Name = "宋体"
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
    '设置作业登记区域的数据校验
    Range("D6:T" & (SumCount + 5)).Select
    Selection.FormulaHidden = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$AV$6:$AV$22"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "作业情况和点名登记错误"
        .InputMessage = ""
        .ErrorMessage = "请按教学过程登记表右上方成绩等级符号填写，点名到和迟到符号不能与作业等级符号相同！"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Worksheets("0-教学过程登记表（填写+打印)").Visible = temp
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub 修订平时成绩表()
    Worksheets("平时成绩表").Visible = True
    Worksheets("平时成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC25+RC29=0,0,ROUND((RC32)/(R5C25+R5C29),0))"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("平时成绩表").Visible = False
End Sub
Sub 修订专业矩阵状态()
     '修订专业矩阵状态工作表
    Dim MyShapes As Shapes
    Dim Shp As Shape
    Set MyShapes = Worksheets("专业矩阵状态").Shapes
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "版本号"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "修订日期"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "主版本号"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "副版本号"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "修复版本号"
    Range("G1:G5").Select
    Selection.Font.Bold = True
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Set NewShp = ActiveSheet.Buttons.Add(695, 2.25, 102, 27) '（位置高度，位置宽度，按钮高度，按钮宽度）
    NewShp.Characters.Text = "发布版本"
    NewShp.OnAction = "生成版本号"
    NewShp.Font.Name = "微软雅黑"
    NewShp.Font.Size = 14
    NewShp.Font.ColorIndex = 3
    
    Range("K6").Select
    AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
    If (AllowEditCount <> 0) Then
        For i = 1 To AllowEditCount
            Sheets("专业矩阵状态").Protection.AllowEditRanges(1).Delete
        Next i
    End If
    ActiveSheet.Protection.AllowEditRanges.Add Title:="专业", Range:=Range("B4:C12")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="学院", Range:=Range("B2:D2")
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
    Call 设置表格线("A2", "AF12", 12)
    Range("A1:I12").Select
    With Selection.Validation
        .Delete
    End With
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "代码发布路径"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "代码备份路径"
    Range("G5").Select
    Selection.Copy
    Range("G6:G7").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Columns("H:H").Select
    Selection.ColumnWidth = 20
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("专业矩阵状态").Visible = False
End Sub
Sub 修订毕业要求达成度评价表()
    Dim SchoolName As String
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    SchoolName = Range("B2").Value
    Worksheets("专业矩阵状态").Visible = False
    Worksheets("3-综合分析（打印）").Visible = True
    Worksheets("3-综合分析（打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    If SchoolName = "计算机信息与安全学院" Then
        If Range("A9").Value = "（4）毕业要求达成度评价" Then
            Rows("9:10").Select
            Selection.Delete Shift:=xlUp
            Range("A11").Select
            ActiveCell.FormulaR1C1 = "（4）改进措施"
        End If
    End If
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("3-综合分析（打印）").Visible = False
End Sub
Sub 课程目标允许编辑区域()
    Dim AllowEditCount As Integer
    Dim i As Integer
    Worksheets("2-课程目标和综合分析（填写）").Activate
    AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
    If (AllowEditCount <> 0) Then
        For i = 1 To AllowEditCount
            Sheets("2-课程目标和综合分析（填写）").Protection.AllowEditRanges(1).Delete
        Next i
    End If
    ActiveSheet.Protection.AllowEditRanges.Add Title:="课号", Range:=Range("B3")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="比例", Range:=Range("D3:O3")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="是否认证", Range:=Range("Q3")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="认证专业", Range:=Range("B7")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="填写日期", Range:=Range("B8")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="课程目标", Range:=Range("B11:B20")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="评价环节支撑比例", Range:=Range("D11:Q20")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="考核结果分析", Range:=Range("B22")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="有效措施", Range:=Range("B23")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="课程目标达成度", Range:=Range("B25")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="毕业要求达成度", Range:=Range("B27")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="改进措施", Range:=Range("B28")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="格式调整", Range:=Range("Y1")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="辅助", Range:=Range("AE2:AE7")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="评价环节名称1", Range:=Range("D2:E2")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="评价环节名称2", Range:=Range("L2:M2")
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
Sub 打开文档()
Dim Grade As String
Dim Major As String
Dim i As Integer
Dim CourseCount As Integer
Dim PointCount As Integer
Dim MatrixSheet As String
On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    MsgBox ("重新导入教学任务，学生名单，专业矩阵等信息，请稍等。。。")
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Grade = Range("D9").Value
    Major = Range("B7").Value
    On Error Resume Next
    If Sheets("学生名单") Is Nothing Then
        Worksheets("使用帮助").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "学生名单"
    End If
    Sheets("学生名单").Visible = True
    Worksheets("学生名单").Activate
    If (Application.WorksheetFunction.CountIf(Range("E:E"), Grade) < 1) Or (Application.WorksheetFunction.CountIf(Range("F:F"), Major) < 1) Then
        Call 导入学生名单
    End If

    Sheets("学生名单").Visible = False
    
    If Sheets("专业矩阵状态") Is Nothing Then
        Worksheets("使用帮助").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "专业矩阵状态"
        Call 新建专业矩阵状态工作表
    End If
    Call 设置学院专业
    Call 导入矩阵
    Sheets("专业矩阵状态").Visible = False
    Call 指标点数据表公式
    Call 课程目标和综合分析公式
    Worksheets("4-质量分析报告（填写+打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    ActiveSheet.Shapes.Range(Array("Button 5")).Select
    Selection.OnAction = "打印"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
       
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Application.ScreenUpdating = True
End Sub
Sub 导入矩阵()
Dim Major As String
Dim MatrixSheet As String
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Major = Range("B7").Value
    Worksheets("专业矩阵状态").Activate
    Call 提取毕业要求矩阵信息(Major)
    Sheets("专业矩阵状态").Visible = True
    On Error Resume Next
    Worksheets("专业矩阵状态").Activate
    MatrixSheet = Application.Index(Range("C4:C" & MajorLastRow), Application.Match(Major, Range("B4:B" & MajorLastRow), 0))
    If Sheets(MatrixSheet) Is Nothing Then
        Worksheets("毕业要求-指标点数据表").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = MatrixSheet
    End If
    Sheets(MatrixSheet).Visible = True
    Worksheets(MatrixSheet).Activate
    CourseCount = Application.CountA(Range("D4:D200"))
    PointCount = Application.CountA(Range("E4:AS200"))
    Sheets(MatrixSheet).Visible = False
    Worksheets("专业矩阵状态").Activate
    If (Range("D" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = "存在") Then
        If (Range("E" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value <> CourseCount) Or (Range("F" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value <> PointCount) Then
            Call 导入毕业要求矩阵(Major)
        End If
    End If
    Sheets("专业矩阵状态").Visible = False
End Sub
Sub 允许事件触发()
    Application.EnableEvents = True
    Worksheets("2-课程目标和综合分析（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("B7").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=OFFSET(专业矩阵状态!$B$4,,,COUNTA(专业矩阵状态!$B$4:$B$12))"
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
Sub 打印()
' 设置格式 宏
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
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    SchoolName = Range("B2").Value
    ErrNum = 0
    ErrorMsg = ""
    Worksheets("2-课程目标和综合分析（填写）").Activate
    CourseTargetCount = Application.CountA(Range("B11:B20"))
    If CourseTargetCount = 0 Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表未填写【课程目标】" & vbCr & vbLf
    End If
    '简单文档检查
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
        ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表未填写【填写日期】" & vbCr & vbLf
    End If
    Worksheets("课程目标达成度汇总用数据").Visible = True
    Worksheets("课程目标达成度汇总用数据").Activate
    If Range("D2").Value = 0 Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表未填写【课程序号】" & vbCr & vbLf
    ElseIf (Range("B2").Value = "") Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、课程信息提取不完整，请在数据源-教学任务.xls中补充【学期】" & vbCr & vbLf
    ElseIf Range("C2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、课程信息提取不完整，请在数据源-教学任务.xls中补充【课程名称】" & vbCr & vbLf
    ElseIf Range("E2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、课程信息提取不完整，请在数据源-教学任务.xls中补充【主讲教师】" & vbCr & vbLf
    ElseIf Range("G2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、课程信息提取不完整，请在数据源-教学任务.xls中补充【学分】" & vbCr & vbLf
    End If
    Worksheets("2-课程目标和综合分析（填写）").Activate
    LinkCount = Application.CountA(Range("D7:Q7")) - Application.CountBlank(Range("D7:Q7"))
    For i = 4 To LinkCount + 4
        If Cells(7, i).Value <> "" Then
            If Application.CountBlank(Cells(11, i).Resize(10, 1)) = 10 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、评价环节" & Cells(5, i).Value & "有平均成绩，但该评价环节未支撑课程目标。" & vbCr & vbLf
            End If
        Else
            If Application.CountBlank(Cells(11, i).Resize(10, 1)) <> 10 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、评价环节" & Cells(5, i).Value & "没有平均成绩，但该评价环节支撑了课程目标。" & vbCr & vbLf
            End If
        End If
    Next i
    
    For i = 0 To CourseTargetCount - 1
        Worksheets("2-课程目标和综合分析（填写）").Activate
        If Range("R" & i + 11).Value <> 100 Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "、课程目标" & i + 1 & "支撑比例合计不为100。" & vbCr & vbLf
        End If
        Worksheets("课程目标达成度汇总用数据").Activate
        If (Not IsNumeric(Cells(2, 2 * i + 9).Value) Or Cells(2, 2 * i + 9).Value = "" Or Cells(2, 2 * i + 9).Value = 0) Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "、课程目标" & i + 1 & "达成情况不正确请检查。" & vbCr & vbLf
        End If
    Next i
    Worksheets("课程目标达成度汇总用数据").Visible = False
    Worksheets("3-毕业要求数据表（填写）").Activate
    RequirementCount = Application.CountA(Range("C7:C18")) - Application.CountBlank(Range("C7:C18"))
    RequirementReachCount = Application.Count(Range("D7:D18"))

    For i = 0 To Application.CountA(Range("B7:B18")) - 1
        If Range("O" & i + 11).Value <> "" And Range("O" & i + 11).Value <> 100 Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "、" & Cells(i + 7, 2).Value & "支撑比例合计不为100。" & vbCr & vbLf
        End If
        If (Cells(i + 7, 3).Value = "√") And (Not IsNumeric(Cells(i + 7, 4).Value) Or Cells(i + 7, 4).Value = "" Or Cells(i + 7, 4).Value = 0) Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "、" & Cells(i + 7, 2).Value & "达成情况不正确" & vbCr & vbLf
        End If
    Next i
    If ErrorMsg <> "" Then
        MsgBox (ErrorMsg)
        Call CreateTXTfile(ErrorMsg)
        Worksheets("课程目标达成度汇总用数据").Visible = False
        Worksheets("毕业要求达成度汇总用数据").Visible = False
        Exit Sub
    End If
    Call 调整表格格式
    Worksheets("2-课程目标和综合分析（填写）").Activate
    If Range("$Q$3").Value = "非认证" Then
        Worksheets("0-教学过程登记表（填写+打印)").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Call 取消各行颜色
        
        
        Call Excel2PDF("0-教学过程登记表（填写+打印)", ThisWorkbook.Path, PDFFileName & "--教学过程登记表.pdf")
        Call 设置各行颜色
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        Worksheets("4-质量分析报告（填写+打印）").Activate
        Call 取消质量分析报告填写区域颜色
        
        Call Excel2PDF("4-质量分析报告（填写+打印）", ThisWorkbook.Path, PDFFileName & "--质量分析报告.pdf")
        Call 设置质量分析报告填写区域颜色
    ElseIf Range("$Q$3").Value = "认证未提交成绩" Then
        '打印教学过程登记表
        Worksheets("0-教学过程登记表（填写+打印)").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Call 取消各行颜色
        
        
        Call Excel2PDF("0-教学过程登记表（填写+打印)", ThisWorkbook.Path, PDFFileName & "--教学过程登记表.pdf")
        Call 设置各行颜色
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        Sheets("2-毕业要求达成度评价（打印）").Visible = True
        Worksheets("2-毕业要求达成度评价（打印）").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        
        Sheets("1-课程目标达成度评价（打印）").Visible = True
        Worksheets("1-课程目标达成度评价（打印）").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Rows("11:20").Select
        Selection.EntireRow.Hidden = False
        SumRow = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
        Rows(SumRow + 11 & ":20").Select
        Selection.EntireRow.Hidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        Call Excel2PDF("1-课程目标达成度评价（打印）", ThisWorkbook.Path, PDFFileName & "--课程目标达成情况评价.pdf")
        Sheets("1-课程目标达成度评价（打印）").Visible = False

        Select Case SchoolName
            Case "计算机信息与安全学院"
                Sheets("2-毕业要求达成度评价（打印）").Visible = False
                Sheets("3-综合分析（打印）").Visible = True
                Worksheets("3-综合分析（打印）").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                
                
                Call Excel2PDF("3-综合分析（打印）", ThisWorkbook.Path, PDFFileName & "--课程综合分析.pdf")
                Sheets("3-综合分析（打印）").Visible = False
                
                Worksheets("4-质量分析报告（填写+打印）").Activate
                Call 取消质量分析报告填写区域颜色
                
                Call Excel2PDF("4-质量分析报告（填写+打印）", ThisWorkbook.Path, PDFFileName & "--质量分析报告.pdf")
                
                Call 设置质量分析报告填写区域颜色
            Case "电子工程与自动化学院"
                Sheets("2-毕业要求达成度评价（打印）").Visible = True
                Worksheets("2-毕业要求达成度评价（打印）").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                
                Call Excel2PDF("2-毕业要求达成度评价（打印）", ThisWorkbook.Path, PDFFileName & "--毕业要求达成情况评价.pdf")
                Sheets("2-毕业要求达成度评价（打印）").Visible = False
                Sheets("3-综合分析（打印）").Visible = True
                Worksheets("3-综合分析（打印）").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                
                
                Call Excel2PDF("3-综合分析（打印）", ThisWorkbook.Path, PDFFileName & "--课程综合分析.pdf")
                Sheets("3-综合分析（打印）").Visible = False
                
                Worksheets("4-质量分析报告（填写+打印）").Activate
                Call 取消质量分析报告填写区域颜色
                
                Call Excel2PDF("4-质量分析报告（填写+打印）", ThisWorkbook.Path, PDFFileName & "--质量分析报告.pdf")
                
                Call 设置质量分析报告填写区域颜色
            End Select
    ElseIf Range("$Q$3").Value = "认证已提交成绩" Then
        Sheets("2-毕业要求达成度评价（打印）").Visible = True
        Worksheets("2-毕业要求达成度评价（打印）").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        
        Sheets("1-课程目标达成度评价（打印）").Visible = True
        Worksheets("1-课程目标达成度评价（打印）").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Rows("11:20").Select
        Selection.EntireRow.Hidden = False
        SumRow = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
        Rows(SumRow + 11 & ":20").Select
        Selection.EntireRow.Hidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        Call Excel2PDF("1-课程目标达成度评价（打印）", ThisWorkbook.Path, PDFFileName & "--课程目标达成情况评价.pdf")
        Sheets("1-课程目标达成度评价（打印）").Visible = False
        Select Case SchoolName
            Case "计算机信息与安全学院"
                Sheets("2-毕业要求达成度评价（打印）").Visible = False
                Worksheets("3-综合分析（打印）").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                Call Excel2PDF("3-综合分析（打印）", ThisWorkbook.Path, PDFFileName & "--课程综合分析.pdf")
                Sheets("3-综合分析（打印）").Visible = False
            Case "电子工程与自动化学院"
                Sheets("2-毕业要求达成度评价（打印）").Visible = True
                Worksheets("2-毕业要求达成度评价（打印）").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                Call Excel2PDF("2-毕业要求达成度评价（打印）", ThisWorkbook.Path, PDFFileName & "--毕业要求达成情况评价.pdf")
                Sheets("2-毕业要求达成度评价（打印）").Visible = False
                Sheets("3-综合分析（打印）").Visible = True
                Worksheets("3-综合分析（打印）").Activate
                ActiveSheet.PageSetup.CenterFooter = ""
                Call Excel2PDF("3-综合分析（打印）", ThisWorkbook.Path, PDFFileName & "--课程综合分析.pdf")
                Sheets("3-综合分析（打印）").Visible = False
            End Select
    Else
        MsgBox ("2-课程目标和综合分析（填写）工作表中的“是否认证”未选择！")
    End If
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    Worksheets(CurrentWorksheet).Activate
End Sub
Sub 毕业要求数据表公式()
    Dim SchoolName As String
    Dim AllowEditCount As Integer
    Dim i As Integer
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    SchoolName = Range("B2").Value
    Worksheets("3-毕业要求数据表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Select Case SchoolName
        Case "计算机信息与安全学院"
            '三院专用
            Range("A7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(ROW(RC1)-ROW(R7C1)+1=1,1,IF(ROW(RC1)-ROW(R7C1)+1<=MAX('毕业要求-指标点数据表'!C1),R[-1]C1+1,""""))"
            Range("B7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(ISNA(MATCH('3-毕业要求数据表（填写）'!RC[-1],'毕业要求-指标点数据表'!C[-1],0)),"""",""毕业要求""&INDEX('毕业要求-指标点数据表'!C[1],MATCH('3-毕业要求数据表（填写）'!RC[-1],'毕业要求-指标点数据表'!C[-1],0)))"
            Range("C7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(OR(R3C2="""",ISERROR(VLOOKUP(MID(RC[-1],5,LEN(RC[-1])-4),'毕业要求-指标点数据表'!R6C3:R46C6,4,0))),"""",IF(VLOOKUP(MID(RC[-1],5,LEN(RC[-1])-4),'毕业要求-指标点数据表'!R6C3:R46C6,4,0)>0,""√"",""""))"
            Range("E7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(RC2="""","""",IF(ROW(RC[-3])-ROW(R7C[-3])=COLUMN(R[-3]C)-COLUMN(R[-3]C5),100,""""))"
            Range("E7").Select
            Selection.AutoFill Destination:=Range("E7:N7"), Type:=xlFillDefault
            Range("E7:N7").Select
            Range("E7:N7").Select
            Selection.AutoFill Destination:=Range("E7:N16"), Type:=xlFillDefault
            Range("E7:N16").Select
            '统计允许编辑区域个数，并全部删除
            AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
            If (AllowEditCount <> 0) Then
                F
