Attribute VB_Name = "模块1"
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
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H10").Value = "弹出消息框"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("专业矩阵状态").Visible = False
    Worksheets(ThisSheet).Activate
    Call 允许事件触发
End Sub
    
''专业及修订
Sub 修订公式()
Dim TempWorkSheetVisible As Boolean
Dim SumRow As Integer
    Worksheets("1-试卷成绩登记表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("AC2:AI403").Select
    Selection.NumberFormatLocal = "G/通用格式"
    Range("T4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",RC[15],IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0)))"
    Selection.AutoFill Destination:=Range("T4:T403"), Type:=xlFillDefault
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "=RC[-15]"
    Range("AC3").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX('2-课程目标和综合分析（填写）'!R3,MATCH('1-试卷成绩登记表（填写）'!R2C,'2-课程目标和综合分析（填写）'!R2,0))"
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
    TempWorkSheetVisible = Worksheets("4-质量分析报告（填写+打印）").Visible
    Worksheets("4-质量分析报告（填写+打印）").Visible = True
    Worksheets("4-质量分析报告（填写+打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    
    Range("F6:G6").Select
    ActiveCell.FormulaR1C1 = _
        "=RC[-4]-COUNTIF('1-试卷成绩登记表（填写）'!C[14],""旷考"")-COUNTIF('1-试卷成绩登记表（填写）'!C[14],""取消"")-COUNTIF('1-试卷成绩登记表（填写）'!C[14],""缓考"")"
    Selection.NumberFormatLocal = "G/通用格式"
    Range("P12").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF('1-试卷成绩登记表（填写）'!C[4],""取消"")=R6C2,COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,""<60"")-COUNTIF('1-试卷成绩登记表（填写）'!C[4],""旷考"")-COUNTIF('1-试卷成绩登记表（填写）'!C[4],""缓考""),COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,""<60""))"
    Selection.NumberFormatLocal = "G/通用格式"
    Range("F16:H16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("I16:J16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("K16:L16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("M16:O16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("P15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""<60"")+COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""旷考"")+COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""取消"")"

    Range("P16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("H4:K5").Select
    Selection.NumberFormatLocal = "yyyy""年""m""月""d""日"";@"
    Range("A6:P16").Select
    Selection.NumberFormatLocal = "G/通用格式"
    Range("P10,F13:P13,F16:P16").Select
    Selection.NumberFormatLocal = "0.00_ "
    Range("D8:E9,J8:K9,P8:P10").Select
    Range("P8").Activate
    Selection.NumberFormatLocal = "0%"
    Range("P10").Select
    Selection.NumberFormatLocal = "0.00%"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("4-质量分析报告（填写+打印）").Visible = TempWorkSheetVisible
    Call 修订专业矩阵状态
    Worksheets("1-课程目标达成度评价（打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Rows("11:21").Select
    Selection.EntireRow.Hidden = False
    SumRow = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
    Rows(SumRow + 11 & ":20").Select
    Selection.EntireRow.Hidden = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub 其他操作()
    '删除专业下拉多余按钮
    Worksheets("2-课程目标和综合分析（填写）").Activate
    'ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    'Dim sh As Shape
    'For Each sh In ActiveSheet.Shapes
    '    If sh.Name = "Drop Down 5606" Then
    '        sh.Delete
    '    End If
    'Next
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Call 允许事件触发
    Call 修订公式
    If Update = vbYes Then
        Worksheets("专业矩阵状态").Visible = True
        Worksheets("专业矩阵状态").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    
        If Range("H8").Value = "更新公式" Then
            Call 重新设置公式按钮
        End If
        Worksheets("专业矩阵状态").Activate
        Range("H8").Value = Range("H1").Value
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("专业矩阵状态").Visible = False
    End If
    Worksheets("2-课程目标和综合分析（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    ActiveWorkbook.BreakLink Name:="E:\01-学期\质量分析报告模版光信息演示\质量分析报告模版V5.xls", Type _
        :=xlExcelLinks
    ActiveSheet.Shapes.Range(Array("Button 4209")).Select
    Selection.OnAction = "保存文件"
    ActiveSheet.Shapes.Range(Array("Button 5")).Select
    Selection.OnAction = "打印"
    ActiveSheet.Shapes.Range(Array("Button 4210")).Select
    Selection.OnAction = "CreateRecordWorkBook"
    ActiveSheet.Shapes.Range(Array("Button 4211")).Select
    Selection.OnAction = "重新设置公式按钮"
    ActiveSheet.Shapes.Range(Array("Button 4212")).Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Selection.OnAction = "允许事件触发"
End Sub
Sub 工作表加密()
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
Sub 设置备份文件信息()
Dim BackupFilePath As String
Dim ReleaseFile As String
Dim RiviseDate As String
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    BackupFilePath = Range("H7").Value
    ReleaseFile = "模块1源代码.bas"
    RiviseDate = Format(Now, "yyyy-mm-dd")
    CodeFileName(0, CStatus) = "更新"
    CodeFileName(0, CMName) = "模块1"
    CodeFileName(0, CBackup) = BackupFilePath & "\模块1源代码-" & Range("H1").Value & "-" & Format(RiviseDate, "YYYYMMDD") & ".bas"
    CodeFileName(0, CRelease) = ReleaseFile
    
    CodeFileName(1, CStatus) = "更新"
    CodeFileName(1, CMName) = "Sheet13"
    CodeFileName(1, CBackup) = BackupFilePath & "\Sheet13-" & Range("H1").Value & "-" & Format(RiviseDate, "YYYYMMDD") & ".cls"
    CodeFileName(1, CRelease) = "Sheet13.cls"
    
    CodeFileName(2, CStatus) = "清除"
    CodeFileName(2, CCMName) = "Sheet20"
    CodeFileName(2, CBackup) = BackupFilePath & "\Sheet20-" & Range("H1").Value & "-" & Format(RiviseDate, "YYYYMMDD") & ".cls"
    CodeFileName(2, CRelease) = "Sheet20.cls"
    
    CodeFileName(3, CStatus) = "清除"
    CodeFileName(3, CMName) = "Sheet3"
    CodeFileName(3, CBackup) = BackupFilePath & "\Sheet3-" & Range("H1").Value & "-" & Format(RiviseDate, "YYYYMMDD") & ".cls"
    CodeFileName(3, CRelease) = "Sheet3.cls"
    
End Sub
Sub 生成版本号()
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
    Call 修订专业矩阵状态
    BatFile = "发布代码.bat"
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
    TestCode = Range("H9").Value
    If (TestCode = "发布版本") Then
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
        If Vbc.Type = 1 And Mid(Vbc.Name, 1, 2) = "模块" Then
            ModuleCount = ModuleCount + 1
            If (Vbc.Name <> "模块1") Then
                Vbc.Name = "模块1"
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
        Call 设置备份文件信息
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
            If (CodeFileName(i, CStatus) = "更新") Then
                Application.VBE.ActiveVBProject.VBComponents(CodeFileName(i, CMName)).Export (CodeFileName(i, CBackup))
                If TestCode = "发布版本" Then
                    Application.VBE.ActiveVBProject.VBComponents(CodeFileName(i, CMName)).Export (ReleaseFilePath & "\" & CodeFileName(i, CRelease))
                End If
            End If
        Next i
        Application.VBE.ActiveVBProject.VBComponents(CodeFileName(i, CMName)).Export (ReleaseFilePath & "\" & CodeFileName(i, CRelease))
        If TestCode = "发布版本" Then
            Call WriteLastLine(CodeFileName(0, CBackup), "'[版本号]" & Range("H1").Value)
            Call WriteLastLine(ReleaseFilePath & "\" & CodeFileName(0, CRelease), "'[版本号]" & Range("H1").Value)
            
            ReadMeisEmpty = 生成Readme(ReleaseFilePath, "Readme.txt", BackupFilePath, "Readme.txt", "模块1", ReleaseFile, Version, RiviseDate)
            
            '生成Git发布批处理文件
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
                Call MsgInfo(NoMsgBox, "ReadMe文件为空，版本发布失败！")
            End If
        End If
    End If
    Set MyTxtObj = Nothing
    Set fso = Nothing
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("专业矩阵状态").Visible = False
End Sub
Function 生成Readme(ReleaseFilePath As String, ReleaseReadmeFile As String, BackupFilePath As String, BackupReadmeFile As String, ModuleName As String, ReleaseFile As String, Version As String, RiviseDate As String)
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
        '删除临时Readme.txt
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
    End If
    生成Readme = isEmpty
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
    Dim xlsApp As New Excel.Application '需要在工程里引用EXCEL对象哦
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
            If (Mid(.VBComponents(i).Name, 1, 2) = "模块") Then
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
Sub 更新工作表代码(FilePath As String)
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
                If (CodeFileName(j, CStatus) = "更新") Then
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
Sub 远程更新代码()
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
    Call 修订专业矩阵状态
    Call 设置备份文件信息
    Application.ScreenUpdating = False
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    AutoUpdate = Range("H12").Value
    If (Range("H10").Value = "弹出消息框") Then
        NoMsgBox = False
    ElseIf (Range("H10").Select = "不弹出消息框") Then
        NoMsgBox = True
    End If
    LastFilePath = ThisWorkbook.Path
    LastReadme = "Readme.txt"
    If Dir((LastFilePath & "\" & LastReadme)) <> "" Then
        Open LastFilePath & "\" & LastReadme For Input As #1
        Close #1
        Kill LastFilePath & "\" & LastReadme
    End If
    If AutoUpdate = "自动更新" Then
        Update = vbYes
    Else
        Update = MsgBox("正在连接远程服务器，检查代码最新版本！" & vbCrLf & "开始更新代码吗？", vbYesNo, "远程自动更新代码")
    End If
    If Update = vbYes Then
        'Call MsgInfo(NoMsgBox, "正在连接远程服务器，检查代码最新版本！")
        Status = DownFile(LastFilePath, LastReadme, True)
        If Status = False Or Dir(LastFilePath & "\" & LastReadme) = "" Or GetLastLine(LastFilePath & "\" & LastReadme) = "文件为空" Then
            GoTo ErrorSub
        End If
        Call GetVersionFromFile(LastFilePath & "\" & LastReadme)
        CurrentVersion = Range("H1").Value
        CurrentRiviseDate = Range("H2").Value
        CtrResult = StrComp(CurrentVersion, ModuleLastRivise(1, CVersion), vbTextCompare)
         '远程代码版本号比当前代码版本号新
        If CtrResult = "-1" Then
            ModuleFile = ModuleLastRivise(1, CFileName)
            Status = DownFile(ThisWorkbook.Path, ModuleFile, False)
            If GetLastLine(ThisWorkbook.Path & "\" & ModuleFile) = "文件为空" Then
                DownComplete = "1"
            Else
                RemoteVersion = Replace(GetLastLine(ThisWorkbook.Path & "\" & ModuleFile), "'[版本号]", "")
                If Status = False Or Dir(ThisWorkbook.Path & "\" & ModuleFile) = "" Then
                    GoTo ErrorSub
                End If
                DownComplete = StrComp(ModuleLastRivise(1, CVersion), RemoteVersion, vbTextCompare)
            End If
            If DownComplete <> "0" Then
                Call MsgInfo(NoMsgBox, "版本为" & LastVersion & "的代码未下载成功，请重新打开文件自动下载最新代码!")
            Else
                Call 更新工作表代码(ThisWorkbook.Path)
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
                Call MsgInfo(NoMsgBox, "已更新代码版本为：" & LastVersion & "修订日期：" & LastRiviseDate)
                Worksheets("专业矩阵状态").Activate
                Range("H8").Value = "更新公式"
            End If
        Else
            Call MsgInfo(NoMsgBox, "该模版代码版本已经为最新版本!")
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
    Worksheets("专业矩阵状态").Visible = False
    Worksheets("2-课程目标和综合分析（填写）").Activate
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
            ShellAndWait = "下载失败"
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
        TempBuf = "文件为空"
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
    While Not (InStr(1, temp(i), "版本号") > 0 Or InStr(1, temp(i), "End Sub") > 0)
        i = i - 1
    Wend
    If InStr(1, temp(i), "版本号") > 0 Then
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
        If InStr(1, StrTxt(i), "[修订版本]") > 0 Then
            LastVersion = Replace(StrTxt(i), "[修订版本]", "")
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
        Call MsgInfo(NoMsgBox, "请检查" & ThisWorkbook.Path & "\wget.exe 文件是否存在！")
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
    '文本文件每行以换行符结束
    If x <> 0 Then
        StrTxt = Split(StrTemp, vbLf)
        n = UBound(StrTxt) - LBound(StrTxt)
    '文本文件每行已回车换行结束
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
Error: GetVersionFromFile = isEmpty
End Function
Sub MsgInfo(NoMsg As Boolean, Msg As String)
    If Not NoMsg Then
        MsgBox (Msg)
    End If
End Sub
Sub 修订课程目标和综合分析公式()
    On Error Resume Next
    Dim MyShapes As Shapes
    Dim Shp As Shape
    Dim SchoolName As String
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    SchoolName = Range("B2").Value
    Worksheets("专业矩阵状态").Visible = False
    '修订"2-课程目标和综合分析（填写）"工作表评价环节课程报告和作业成绩公式
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Application.EnableEvents = False
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    '2019.5.3修订，解决实验课程实验1，实验2，实验3等满分为100分，合计考核分超过100分的情况
    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C="""","""",COUNTIF('1-试卷成绩登记表（填写）'!C24,R9C4&""-""&R7C2&""-""&""认证""))"
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
    Application.ScreenUpdating = True
End Sub
Sub 修订教学过程登记表公式()
    On Error Resume Next
    Dim temp As Boolean
    '"0-教学过程登记表（填写+打印)"工作表修订标题，学号，姓名关键词，设置字体，设置允许编辑区域
    Application.ScreenUpdating = False
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
    Application.ScreenUpdating = True
End Sub
Sub 修订平时成绩表()
    On Error Resume Next
    Application.ScreenUpdating = False
    Worksheets("平时成绩表").Visible = True
    Worksheets("平时成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC25+RC29=0,0,ROUND((RC32)/(R5C25+R5C29),0))"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("平时成绩表").Visible = False
    Application.ScreenUpdating = True
End Sub
Sub 修订专业矩阵状态()
    On Error Resume Next
     '修订专业矩阵状态工作表
    Dim MyShapes As Shapes
    Dim Shp As Shape
    Dim LastVersion As String
    Application.ScreenUpdating = False
    Set MyShapes = Worksheets("专业矩阵状态").Shapes
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:I12").Select
    With Selection.Validation
        .Delete
    End With
    Range("H1").Select
    ActiveCell.FormulaR1C1 = _
        "=""V""&R[2]C&"".""&TEXT(R[3]C,""00"")&"".""&TEXT(R[4]C,""00"")"
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
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "代码发布路径"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "代码备份路径"
    
    Range("G1:G12").Select
    Selection.Font.Bold = True
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Set NewShp = ActiveSheet.Buttons.Add(745, 2, 86, 25) '（位置高度，位置宽度，按钮高度，按钮宽度）
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
    ActiveSheet.Protection.AllowEditRanges.Add Title:="修复版本号", Range:=Range("H5")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="代码发布版本", Range:=Range("H9")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="消息框状态", Range:=Range("H10")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="生成PDF后打开文档", Range:=Range("H11")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="自动更新代码", Range:=Range("H12")
    Range("G9").Select
    ActiveCell.FormulaR1C1 = "代码发布版本"
    Range("H9").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="测试版本,发布版本"
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
    ActiveCell.FormulaR1C1 = "消息框状态"
    Range("H10").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="弹出消息框,不弹出消息框"
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
    ActiveCell.FormulaR1C1 = "生成PDF后打开文档"
    Range("H11").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="打开PDF,不打开PDF"
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
    ActiveCell.FormulaR1C1 = "自动更新代码"
    Range("H12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="自动更新,手动更新"
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
    Call 设置表格线("A2", "H12", 12)
    Columns("H:H").Select
    Selection.ColumnWidth = 20
    Range("G1:H12").Select
    With Selection.Font
        .Name = "宋体"
        .Size = 12
    End With
    Selection.Font.Bold = True
    Range("A2:F12").Select
    With Selection.Font
        .Name = "宋体"
        .Size = 12
    End With
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("专业矩阵状态").Visible = False
    Application.ScreenUpdating = True
End Sub
Sub 修订毕业要求达成度评价表()
    On Error Resume Next
    Dim SchoolName As String
    Application.ScreenUpdating = False
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
    Application.ScreenUpdating = True
End Sub
Sub 课程目标允许编辑区域()
    On Error Resume Next
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
On Error Resume Next
Dim Grade As String
Dim Major As String
Dim i As Integer
Dim CourseCount As Integer
Dim PointCount As Integer
Dim MatrixSheet As String
On Error Resume Next
    Application.EnableEvents = True
    Call MsgInfo(NoMsgBox, "重新导入教学任务，学生名单，专业矩阵等信息，请稍等。。。")
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
End Sub
Sub 导入矩阵()
On Error Resume Next
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
        'If (Range("E" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value <> CourseCount) Or (Range("F" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value <> PointCount) Then
            Call 导入毕业要求矩阵(Major)
        'End If
    End If
    Sheets("专业矩阵状态").Visible = False
End Sub
Sub 允许事件触发()
    On Error Resume Next
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
    Call 课程目标允许编辑区域
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H10").Value = "弹出消息框"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-课程目标和综合分析（填写）").Activate
End Sub
Sub 提交前检查()
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
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ShoolName = Range("B2").Value
    ErrNum = 0
    ErrorMsg = ""
    ThisFileName = Mid(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1)
    '检查2-课程目标和综合分析（填写） 工作表
    Worksheets("2-课程目标和综合分析（填写）").Activate
    CourseTargetCount = Application.CountA(Range("B11:B20"))
    IdentifyStatus = Range("$Q$3").Value
    If Range("$B$3").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表未填写【课程序号】" & vbCrLf
    End If
    If IdentifyStatus = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表未选择【认证状态】" & vbCrLf
    End If
    If Range("$B$7").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表未选择【认证专业】" & vbCrLf
    End If
    If CourseTargetCount = 0 Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表未填写【课程目标】" & vbCrLf
    End If
    If Range("$B$8").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表未填写【填写日期】" & vbCrLf
    End If

    For i = 1 To 5
        Mark = Application.Index(Range("M7:Q7"), Application.Match(Cells(2, 2 * i + 2).Value, Range("M5:Q5"), 0))
        If Cells(3, 2 * i + 2).Value <> "" And Cells(3, 2 * i + 2).Value <> "0" Then
            If Mark = "" Or Mark = 0 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表评价环节" & Cells(2, 2 * i + 2).Value & "平均得分为空,请检查：（1）如果没有该评价环节，删除比例；（2）试卷成绩登记表中缺少该项成绩" & vbCrLf
            End If
        ElseIf Cells(3, 2 * i + 2).Value = "" Then
            If Mark <> "" Or Mark = 0 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、2-课程目标和综合分析（填写）工作表评价环节" & Cells(2, 2 * i + 2).Value & "比例为空，但平均得分不为空,请检查：（1）是否缺少该评价环节；（2）试卷成绩登记表导入成绩时是否多导入了该项成绩" & vbCrLf
            End If
        End If
    Next i
    LinkCount = Application.CountA(Range("D5:Q5"))
    For i = 4 To LinkCount + 4
        If Cells(7, i).Value <> "" Then
            If Application.CountBlank(Cells(11, i).Resize(10, 1)) = 10 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、评价环节" & Cells(5, i).Value & "有平均成绩，但该评价环节未支撑课程目标。" & vbCrLf
            End If
        Else
            If Application.CountBlank(Cells(11, i).Resize(10, 1)) <> 10 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、评价环节" & Cells(5, i).Value & "没有平均成绩，但该评价环节支撑了课程目标。" & vbCrLf
            End If
        End If
    Next i
    
    For i = 11 To 20
        If Range("B" & i).Value <> "" Then
            If (Range("R" & i).Value = "") Or (Range("R" & i).Value <> 100) Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、课程目标" & i - 10 & "支撑比例合计不是100%！" & vbCrLf
            End If
        ElseIf (Range("R" & i).Value <> "") Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "、没有课程目标" & i - 10 & "，请删除对应的支撑比例" & vbCrLf
        End If
    Next i
    If Range("$B$22").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、缺少（1）考核结果分析" & vbCrLf
    End If
    If Range("$B$23").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、缺少（2）有效的教学方法和措施" & vbCrLf
    End If
    If Range("$B$25").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、缺少（3）课程目标达成度评价" & vbCrLf
    End If
    If ShoolName = "电子工程与自动化学院" Then
        If Range("$B$27").Value = "" Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "、缺少（4）毕业要求达成度评价" & vbCrLf
        End If
    End If
    If Range("$B$28").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、缺少（5）改进措施" & vbCrLf
    End If

    For i = 0 To CourseTargetCount - 1
        Worksheets("课程目标达成度汇总用数据").Activate
        If (Not IsNumeric(Cells(2, 2 * i + 9).Value) Or Cells(2, 2 * i + 9).Value = "" Or Cells(2, 2 * i + 9).Value = 0) Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "、课程目标" & i + 1 & "达成情况不正确请检查。" & vbCrLf
        End If
    Next i
    
    If (ShoolName = "电子工程与自动化学院") Then
        Worksheets("3-毕业要求数据表（填写）").Visible = True
        Worksheets("3-毕业要求数据表（填写）").Activate
        RequirementCount = Application.CountA(Range("C7:C18")) - Application.CountBlank(Range("C7:C18"))
        RequirementReachCount = Application.Count(Range("D7:D18"))
        For i = 0 To Application.CountA(Range("B7:B18")) - 1
            If Range("O" & i + 11).Value <> "" And Range("O" & i + 11).Value <> 100 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、" & Cells(i + 7, 2).Value & "支撑比例合计不为100。" & vbCrLf
            End If
            If (Cells(i + 7, 3).Value = "√") And (Not IsNumeric(Cells(i + 7, 4).Value) Or Cells(i + 7, 4).Value = "" Or Cells(i + 7, 4).Value = 0) Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、" & Cells(i + 7, 2).Value & "达成情况不正确" & vbCrLf
            End If
        Next i
    End If

    Worksheets("课程目标达成度汇总用数据").Visible = True
    Worksheets("课程目标达成度汇总用数据").Activate
    If (Range("B2").Value = "") Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、课程信息提取不完整，请在数据源-教学任务.xls中补充【学期】" & vbCrLf
    ElseIf Range("C2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、课程信息提取不完整，请在数据源-教学任务.xls中补充【课程名称】" & vbCrLf
    ElseIf Range("E2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、课程信息提取不完整，请在数据源-教学任务.xls中补充【主讲教师】" & vbCrLf
    ElseIf Range("G2").Value = "" Then
        ErrNum = ErrNum + 1
        ErrorMsg = ErrorMsg & ErrNum & "、课程信息提取不完整，请在数据源-教学任务.xls中补充【学分】" & vbCrLf
    End If
    Worksheets("课程目标达成度汇总用数据").Visible = False
    Worksheets("毕业要求达成度汇总用数据").Visible = False
    

    If IdentifyStatus = "非认证" Or IdentifyStatus = "认证未提交成绩" Then
        Worksheets("0-教学过程登记表（填写+打印)").Activate
        If (Application.WorksheetFunction.CountIf(Range("Z6:Z185"), "取消") <> Application.WorksheetFunction.CountIf(Range("AD6:AD185"), "取消")) Then
            ErrNum = ErrNum + 1
            ErrorMsg = ErrorMsg & ErrNum & "、以下同学平时成绩为0，请核实是否取消考试资格，若取消，请在成绩类别列选择“取消”" & vbCrLf
            For i = 6 To 185
                If (Range("B" & i).Value <> "") And (Range("Z" & i).Value = "取消") Then
                    If (Range("AD" & i).Value = "") Then
                        ErrorMsg = ErrorMsg & Range("B" & i).Value & "  "
                    End If
                End If
            Next i
            ErrorMsg = ErrorMsg & vbCrLf
        End If
    End If
    '检查“1-试卷成绩登记表（填写）”工作表
    Worksheets("1-试卷成绩登记表（填写）").Activate
    For i = 1 To 9
        If Cells(3, i + 4) = "" Then
            If Application.WorksheetFunction.Count(Cells(4, i + 4).Resize(403, 1)) <> 0 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、试卷成绩登记表中" & Cells(2, i + 4) & "满分为空，但该列有成绩，请填写满分。" & vbCrLf
            End If
        Else
            If Application.WorksheetFunction.Count(Cells(4, i + 4).Resize(403, 1)) = 0 Then
                ErrNum = ErrNum + 1
                ErrorMsg = ErrorMsg & ErrNum & "、试卷成绩登记表中" & Cells(2, i + 4) & "满分不为空，但该列没有成绩。" & vbCrLf
            End If
        End If
    Next i
 
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Term = Range("$B$2").Value
    Term = Mid(Term, 3, 2) & "-" & Mid(Term, 8, 2) & "-" & Mid(Term, 14, 1)
    CourseNum = Range("$B$3").Value
    CourseName = Range("$B$4").Value
    Teacher = Range("$B$5").Value
    Major = Range("$B$7").Value
    If Dir(ThisWorkbook.Path & "\错误报告\" & ThisFileName & "-错误检查报告.txt") <> "" Then
        Open ThisWorkbook.Path & "\错误报告\" & ThisFileName & "-错误检查报告.txt" For Input As #1
        Close #1
        Kill ThisWorkbook.Path & "\错误报告\" & ThisFileName & "-错误检查报告.txt"
    End If
    If ErrorMsg <> "" Then
        If Dir(ThisWorkbook.Path & "\错误报告\") = "" Then
            MkDir ThisWorkbook.Path & "\错误报告\"
        End If
        Call CreateTXTfile(ThisWorkbook.Path & "\错误报告\" & ThisFileName & "-错误检查报告.txt", ErrorMsg, True)
        NoError = False
    End If
End Sub
Sub 打印()
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H11").Value = "打开PDF"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("专业矩阵状态").Visible = False
    Call 生成PDF
End Sub
Sub 生成PDF()
    On Error Resume Next
    '设置格式 宏
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
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    If (Range("H11").Value = "打开PDF") Then
        isOpenAfterPublish = True
    ElseIf (Range("H11").Value = "不打开PDF") Then
        isOpenAfterPublish = False
    End If
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("专业矩阵状态").Visible = False
    
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Term = Range("$B$2").Value
    Term = Mid(Term, 3, 2) & "-" & Mid(Term, 8, 2) & "-" & Mid(Term, 14, 1)
    CourseNum = Range("$B$3").Value
    CourseName = Range("$B$4").Value
    Teacher = Range("$B$5").Value
    Major = Range("$B$7").Value
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    SchoolName = Range("B2").Value
    Worksheets("专业矩阵状态").Visible = False
    PDFFilePath = ThisWorkbook.Path & "\PDF版"
    If Dir(PDFFilePath) = "" Then
        MkDir PDFFilePath
    End If
    PDFFileName = Term & "-" & CourseNum & "-" & Major & "-" & Teacher & "-" & CourseName
    '错误检查报告不存在，没有错误
    If (Dir(ThisWorkbook.Path & "\错误报告\" & PDFFileName & "-错误检查报告.txt") = "") Then
        Call 调整表格格式
        Worksheets("2-课程目标和综合分析（填写）").Activate
        If Range("$Q$3").Value = "非认证" Then
            Worksheets("0-教学过程登记表（填写+打印)").Activate
            ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
            Call Excel2PDF("0-教学过程登记表（填写+打印)", PDFFilePath, PDFFileName & "--教学过程登记表.pdf", isOpenAfterPublish)
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
            
            Worksheets("4-质量分析报告（填写+打印）").Activate
            Call 取消质量分析报告填写区域颜色
            
            Call Excel2PDF("4-质量分析报告（填写+打印）", PDFFilePath, PDFFileName & "--质量分析报告.pdf", isOpenAfterPublish)
            Call 设置质量分析报告填写区域颜色
        ElseIf Range("$Q$3").Value = "认证未提交成绩" Then
            '打印教学过程登记表
            Worksheets("0-教学过程登记表（填写+打印)").Activate
            ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
            Call Excel2PDF("0-教学过程登记表（填写+打印)", PDFFilePath, PDFFileName & "--教学过程登记表.pdf", isOpenAfterPublish)
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
            
            Sheets("2-毕业要求达成度评价（打印）").Visible = True
            Worksheets("2-毕业要求达成度评价（打印）").Activate
            ActiveSheet.PageSetup.CenterFooter = ""
            
            Sheets("1-课程目标达成度评价（打印）").Visible = True
            Worksheets("1-课程目标达成度评价（打印）").Activate
            ActiveSheet.PageSetup.CenterFooter = ""
            ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
            Rows("11:21").Select
            Selection.EntireRow.Hidden = False
            SumRow = Application.WorksheetFunction.CountA(Range("B11:B20")) - Application.WorksheetFunction.CountBlank(Range("B11:B20"))
            Rows(SumRow + 11 & ":20").Select
            Selection.EntireRow.Hidden = True
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
            
            Call Excel2PDF("1-课程目标达成度评价（打印）", PDFFilePath, PDFFileName & "--课程目标达成情况评价.pdf", isOpenAfterPublish)
            Sheets("1-课程目标达成度评价（打印）").Visible = False
    
            Select Case SchoolName
                Case "计算机信息与安全学院"
                    Sheets("2-毕业要求达成度评价（打印）").Visible = False
                    Sheets("3-综合分析（打印）").Visible = True
                    Worksheets("3-综合分析（打印）").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    
                    
                    Call Excel2PDF("3-综合分析（打印）", PDFFilePath, PDFFileName & "--课程综合分析.pdf", isOpenAfterPublish)
                    Sheets("3-综合分析（打印）").Visible = False
                    
                    Worksheets("4-质量分析报告（填写+打印）").Activate
                    Call 取消质量分析报告填写区域颜色
                    
                    Call Excel2PDF("4-质量分析报告（填写+打印）", PDFFilePath, PDFFileName & "--质量分析报告.pdf", isOpenAfterPublish)
                    
                    Call 设置质量分析报告填写区域颜色
                Case "电子工程与自动化学院"
                    Sheets("2-毕业要求达成度评价（打印）").Visible = True
                    Worksheets("2-毕业要求达成度评价（打印）").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    
                    Call Excel2PDF("2-毕业要求达成度评价（打印）", PDFFilePath, PDFFileName & "--毕业要求达成情况评价.pdf", isOpenAfterPublish)
                    Sheets("2-毕业要求达成度评价（打印）").Visible = False
                    Sheets("3-综合分析（打印）").Visible = True
                    Worksheets("3-综合分析（打印）").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    
                    
                    Call Excel2PDF("3-综合分析（打印）", PDFFilePath, PDFFileName & "--课程综合分析.pdf", isOpenAfterPublish)
                    Sheets("3-综合分析（打印）").Visible = False
                    
                    Worksheets("4-质量分析报告（填写+打印）").Activate
                    Call 取消质量分析报告填写区域颜色
                    
                    Call Excel2PDF("4-质量分析报告（填写+打印）", PDFFilePath, PDFFileName & "--质量分析报告.pdf", isOpenAfterPublish)
                    
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
            
            Call Excel2PDF("1-课程目标达成度评价（打印）", PDFFilePath, PDFFileName & "--课程目标达成情况评价.pdf", isOpenAfterPublish)
            Sheets("1-课程目标达成度评价（打印）").Visible = False
            Select Case SchoolName
                Case "计算机信息与安全学院"
                    Sheets("2-毕业要求达成度评价（打印）").Visible = False
                    Worksheets("3-综合分析（打印）").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    Call Excel2PDF("3-综合分析（打印）", PDFFilePath, PDFFileName & "--课程综合分析.pdf", isOpenAfterPublish)
                    Sheets("3-综合分析（打印）").Visible = False
                Case "电子工程与自动化学院"
                    Sheets("2-毕业要求达成度评价（打印）").Visible = True
                    Worksheets("2-毕业要求达成度评价（打印）").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    Call Excel2PDF("2-毕业要求达成度评价（打印）", PDFFilePath, PDFFileName & "--毕业要求达成情况评价.pdf", isOpenAfterPublish)
                    Sheets("2-毕业要求达成度评价（打印）").Visible = False
                    Sheets("3-综合分析（打印）").Visible = True
                    Worksheets("3-综合分析（打印）").Activate
                    ActiveSheet.PageSetup.CenterFooter = ""
                    Call Excel2PDF("3-综合分析（打印）", PDFFilePath, PDFFileName & "--课程综合分析.pdf", isOpenAfterPublish)
                    Sheets("3-综合分析（打印）").Visible = False
            End Select
        Else
            Call MsgInfo(NoMsgBox, "2-课程目标和综合分析（填写）工作表中的“是否认证”未选择！")
        End If
    End If
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    Worksheets(CurrentWorksheet).Activate
End Sub
Sub 毕业要求数据表公式()
    On Error Resume Next
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
                For i = 1 To AllowEditCount
                    Sheets("3-毕业要求数据表（填写）").Protection.AllowEditRanges(1).Delete
                Next i
            End If
            Range("E7:N18").Select
            Range("N18").Activate
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Case "电子工程与自动化学院"
            '八院专用
            Range("A7").Select
            ActiveCell.FormulaR1C1 = "=ROW(RC1)-ROW(R7C1)+1"
            Range("B7").Select
            ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",""毕业要求""&RC1)"
            Range("C7").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(OR(R3C2="""",ISERROR(VLOOKUP(RC[-1],'毕业要求-指标点数据表'!R6C2:R46C6,5,0))),IF(RC[1]<>"""",""√"",""""),IF(VLOOKUP(RC[-1],'毕业要求-指标点数据表'!R6C2:R46C6,5,0)>0,""√"",IF(RC[1]<>"""",""√"","""")))"
            '统计允许编辑区域个数，并全部删除
            AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
            If (AllowEditCount <> 0) Then
                For i = 1 To AllowEditCount
                    Sheets("3-毕业要求数据表（填写）").Protection.AllowEditRanges(1).Delete
                Next i
            End If
            Range("E7:N18").Select
            Range("N18").Activate
            ActiveSheet.Protection.AllowEditRanges.Add Title:="区域1", Range:=Range( _
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
        "=IF(SUM(RC[-10]:RC[-1])=0,"""",IF(SUM(RC[-10]:RC[-1])<>100,""比例合计不为100"",SUM(RC[-10]:RC[-1])))"

    Range("Q7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]="""","""",IF(RC[-2]<>100,1,IF(RC[-14]<>""√"",1,"""")))"
    Range("R7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]="""","""",IF(RC[-3]<>100,1,IF(RC[-15]<>""√"",1,"""")))"
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
    Worksheets("专业矩阵状态").Visible = False
End Sub
Sub 设置表格主题()
    On Error Resume Next
    Application.ScreenUpdating = False
    Worksheets("3-毕业要求数据表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("E1:K1").Select
    ActiveCell.FormulaR1C1 = _
        "='2-课程目标和综合分析（填写）'!R[6]C[-3]&""专业-毕业要求达成情况评价数据表"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("1-课程目标达成度评价（打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:R1").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[1]C[1]&""  课程目标达成情况评价表"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-毕业要求达成度评价（打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:N1").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[1]C[1]&""  毕业要求达成情况评价表"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.ScreenUpdating = True
End Sub
Sub 毕业要求达成度评价公式()
    On Error Resume Next
    Worksheets("2-毕业要求达成度评价（打印）").Visible = True
    Worksheets("2-毕业要求达成度评价（打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "='3-毕业要求数据表（填写）'!R[1]C"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "='3-毕业要求数据表（填写）'!R[1]C"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]="""","""",VLOOKUP(RC[-1],'3-毕业要求数据表（填写）'!R7C[-1]:R18C[1],3,0))"
    Range("D4:M4").Select
    Selection.FormulaArray = "=TRANSPOSE('2-课程目标和综合分析（填写）'!R[7]C[-1]:R[16]C[-1])"
    Range("B6:C6").Select
    Selection.AutoFill Destination:=Range("B6:C17"), Type:=xlFillDefault
    Range("B6:C17").Select
    Range("D6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('3-毕业要求数据表（填写）'!R[1]C[1]="""","""",'3-毕业要求数据表（填写）'!R[1]C[1])"
    Range("D6").Select
    Selection.AutoFill Destination:=Range("D6:M6"), Type:=xlFillDefault
    Range("D6:M6").Select
    Selection.AutoFill Destination:=Range("D6:M17"), Type:=xlFillDefault
    Range("D6:M17").Select
    Range("N6").Select
    ActiveCell.FormulaR1C1 = "='3-毕业要求数据表（填写）'!R[1]C[1]"
    Range("N6").Select
    Selection.AutoFill Destination:=Range("N6:N17"), Type:=xlFillDefault
    Range("A1:N1").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[1]C[1]&""  毕业要求达成情况评价表"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-毕业要求达成度评价（打印）").Visible = False
End Sub


Sub 指标点数据表公式()
    On Error Resume Next
    Dim SchoolName As String
    Dim Major As String
    Dim i As Integer
    Dim MajorCount As Integer
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Major = Range("B7").Value
    Worksheets("专业矩阵状态").Visible = True
    Worksheets("专业矩阵状态").Activate
    SchoolName = Range("B2").Value
    MajorCount = Application.WorksheetFunction.CountA(Range("B4:B" & MajorLastRow))
    Sheets("毕业要求-指标点数据表").Visible = True
    Worksheets("毕业要求-指标点数据表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("B2:C2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[2]C"
    Range("B3:C3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R2C2="""","""",IF(OR(MID(R[-1]C,LEN(R[-1]C),1)=""A"",MID(R[-1]C,LEN(R[-1]C),1)=""B"",MID(R[-1]C,LEN(R[-1]C),1)=""C""),MID(R[-1]C,1,LEN(R[-1]C)-1),IF(ISNUMBER(FIND(""（"",R[-1]C,1)),MID(R[-1]C,1,FIND(""（"",R[-1]C,1)-1),IF(ISNUMBER(FIND(""("",R[-1]C,1)),MID(R[-1]C,1,FIND(""("",R[-1]C,1)-1),R[-1]C))))"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC5=1,IF(R[-1]C=""序号"",1,R[-1]C+1),R[-1]C)"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R4C3="""",ISNA(MATCH('2-课程目标和综合分析（填写）'!R7C2,专业矩阵状态!C2,0))),"""",IF(VLOOKUP(R4C3,INDIRECT(""'""&INDEX(专业矩阵状态!C3,MATCH('2-课程目标和综合分析（填写）'!R7C2,专业矩阵状态!C2,0))&""'!$B$3:$AS$3""),ROW(RC3)-2,0)="""","""",VLOOKUP(R4C3,INDIRECT(""'""&INDEX(专业矩阵状态!C3,MATCH('2-课程目标和综合分析（填写）'!R7C2,专业矩阵状态!C2,0))&""'!$B$3:$AS$3""),ROW(RC3)-2,0)))"
    Range("E6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R2C2="""",RC[-1]="""",ISNA(MATCH('2-课程目标和综合分析（填写）'!R7C2,专业矩阵状态!C2,0))),"""",IF(ISNA(VLOOKUP(R4C[-2]&""-""&R3C2&""*"",INDIRECT(""'""&INDEX(专业矩阵状态!C3,MATCH('2-课程目标和综合分析（填写）'!R7C2,专业矩阵状态!C2,0))&""'!$A:$AS""),ROW(RC1)-1,0)),0,VLOOKUP(R4C[-2]&""-""&R3C2&""*"",INDIRECT(""'""&INDEX(专业矩阵状态!C3,MATCH('2-课程目标和综合分析（填写）'!R7C2,专业矩阵状态!C2,0))&""'!$A:$AS""),ROW(RC1)-1,0)))"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C2,RC2,C[-1])"
    Range("C4:G4").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R7C2"
    Range("A6").Select
    Selection.AutoFill Destination:=Range("A6:A46"), Type:=xlFillDefault
    Range("D6:G6").Select
    Range("D6").Activate
    Selection.AutoFill Destination:=Range("D6:G46"), Type:=xlFillDefault
    Range("D6:G46").Select
    Columns("H:S").Select
    Selection.Delete Shift:=xlToLeft
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets("毕业要求-指标点数据表").Visible = False
    Worksheets("专业矩阵状态").Visible = False
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
Sub 保存文件()
' 保存文件 宏
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
    Call 设置区域颜色("1-试卷成绩登记表（填写）", "E2:M" & MaxLineCout, xlThemeColorDark1)
    Call 设置区域颜色("3-毕业要求数据表（填写）", "E7:N18", xlThemeColorDark1)

    Worksheets("2-课程目标和综合分析（填写）").Activate
    CourseName = Range("$B$4").Value
    
    Worksheets("课内实验成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("成绩核对").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password

    Worksheets("成绩录入").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("0-教学过程登记表（填写+打印)").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("1-试卷成绩登记表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("2-课程目标和综合分析（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("3-毕业要求数据表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("4-质量分析报告（填写+打印）").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("1-课程目标达成度评价（打印）").Activate
    ActiveSheet.PageSetup.CenterFooter = "1-" & CourseName & "课程目标达成度评价表"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("2-毕业要求达成度评价（打印）").Activate
    ActiveSheet.PageSetup.CenterFooter = "2-" & CourseName & "毕业要求达成度评价表"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("3-综合分析（打印）").Activate
    ActiveSheet.PageSetup.CenterFooter = "3 -" & CourseName & "课程综合分析表"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
   
    Application.ScreenUpdating = False
    Call MsgInfo(NoMsgBox, "正在进行工作表格式调整，并保存文件，请耐心等待！")
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Call 设置教学过程登记表
    Call 设置质量分析报告格式
    'Application.Calculation = xlManual
    Call 调整表格格式
    'Call 文档填写检查
    FileName = ThisWorkbook.Name

    Worksheets("2-课程目标和综合分析（填写）").Activate
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
Sub 调整表格格式()
    On Error Resume Next
    '"1-课程目标达成度评价（打印）" 按2页自动设置课程目标行高
    Worksheets("1-课程目标达成度评价（打印）").Visible = True
    Worksheets("1-课程目标达成度评价（打印）").Activate
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
    Worksheets("1-课程目标达成度评价（打印）").Activate
    Worksheets("1-课程目标达成度评价（打印）").Visible = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Worksheets("2-毕业要求达成度评价（打印）").Visible = True
    Worksheets("2-毕业要求达成度评价（打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Rows("3:17").Select
    Selection.RowHeight = 25
    Worksheets("2-毕业要求达成度评价（打印）").Visible = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password

    '"3-综合分析（打印）" 按1页自动设行高
    Worksheets("3-综合分析（打印）").Visible = True
    Worksheets("3-综合分析（打印）").Activate
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
    Worksheets("3-综合分析（打印）").Activate
    Worksheets("3-综合分析（打印）").Visible = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub 重新设置公式按钮()
  On Error Resume Next
  Dim SumCount As Integer
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Worksheets("0-教学过程登记表（填写+打印)").Activate
  SumCount = Application.WorksheetFunction.Count(Range("A6:A185"))
  ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
  Call 设置各行颜色
  ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
  Call 课程目标和综合分析公式
  Call 重新设置公式(SumCount)
  Call 试卷成绩登记表公式
  Call 成绩录入公式
  Call 实验成绩表公式
  Call 毕业要求数据表公式
  Call 指标点数据表公式
  Call 平时成绩表公式
  Call 成绩核对表公式
  Call 成绩表公式
  Call 毕业要求达成度评价公式
  Call 评价环节比例设置公式
  Call 质量分析报告公式
  Worksheets("0-教学过程登记表（填写+打印)").Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

Sub 设置教学过程登记表()
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
    Worksheets("0-教学过程登记表（填写+打印)").Activate
    CourseNumber = Range("AN1").Value
    SumCount = Range("AM1").Value
    If (SumCount > 0) Then
        Call 重新设置公式(SumCount)
    ElseIf (SumCount = 0) Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    If (CourseNumber <> "") Then
    '获取人数
        'Count = Application.WorksheetFunction.Count(Range("B6:B185"))
        Call 设置表格线("A6", "AG" & (SumCount + 5), 9)
        Call 删除表格线(SumCount + 6, MaxRecord)
        Call 最后一行表格线(SumCount + 5)
        HPageBreaksCount = 0
        ActiveSheet.ResetAllPageBreaks
        '每30行记录后插入分页符
        Pages = Int(SumCount / 30)
        If (SumCount Mod 30 = 0) Then
            Pages = Pages - 1
        End If
        For i = 1 To Pages
          Rows(30 * i + 6 & ":" & 30 * i + 6).Select
          ActiveWindow.SelectedSheets.HPageBreaks.Add before:=ActiveCell
          HPageBreaksCount = HPageBreaksCount + 1
        Next i
        '设置实验成绩和考核成绩为可以带小数格式
        Range("AC6:AD" & (SumCount + 6)).Select
        Selection.NumberFormatLocal = "G/通用格式"
        Range("A" & (SumCount + 6) & ":AF" & (MaxRecord + 3)).Select
        Selection.ClearContents
        '取消所有单元格数据校验
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

        '设置成绩类别列的数据校验
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
        '设置考试类别列的数据校验
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
        '设置未批改作业成绩评价方式
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
        '设置最后一行信息格式
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
    
        '写入最后一行成绩构成比例，填表人，填写日期，教研室主任等信息
        Range("A" & SumCount + 7).Select
        ActiveCell.FormulaR1C1 = "=R2C34"
        Range("AJ1").Value = Count + 7
        '插入分页符
        Rows(SumCount + 8 & ":" & SumCount + 8).Select
        Range("A" & SumCount + 8).Activate
        ActiveWindow.SelectedSheets.HPageBreaks.Add before:=ActiveCell
    
        Range("A1:AG" & SumCount + 7).Select
        ActiveSheet.PageSetup.PrintArea = "$A$1:$AG$" & SumCount + 7
    Else
        Call 设置表格线("A6", "AG185", 9)
    End If
        '设置教学过程登记表表头格式
    Call 设置表格线("A4", "AG5", 9)
    Range("A4:AG5").Select
    With Selection.Font
        .Name = "宋体"
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
        .FontStyle = "常规"
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
    Call 设置各行颜色
    Columns("AH:AR").Select
    Range("AH2").Activate
    Selection.EntireColumn.Hidden = True
    
    Columns("AY:AZ").Select
    Selection.EntireColumn.Hidden = True
    Columns("BE:BE").Select
    Selection.EntireColumn.Hidden = True
    
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
    '保护工作表
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub 成绩表公式()
    On Error Resume Next
    Worksheets("成绩表").Activate
    Sheets("成绩表").Visible = 1
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "总评成绩"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "序号"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "学号"
    Range("R1").Select
    Selection.FormulaR1C1 = "文本学号"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "姓名"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "总评成绩"
    Range("U1").Select
    Selection.FormulaR1C1 = "平时成绩"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "成绩类别"
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "考试类别"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""序号"",1,R[-1]C+1)"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[-1],OFFSET(R1C1,R3C33-1,R3C34-1,601,1),0)),INDEX(OFFSET(R1C1,R4C33-1,R4C34-1,601,1),MATCH(RC[-1],OFFSET(R1C1,R3C33-1,R3C34-1,601,1),0)),IF(ISNUMBER(MATCH(RC[-1],OFFSET(R1C1,R3C33-1,R3C37-1,601,1),0)),INDEX(OFFSET(R1C1,R4C33-1,R4C37-1,601,1),MATCH(RC[-1],OFFSET(R1C1,R3C33-1,R3C37-1,601,1),0)),""""))"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]=0,"""",IF(ISNUMBER(RC[-1]),TEXT(RC[-1],""0000000000""),RC[-1]))"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),""""))"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-2]="""",INDEX(C33,MATCH(R1C,C30,0))=0),"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),""""))"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),""""))"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),""""))"
    Range("W2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0))),""""))"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0))),""""))"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0))),""""))"
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0))),""""))"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),VALUE(INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0))),""""))"
    Range("AB2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),""""))"
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(INDEX(C33,MATCH(R1C,C30,0))=0,"""",IF(ISNUMBER(MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),INDEX(OFFSET(R1C1,INDEX(C33,MATCH(R1C,C30,0))-1,INDEX(C34,MATCH(R1C,C30,0))-1,601,1),MATCH(RC17,OFFSET(R1C1,INDEX(C33,MATCH(""学号"",C30,0))-1,INDEX(C34,MATCH(""学号"",C30,0))-1,601,1),0)),""""))"
    Range("P2:AC2").Select
    Selection.AutoFill Destination:=Range("P2:AC601"), Type:=xlFillDefault
    Range("P2:AC601").Select
    
    Range("AD6:AD15").Select
    Selection.FormulaArray = "=TRANSPOSE(R[-5]C[-10]:R[-5]C[-1])"
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "偏移"
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "栏数"
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "关键词所在行"
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "第1栏关键词所在列"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[1]=2,2,1)"
    Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=COUNT(R[2]C[2]:R[2]C[5])"
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(OFFSET(R[-1]C1,0,R4C34-1,200))"
    Range("AD3").Select
    ActiveCell.FormulaR1C1 = "序号"
    Range("AD4").Select
    ActiveCell.FormulaR1C1 = "学号"
    Range("AD5").Select
    ActiveCell.FormulaR1C1 = "姓名"
    Range("AE3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[3]<>"""",OFFSET(R1C1,RC33-1,RC34-1,1,1),"""")"
    Range("AE3").Select
    Selection.AutoFill Destination:=Range("AE3:AE15"), Type:=xlFillDefault
    Range("AE3:AE15").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets("成绩表").Visible = 0
End Sub
Sub 课程目标和综合分析公式()
    Dim Evaluation1 As String
    Dim Evaluation2 As String
    On Error Resume Next
    Call 导入教学任务
    Application.ScreenUpdating = False
    Worksheets("2-课程目标和综合分析（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Application.EnableEvents = False
    Evaluation1 = Range("D2").Value
    Range("D2:E2").Select
    ActiveCell.FormulaR1C1 = Evaluation1
    Range("F2:G2").Select
    ActiveSheet.Unprotect
    ActiveCell.FormulaR1C1 = "作业成绩"
    Range("H2:I2").Select
    ActiveCell.FormulaR1C1 = "实验成绩"
    Range("J2:K2").Select
    ActiveCell.FormulaR1C1 = "课堂测验"
    Evaluation2 = Range("L2").Value
    Range("L2:M2").Select
    ActiveCell.FormulaR1C1 = Evaluation2
    Range("N2:O2").Select
    ActiveCell.FormulaR1C1 = "考核成绩"
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
        .Name = "宋体"
        .FontStyle = "加粗"
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
    Range("D4:L4").Select
    ActiveCell.FormulaR1C1 = "=IF(R5C4=""一"",""考试"",""各实验成绩"")"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-试卷成绩登记表（填写）'!R[-3]C[2]="""","""",'1-试卷成绩登记表（填写）'!R[-3]C[2])"
    Range("N5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-试卷成绩登记表（填写）'!R[-3]C[2]="""","""",'1-试卷成绩登记表（填写）'!R[-3]C[2])"
    Range("O5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-试卷成绩登记表（填写）'!R[-3]C[2]="""","""",'1-试卷成绩登记表（填写）'!R[-3]C[2])"
    Range("P5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-试卷成绩登记表（填写）'!R[-3]C[2]="""","""",'1-试卷成绩登记表（填写）'!R[-3]C[2])"
    Range("Q5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('1-试卷成绩登记表（填写）'!R[-3]C[2]="""","""",'1-试卷成绩登记表（填写）'!R[-3]C[2])"

    Range("V1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[2]C[-20]="""",0,IF(ISNUMBER(MATCH(TEXT(R[2]C[-20],""0000000""),教学过程登记表!C[-21],0)),COUNTIF(教学过程登记表!C[-20],R[2]C[-20]&""*""),IF(ISNUMBER(MATCH(R[2]C[-20],教学过程登记表!C[-21],0)),COUNTIF(教学过程登记表!C[-20],R[2]C[-20]&""*""),0)))"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R3C2="""","""",IF(R4C2=""请在数据源-教学任务中添加该课号的课程信息"","""",MID(VLOOKUP(R3C2,教学任务!C1:C13,2,0),1,9)&""学年第 ""&MID(VLOOKUP(R3C2,教学任务!C1:C13,2,0),11,1)&"" 学期""))"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SUM(RC[-12],RC[-10],RC[-8],RC[-6],RC[-4],RC[-2])=0,"""",SUM(RC[-12],RC[-10],RC[-8],RC[-6],RC[-4],RC[-2]))"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R3C2="""","""",IF(ISERROR(VLOOKUP(R3C2,教学任务!R1C1:R65536C13,5,0)),""请在数据源-教学任务中添加该课号的课程信息"",VLOOKUP(R3C2,教学任务!R1C1:R65536C13,5,0)))"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R3C2="""","""",IF(R4C2=""请在数据源-教学任务中添加该课号的课程信息"","""",VLOOKUP(R3C2,教学任务!R1C1:R65536C13,6,0)))"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C="""","""",COUNTIF('1-试卷成绩登记表（填写）'!C24,R9C4&""-""&R7C2&""-""&""认证""))"
    Range("B9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-6]C="""","""",IF(R4C2=""请在数据源-教学任务中添加该课号的课程信息"","""",VLOOKUP(R3C2,教学任务!R1C1:R65536C13,4,0)))"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[1]="""","""",""课程目标""&(ROW(RC)-10))"
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
        "=IF('1-试卷成绩登记表（填写）'!R[-3]C[1]="""","""",'1-试卷成绩登记表（填写）'!R[-3]C[1])"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2="""",R[-1]C="""",COUNTIF('1-试卷成绩登记表（填写）'!C24,R9C4&""-""&R7C2&""-""&""认证"")=0,COUNT(OFFSET('1-试卷成绩登记表（填写）'!R4C1,,MATCH(R5C,'1-试卷成绩登记表（填写）'!R2,0)-1,205))=0),"""",SUMIF('1-试卷成绩登记表（填写）'!C24,R9C4&""-""&R7C2&""-""&""认证"",OFFSET('1-试卷成绩登记表（填写）'!R1C1,,MATCH(R5C,'1-试卷成绩登记表（填写）'!R2,0)-1,183))/R6C2)"
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
        "=IF(OR(R6C2=0,R6C2="""",R[-1]C="""",COUNTIF('1-试卷成绩登记表（填写）'!C24,R9C4&""-""&R7C2&""-""&""认证"")=0,COUNT(OFFSET('1-试卷成绩登记表（填写）'!R4C1,,MATCH(R5C,'1-试卷成绩登记表（填写）'!R2,0)-1,205))=0),"""",SUMIF('1-试卷成绩登记表（填写）'!C24,R9C4&""-""&R7C2&""-""&""认证"",OFFSET('1-试卷成绩登记表（填写）'!R1C1,,MATCH(R5C,'1-试卷成绩登记表（填写）'!R2,0)-1,183))/R6C2)"
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C="""","""",ROUND(R[-1]C*100/R[-2]C,1))"
    Range("M6:M8").Select
    Selection.AutoFill Destination:=Range("M6:R8"), Type:=xlFillDefault
    Range("M6:R8").Select
    Range("R7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(COUNTBLANK(RC[-14]:RC[-6])=14,SUM(R6C4:R6C12)=0),"""",SUM('2-课程目标和综合分析（填写）'!R7C4:R7C12)*100/SUM(R6C4:R6C12))"
    Range("R8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C18="""",R7C18="""",R7C18=0),"""",ROUND(R7C18*100/R6C18,1))"
    Range("D9:E9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-6]C[-2]="""","""",IF(R4C2=""请在数据源-教学任务中添加该课号的课程信息"","""",VLOOKUP(R3C2,教学任务!R1C1:R65536C13,8,0)))"
    Range("H9:I9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-6]C[-6]="""","""",IF(R4C2=""请在数据源-教学任务中添加该课号的课程信息"","""",VLOOKUP(R3C2,教学任务!R1C1:R65536C13,3,0)))"
    Range("N9:R9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SUM(R[2]C[-11]:R[11]C[-11])=0,"""",ROUND(AVERAGE(R[2]C[-11]:R[11]C[-11]),0))"
    Range("B21:R21").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2=""""),"""",IF(CONCATENATE(R[-14]C[17],R[-14]C[18],R[-14]C[19],R[-14]C[20],R[-14]C[21],R[-14]C[22],R[-14]C[23],R[-14]C[24],R[-14]C[25])="""","""",""试卷各题平均得分比为：""&MID(CONCATENATE(R[-14]C[17],R[-14]C[18],R[-14]C[19],R[-14]C[20],R[-14]C[21],R[-14]C[22],R[-14]C[23],R[-14]C[24],R[-14]C[25]),1,LEN(CONCATENATE(R[-14]C[17],R[-14]C[18],R[-14]C[19],R[-14]C[2" & _
        "0],R[-14]C[21],R[-14]C[22],R[-14]C[23],R[-14]C[24],R[-14]C[25]))-1)))" & _
        ""
    Range("B22:R22").Select
    Range("B24:R24").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2=""""),"""",IF(CONCATENATE(R[-13]C[31],R[-12]C[31],R[-11]C[31],R[-10]C[31],R[-9]C[31],R[-8]C[31],R[-7]C[31],R[-6]C[31],R[-5]C[31],R[-4]C[31])="""","""",""课程目标达成度分别为：""&MID(CONCATENATE(R[-13]C[31],R[-12]C[31],R[-11]C[31],R[-10]C[31],R[-9]C[31],R[-8]C[31],R[-7]C[31],R[-6]C[31],R[-5]C[31],R[-4]C[31]),1,LEN(CONCATENATE(R[-13]C[31],R[-12]C[31],R[-11]C[3" & _
        "1],R[-10]C[31],R[-9]C[31],R[-8]C[31],R[-7]C[31],R[-6]C[31],R[-5]C[31],R[-4]C[31]))-1)))" & _
        ""
    Range("B26:R26").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R6C2=0,R6C2=""""),"""",IF(CONCATENATE(RC[18],RC[19],RC[20],RC[21],RC[22],RC[23],RC[24],RC[25],RC[26],RC[27],RC[28],RC[29])="""","""",""毕业要求达成度分别为：""&MID(CONCATENATE(RC[18],RC[19],RC[20],RC[21],RC[22],RC[23],RC[24],RC[25],RC[26],RC[27],RC[28],RC[29]),1,LEN(CONCATENATE(RC[18],RC[19],RC[20],RC[21],RC[22],RC[23],RC[24],RC[25],RC[26],RC[27],RC[28],RC[29]))-1)))"
    Range("B27:R27").Select
    Range("T1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(R[2]C[-18]&""-1"",教学过程登记表!C[-18],0)),""无名单"",IF(MATCH(R[2]C[-18]&""-1"",教学过程登记表!C[-18],0)-6<0,0,MATCH(R[2]C[-18]&""-1"",教学过程登记表!C[-18],0)-6))"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[2]C[-20]="""",0,IF(ISNUMBER(MATCH(TEXT(R[2]C[-20],""0000000""),教学过程登记表!C[-21],0)),COUNTIF(教学过程登记表!C[-20],R[2]C[-20]&""*""),IF(ISNUMBER(MATCH(R[2]C[-20],教学过程登记表!C[-21],0)),COUNTIF(教学过程登记表!C[-20],R[2]C[-20]&""*""),0)))"
    Range("W1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]<>""无名单"",MATCH(R[2]C[-21]&""-1"",教学过程登记表!C[-21],0)+COUNTIF(教学过程登记表!C[-22],R[2]C[-21])-1,"""")"
    Range("AG1").Select
    Selection.FormulaR1C1 = "=IF(R3C2="""","""",VLOOKUP(R3C2,教学任务!C1:C13,2,0))"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "=ISNUMBER(MATCH(R[1]C[-18],教学过程登记表!C[-19],0))"
    Range("U2:W2").Select
    ActiveCell.FormulaR1C1 = "='0-教学过程登记表（填写+打印)'!R[2]C[9]"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "='4-质量分析报告（填写+打印）'!R[4]C[-22]"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "='4-质量分析报告（填写+打印）'!R[3]C[-17]"
    Range("T4").Select
    Selection.FormulaR1C1 = _
        "=VLOOKUP(""比例"",评价环节比例设置!R3C[-19]:R3C[-12],MATCH(""折合平时"",评价环节比例设置!R2C[-19]:R2C[-12],0),0)"
    Range("U4").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(""比例"",评价环节比例设置!R3C[-20]:R3C[-13],MATCH(""实验成绩"",评价环节比例设置!R2C[-20]:R2C[-13],0),0)"
    Range("V4").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(""比例"",评价环节比例设置!R3C[-21]:R3C[-14],MATCH(""考核成绩"",评价环节比例设置!R2C[-21]:R2C[-14],0),0)"
    Range("W4").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(""比例"",评价环节比例设置!R3C[-22]:R3C[-15],MATCH(""期中成绩"",评价环节比例设置!R2C[-22]:R2C[-15],0),0)"
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
        "=IF(R[-1]C="""","""",""第""&R[-2]C[-15]&""题：""&TEXT(R[-1]C,""00.0%"")&""，"")"
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
        "=IF(ISNA(MATCH(R[-3]C&""*"",'3-毕业要求数据表（填写）'!C2,0)),"""",INDEX('3-毕业要求数据表（填写）'!C4,MATCH(R[-3]C&""*"",'3-毕业要求数据表（填写）'!C2,0)))"
    Range("T25").Select
    Selection.AutoFill Destination:=Range("T25:AE25"), Type:=xlFillDefault
    Application.EnableEvents = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub 质量分析报告公式()
    On Error Resume Next
    Worksheets("4-质量分析报告（填写+打印）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("H4:K5").Select
    Selection.NumberFormatLocal = "yyyy""年""m""月""d""日"";@"
    Range("A6:P16").Select
    Selection.NumberFormatLocal = "G/通用格式"
    Range("P10,F13:P13,F16:P16").Select
    Selection.NumberFormatLocal = "0.00_ "
    Range("D8:E9,J8:K9,P8:P10").Select
    Range("P8").Activate
    Selection.NumberFormatLocal = "0%"
    Range("P10").Select
    Selection.NumberFormatLocal = "0.00%"
    Range("A2:P2").Select
    ActiveCell.FormulaR1C1 = _
        "=""填表日期： ""&TEXT('2-课程目标和综合分析（填写）'!R[6]C[1],""YYYY年MM月DD日"")&""    """
    Range("C3:E3").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[2]C[-1]"
    Range("K3:P3").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[1]C[-9]"
    Range("C4:E5").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[-1]C[-1]"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTA('1-试卷成绩登记表（填写）'!R4C2:R183C2)-COUNTBLANK('1-试卷成绩登记表（填写）'!R4C2:R183C2)"
    Range("F6:G6").Select
    ActiveCell.FormulaR1C1 = _
        "=RC[-4]-COUNTIF('1-试卷成绩登记表（填写）'!C[14],""旷考"")-COUNTIF('1-试卷成绩登记表（填写）'!C[14],""取消"")-COUNTIF('1-试卷成绩登记表（填写）'!C[14],""缓考"")"
    Range("H6:P6").Select
    ActiveCell.FormulaR1C1 = _
        "=""人；缓考：""&COUNTIF('1-试卷成绩登记表（填写）'!C[12],""缓考"")&""人;旷考：""&COUNTIF('1-试卷成绩登记表（填写）'!C[13],""旷考"")&""人；取消考试资格：""&COUNTIF('1-试卷成绩登记表（填写）'!C[13],""取消"")&""人"""
    Range("D8:E8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(评价环节比例设置!R3C1,,MATCH(R8C1,评价环节比例设置!R2,0)-1,1,1)=0,"""",OFFSET(评价环节比例设置!R3C1,,MATCH(R8C1,评价环节比例设置!R2,0)-1,1,1))"
    Range("J8:K8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(评价环节比例设置!R3C1,,MATCH(R8C6,评价环节比例设置!R2,0)-1,1,1)=0,"""",OFFSET(评价环节比例设置!R3C1,,MATCH(R8C6,评价环节比例设置!R2,0)-1,1,1))"
    Range("P8").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(评价环节比例设置!R3C1,,MATCH(R8C12,评价环节比例设置!R2,0)-1,1,1)=0,"""",OFFSET(评价环节比例设置!R3C1,,MATCH(R8C12,评价环节比例设置!R2,0)-1,1,1))"
    Range("D9:E9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(评价环节比例设置!R3C1,,MATCH(R9C1,评价环节比例设置!R2,0)-1,1,1)=0,"""",OFFSET(评价环节比例设置!R3C1,,MATCH(R9C1,评价环节比例设置!R2,0)-1,1,1))"
    Range("J9:K9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(评价环节比例设置!R3C1,,MATCH(R9C6,评价环节比例设置!R2,0)-1,1,1)=0,"""",OFFSET(评价环节比例设置!R3C1,,MATCH(R9C6,评价环节比例设置!R2,0)-1,1,1))"
    Range("P9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OFFSET(评价环节比例设置!R3C1,,MATCH(R9C12,评价环节比例设置!R2,0)-1,1,1)=0,"""",OFFSET(评价环节比例设置!R3C1,,MATCH(R9C12,评价环节比例设置!R2,0)-1,1,1))"
    Range("F10:I10").Select
    ActiveCell.FormulaR1C1 = "=""最高""&MAX('1-试卷成绩登记表（填写）'!R4C14:R400C14)&""分"""
    Range("J10:L10").Select
    ActiveCell.FormulaR1C1 = "=""最低""&MIN('1-试卷成绩登记表（填写）'!R4C14:R400C14)&""分"""
    Range("F12:H12").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,"">=90"")"
    Range("I12:J12").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,"">=80"")-COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,"">=90"")"
    Range("K12:L12").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,"">=70"")-COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,"">=80"")"
    Range("M12:O12").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,"">=60"")-COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,"">=70"")"
    Range("P12").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF('1-试卷成绩登记表（填写）'!C[4],""取消"")=R6C2,COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,""<60"")-COUNTIF('1-试卷成绩登记表（填写）'!C[4],""旷考"")-COUNTIF('1-试卷成绩登记表（填写）'!C[4],""缓考""),COUNTIF('1-试卷成绩登记表（填写）'!R4C14:R400C14,""<60""))"
    Range("F15:H15").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,"">=90"")"
    Range("I15:J15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,"">=80"")-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,"">=90"")"
    Range("K15:L15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,"">=70"")-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,"">=80"")"
    Range("M15:O15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,"">=60"")-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,"">=70"")"
    Range("P15").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""<60"")+COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""旷考"")+COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""取消"")"
  
    
    Range("A9:C9").Select
    ActiveCell.FormulaR1C1 = "='0-教学过程登记表（填写+打印)'!R[-5]C[26]"
    Range("L9:O9").Select
    ActiveCell.FormulaR1C1 = "='0-教学过程登记表（填写+打印)'!R[-5]C[13]"
    
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
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("I16:J16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("K16:L16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("M16:O16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("P16").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R6C6=0,"""",R[-1]C*100/(R6C2-COUNTIF('1-试卷成绩登记表（填写）'!R4C20:R400C20,""缓考"")))"
    Range("D17:P17").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[5]C[-2]"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub 成绩核对表公式()
    On Error Resume Next
    Worksheets("成绩核对").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "学号"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "姓名"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "平时成绩"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "考核成绩"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "总评成绩"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "成绩类别"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "考试类别"
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
        "=IF(ISNA(MATCH(RC[-1],'0-教学过程登记表（填写+打印)'!R6C1:R179C1,0)),"""",IF(INDEX('0-教学过程登记表（填写+打印)'!R6C2:R179C2,MATCH(RC[-1],'0-教学过程登记表（填写+打印)'!R6C1:R179C1,0))=0,"""",INDEX('0-教学过程登记表（填写+打印)'!R6C2:R179C2,MATCH(RC[-1],'0-教学过程登记表（填写+打印)'!R6C1:R179C1,0))))"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(RC[-2],'0-教学过程登记表（填写+打印)'!R6C1:R179C1,0)),"""",IF(INDEX('0-教学过程登记表（填写+打印)'!R6C2:R179C2,MATCH(RC[-2],'0-教学过程登记表（填写+打印)'!R6C1:R179C1,0))=0,"""",INDEX('0-教学过程登记表（填写+打印)'!R6C3:R179C3,MATCH(RC[-2],'0-教学过程登记表（填写+打印)'!R6C1:R179C1,0))))"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]=""请导入名单"","""",INDEX(成绩录入!R3C[1]:R183C[1],MATCH(RC2,成绩录入!R3C1:R183C1,0)))"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]=""请导入名单"","""",INDEX(成绩录入!R3C[1]:R183C[1],MATCH(RC2,成绩录入!R3C1:R183C1,0)))"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]=""请导入名单"","""",INDEX(成绩录入!R3C[1]:R183C[1],MATCH(RC2,成绩录入!R3C1:R183C1,0)))"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-5]=""请导入名单"","""",INDEX(成绩录入!R3C[1]:R183C[1],MATCH(RC2,成绩录入!R3C1:R183C1,0)))"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]=""请导入名单"","""",INDEX(成绩录入!R3C[1]:R183C[1],MATCH(RC2,成绩录入!R3C1:R183C1,0)))"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""请导入名单""),"""",IF(COUNTBLANK(R3C16:R182C16)=180,""未导入"",IF(ISNA(MATCH(RC2,C16,0)),""退课"",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))))"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""请导入名单"",RC10=""未导入"",RC10=""退课""),"""",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""请导入名单"",RC10=""未导入"",RC10=""退课""),"""",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""请导入名单"",RC10=""未导入"",RC10=""退课""),"""",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""请导入名单"",RC10=""未导入"",RC10=""退课""),"""",IF(RC[-6]=INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0)),""OK"",INDEX(OFFSET(R1C1,,MATCH(MID(R2C,3,4),R2,0)-1,185),MATCH(RC2,C16,0))))"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(成绩核对!RC2,成绩表!C18,0)),"""",VLOOKUP(RC2,成绩表!C18:C29,MATCH(""文本""&R2C16,成绩表!R1C18:R1C29,0),0))"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(成绩核对!RC2,成绩表!C18,0)),"""",VLOOKUP(RC2,成绩表!C18:C29,MATCH(R2C,成绩表!R1C18:R1C29,0),0))"
    Range("T3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(成绩核对!RC2,成绩表!C18,0)),"""",VLOOKUP(RC2,成绩表!C18:C29,MATCH(R2C,成绩表!R1C18:R1C29,0),0))"
    Range("U3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(成绩核对!RC2,成绩表!C18,0)),"""",VLOOKUP(RC2,成绩表!C18:C29,MATCH(R2C,成绩表!R1C18:R1C29,0),0))"
    Range("V3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(成绩核对!RC2,成绩表!C18,0)),"""",IF(ISNUMBER(VLOOKUP(RC2,成绩表!C18:C29,MATCH(R2C,成绩表!R1C18:R1C29,0),0)),ROUND(VLOOKUP(RC2,成绩表!C18:C29,MATCH(R2C,成绩表!R1C18:R1C29,0),0),0),VLOOKUP(RC2,成绩表!C18:C29,MATCH(R2C,成绩表!R1C18:R1C29,0),0)))"
    Range("W3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(成绩核对!RC2,成绩表!C18,0)),"""",VLOOKUP(RC2,成绩表!C18:C29,MATCH(R2C,成绩表!R1C18:R1C29,0),0))"
    Range("X3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(成绩核对!RC2,成绩表!C18,0)),"""",VLOOKUP(RC2,成绩表!C18:C29,MATCH(R2C,成绩表!R1C18:R1C29,0),0))"
    Range("B3:X3").Select
    Selection.AutoFill Destination:=Range("B3:X182"), Type:=xlFillDefault
    Range("B3").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub 平时成绩表公式()
    On Error Resume Next
    Worksheets("平时成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password

    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(RC1,'0-教学过程登记表（填写+打印)'!C1,0)),"""",INDEX('0-教学过程登记表（填写+打印)'!C,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C1,0)))"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(MATCH(RC1,'0-教学过程登记表（填写+打印)'!C1,0)),"""",INDEX('0-教学过程登记表（填写+打印)'!C,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C1,0)))"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR('0-教学过程登记表（填写+打印)'!RC=INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH(""已交未批改"",'0-教学过程登记表（填写+打印)'!C46,0)),ISNUMBER(MATCH('0-教学过程登记表（填写+打印)'!RC,'0-教学过程登记表（填写+打印)'!R18C48:R23C48,0)),'0-教学过程登记表（填写+打印)'!RC=""X""),'0-教学过程登记表（填写+打印)'!RC,IF(ISNA(MATCH('0-教学过程登记表（填写+打印)'!RC,'0-教学过程登记表（填写+打印)'!R6C48:R23C48,0)),0,INDEX('0-教学过程登记表（填写+打印)'!R6C49:R23C49,MATCH('0-教学过程登记表（填写+打印)'!RC,'0" & _
        "-教学过程登记表（填写+打印)'!R6C48:R23C48,0))))" & _
        ""
    Range("D6").Select
    Selection.AutoFill Destination:=Range("D6:U6"), Type:=xlFillDefault
    Range("D6:U6").Select
    Range("V6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER('0-教学过程登记表（填写+打印)'!RC),'0-教学过程登记表（填写+打印)'!RC,IF(ISNA(MATCH('0-教学过程登记表（填写+打印)'!RC,'0-教学过程登记表（填写+打印)'!R6C48:R23C48,0)),0,INDEX('0-教学过程登记表（填写+打印)'!R6C49:R23C49,MATCH('0-教学过程登记表（填写+打印)'!RC,'0-教学过程登记表（填写+打印)'!R6C48:R23C48,0))))"
    Range("W6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER('0-教学过程登记表（填写+打印)'!RC),'0-教学过程登记表（填写+打印)'!RC,IF(ISNA(MATCH('0-教学过程登记表（填写+打印)'!RC,'0-教学过程登记表（填写+打印)'!R6C48:R23C48,0)),0,INDEX('0-教学过程登记表（填写+打印)'!R6C49:R23C49,MATCH('0-教学过程登记表（填写+打印)'!RC,'0-教学过程登记表（填写+打印)'!R6C48:R23C48,0))))"
    Range("X6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER('0-教学过程登记表（填写+打印)'!RC),'0-教学过程登记表（填写+打印)'!RC,IF(ISNA(MATCH('0-教学过程登记表（填写+打印)'!RC,'0-教学过程登记表（填写+打印)'!R6C48:R23C48,0)),0,INDEX('0-教学过程登记表（填写+打印)'!R6C49:R23C49,MATCH('0-教学过程登记表（填写+打印)'!RC,'0-教学过程登记表（填写+打印)'!R6C48:R23C48,0))))"
    Range("Y6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[5]:RC[6])"
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC25+RC29=0,0,ROUND((RC32)/(R5C25+R5C29),0))"
    Range("AA6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R5C28=0,RC[-25]=0),0,ROUND(SUMIF(R4C[-5]:R4C[-3],""测验"",RC[-5]:RC[-3])/R5C28,0))"
    Range("AB6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(OFFSET(RC1,,R3C[6]-1,,R3C38),"">0"")"
    Columns("AC:AC").Select
    Selection.NumberFormatLocal = "G/通用格式"
    Range("AC6").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF('0-教学过程登记表（填写+打印)'!RC4:RC21,INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH(""点名到"",'0-教学过程登记表（填写+打印)'!C46,0)))+COUNTIF('0-教学过程登记表（填写+打印)'!RC4:RC21,INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH(""点名迟到"",'0-教学过程登记表（填写+打印)'!C46,0)))+COUNTIF('0-教学过程登记表（填写+打印)'!RC4:RC21,INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH(""点名请假"",'0-教学过程登记表（填写+打印)'!C46,0)))"
    Range("AD6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(RC4:RC21,"">0"")"
    Range("AE6").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF(RC4:RC21,INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH(""已交未批改"",'0-教学过程登记表（填写+打印)'!C46,0)))"
    Range("AF6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[3])"
    Range("AG6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC4:RC21)"
    Range("AH6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('0-教学过程登记表（填写+打印)'!R6C50=""默认成绩"",RC31*VLOOKUP(INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH(""已交未批改"",'0-教学过程登记表（填写+打印)'!C46,0)),'0-教学过程登记表（填写+打印)'!C48:C49,2,0),RC31*RC36)"
    Columns("AC:AJ").Select
    Selection.NumberFormatLocal = "G/通用格式"
    Range("AI6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R5C29=0,0,COUNTIF('0-教学过程登记表（填写+打印)'!RC4:RC21,INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH('0-教学过程登记表（填写+打印)'!R18C46,'0-教学过程登记表（填写+打印)'!C46,0)))*VLOOKUP(INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH('0-教学过程登记表（填写+打印)'!R18C46,'0-教学过程登记表（填写+打印)'!C46,0)),'0-教学过程登记表（填写+打印)'!C48:C49,2,0)+COUNTIF('0-教学过程登记表（填写+打印)'!RC4:RC21,INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH('0-教学过程登记表（填写+打印)'!R19C46,'0-教" & _
        "学过程登记表（填写+打印)'!C46,0)))*VLOOKUP(INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH('0-教学过程登记表（填写+打印)'!R19C46,'0-教学过程登记表（填写+打印)'!C46,0)),'0-教学过程登记表（填写+打印)'!C48:C49,2,0)+COUNTIF('0-教学过程登记表（填写+打印)'!RC4:RC21,INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH('0-教学过程登记表（填写+打印)'!R20C46,'0-教学过程登记表（填写+打印)'!C46,0)))*VLOOKUP(INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH('0-教学过程登记表（填写+打印)'!R20C46,'0-教学过程登记表（填写+打印)'!C46" & _
        ",0)),'0-教学过程登记表（填写+打印)'!C48:C49,2,0)+COUNTIF('0-教学过程登记表（填写+打印)'!RC4:RC21,INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH('0-教学过程登记表（填写+打印)'!R21C46,'0-教学过程登记表（填写+打印)'!C46,0)))*VLOOKUP(INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH('0-教学过程登记表（填写+打印)'!R21C46,'0-教学过程登记表（填写+打印)'!C46,0)),'0-教学过程登记表（填写+打印)'!C48:C49,2,0))" & _
        ""
    Range("AJ6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC25=0,0,IF(RC[-6]=0,INDEX('0-教学过程登记表（填写+打印)'!R6C[13]:R19C[13],MATCH(INDEX('0-教学过程登记表（填写+打印)'!C48,MATCH(""已交未批改"",'0-教学过程登记表（填写+打印)'!C46,0)),'0-教学过程登记表（填写+打印)'!R6C[12]:R19C[12],0)),ROUND(RC33/RC30,1)))"
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

Sub 试卷成绩登记表公式()
On Error Resume Next
    Worksheets("1-试卷成绩登记表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("AB3").Select
    ActiveCell.FormulaR1C1 = "剔除不及格学生"
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
    ActiveCell.FormulaR1C1 = "学号"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "姓名"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "='0-教学过程登记表（填写+打印)'!R[2]C[16]"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[-11]"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[-10]"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[-9]"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[-8]"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[-7]"
    If (Range("T2").Value <> "总评成绩") Then
        Columns("T:T").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    If (Range("U2").Value <> "成绩类别") Then
        Columns("U:U").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("T2").Select
        ActiveCell.FormulaR1C1 = "总评成绩"
        Range("U2").Select
        ActiveCell.FormulaR1C1 = "成绩类别"
        Range("V2").Select
        ActiveCell.FormulaR1C1 = "专业"
        Range("W2").Select
        ActiveCell.FormulaR1C1 = "年级"
        Range("X2").Select
        ActiveCell.FormulaR1C1 = "年级-专业-认证状态"
        Range("Y2").Select
        ActiveCell.FormulaR1C1 = "认证状态"
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
    Worksheets("1-试卷成绩登记表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Call 试卷成绩登记表核心公式
    
    If Sheets("学生名单") Is Nothing Then
        Worksheets("使用帮助").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "学生名单"
    End If
    Worksheets("学生名单").Activate
    If (Application.WorksheetFunction.CountIf(Range("E:E"), Grade) < 1) Then
        Call 导入学生名单
    End If

End Sub
Sub 试卷成绩登记表核心公式()
On Error Resume Next
    Worksheets("1-试卷成绩登记表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("B4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C2="""","""",IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(VLOOKUP(RC1,成绩表!C16:C19,MATCH(""文本""&R2C2,成绩表!R1C16:R1C19,0),0)="""",IF(ISNA(VLOOKUP(RC1,'0-教学过程登记表（填写+打印)'!C1:C3,MATCH(R2C2,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0)),"""",IF(VLOOKUP(RC1,'0-教学过程登记表（填写+打印)'!C1:C3,MATCH(R2C2,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0)=""请导入名单"","""",VLOOKUP(RC1,'0-教学过程登记表" & _
        "（填写+打印)'!C1:C3,MATCH(R2C2,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0))),VLOOKUP(RC1,成绩表!C16:C19,MATCH(""文本""&R2C2,成绩表!R1C16:R1C19,0),0)),IF(ISNA(VLOOKUP(RC1,'0-教学过程登记表（填写+打印)'!C1:C3,MATCH(R2C2,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0)),"""",VLOOKUP(RC1,'0-教学过程登记表（填写+打印)'!C1:C3,MATCH(R2C2,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0))))" & _
        ""
    Range("C4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C2="""","""",IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(VLOOKUP(RC1,成绩表!C16:C19,MATCH(""文本""&R2C2,成绩表!R1C16:R1C19,0),0)="""",IF(ISNA(VLOOKUP(RC1,'0-教学过程登记表（填写+打印)'!C1:C3,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0)),"""",VLOOKUP(RC1,'0-教学过程登记表（填写+打印)'!C1:C3,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0)),VLOOKUP(RC2,成绩表!C18:C19,MATCH(R2C3,成绩表!R1" & _
        "C18:R1C19,0),0)),IF(ISNA(VLOOKUP(RC1,'0-教学过程登记表（填写+打印)'!C1:C3,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0)),"""",VLOOKUP(RC1,'0-教学过程登记表（填写+打印)'!C1:C3,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C1:R4C3,0),0))))" & _
        ""
    Range("N4").Select
    ActiveCell.FormulaR1C1 = _
                "=IF(OR(RC2="""",COUNT(RC[-9]:RC[-1])=0),"""",SUM(RC[-9]:RC[-1]))"

    Range("O4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(OR(RC[-13]="""",ISNA(VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0)))"
    Range("P4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(OR(RC[-14]="""",ISNA(VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0)))"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(OR(RC[-15]="""",ISNA(VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0)))"
    Range("R4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(OR(RC[-16]="""",ISNA(VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C&""平均分"",'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0)))"
    Range("S4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(OR(RC[-17]="""",ISNA(VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0))),"""",VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),IF(OR(RC2="""",R3C19=""""),"""",VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C&""得分"",'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0)/INDEX(评价环节比例设置!R3,MATCH(R2C19,评价环节比例设置!R2,0))))"
    Range("T4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(ISNA(VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),"""",VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0)))"
    Range("U4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(ISNA(VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),"""",VLOOKUP(RC2,成绩表!C18:C27,MATCH(R2C,成绩表!R1C18:R1C27,0),0)),IF(OR(RC2=""""),"""",IF(VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0)=0,"""",VLOOKUP(RC2,'0-教学过程登记表（填写+打印)'!C2:C44,MATCH(R2C,'0-教学过程登记表（填写+打印)'!R4C2:R4C44,0),0))))"
    Range("V4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-20]=""""),"""",IF(ISNA(MATCH(RC2&""-""&'2-课程目标和综合分析（填写）'!R7C2,学生名单!C7,0)),IF(ISNUMBER(MATCH(VALUE(RC2),学生名单!C2,0)),VLOOKUP(VALUE(RC2),学生名单!C2:C6,5,0),IF(ISNUMBER(MATCH(RC2,学生名单!C2,0)),VLOOKUP(RC2,学生名单!C2:C6,5,0),"""")),INDEX(学生名单!C6,MATCH(RC2&""-""&'2-课程目标和综合分析（填写）'!R7C2,学生名单!C7,0))))"
    Range("W4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-21]=""""),"""",IF(ISNA(VLOOKUP(VALUE(RC[-21]),学生名单!C2:C6,4,0)),IF(ISNA(VLOOKUP(RC[-21],学生名单!C2:C6,4,0)),"""",VLOOKUP(RC[-21],学生名单!C2:C6,4,0)),VLOOKUP(VALUE(RC[-21]),学生名单!C2:C6,4,0)))"
    Range("X4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-22]="""",RC[1]=""""),"""",RC[-1]&""-""&RC[-2]&""-""&RC[1])"
    Range("Y4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(AND(RC[-3]<>"""",R2C26=""""),AND(RC[-3]<>"""",R2C26=""剔除不及格学生"",RC[-5]>=60)),IF(ISNUMBER(RC[-2]),IF(AND(RC[-2]='2-课程目标和综合分析（填写）'!R9C4,RC[-3]='2-课程目标和综合分析（填写）'!R7C2),""认证"",""""),IF(AND(VALUE(RC[-2])='2-课程目标和综合分析（填写）'!R9C4,RC[-3]='2-课程目标和综合分析（填写）'!R7C2),""认证"","""")),"""")"
    
    Range("A4:C4").Select
    Selection.AutoFill Destination:=Range("A4:C" & MaxLineCout), Type:=xlFillDefault
    Range("A4:C" & MaxLineCout).Select
    Range("N4:Y4").Select
    Range("W4").Activate
    Selection.AutoFill Destination:=Range("N4:Y" & MaxLineCout), Type:=xlFillDefault
    Range("N4:Y" & MaxLineCout).Select
    Worksheets("1-试卷成绩登记表（填写）").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub 成绩录入公式()
On Error Resume Next
    Worksheets("成绩录入").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=MID('2-课程目标和综合分析（填写）'!R[1]C[1],1,9)&""-""&MID('2-课程目标和综合分析（填写）'!R[1]C[1],14,1)&""学期 期末成绩"""
    Range("A2:H2").Select
    ActiveCell.FormulaR1C1 = _
        "='2-课程目标和综合分析（填写）'!R[2]C[1]&""（成绩比例=""&CONCATENATE(评价环节比例设置!R12C2,评价环节比例设置!R12C3,评价环节比例设置!R12C4,评价环节比例设置!R12C5)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=""课号：""&'2-课程目标和综合分析（填写）'!R[1]C[-7]"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('0-教学过程登记表（填写+打印)'!R[2]C[1]<>"""",'0-教学过程登记表（填写+打印)'!R[2]C[1],"""")"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",'0-教学过程登记表（填写+打印)'!R[2]C[1])"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]="""","""",INDEX('0-教学过程登记表（填写+打印)'!C39,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0)))"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-5]="""","""",INDEX('0-教学过程登记表（填写+打印)'!C30,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0)))"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC1="""","""",IF(OR(INDEX('0-教学过程登记表（填写+打印)'!C31,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0))=""取消"",INDEX('0-教学过程登记表（填写+打印)'!C31,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0))=""旷考"",INDEX('0-教学过程登记表（填写+打印)'!C31,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0))=""退课"",INDEX('0-教学过程登记表（填写+打印)'!C31,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0))=""缓考""),0,INDEX('0-教学过程登记表（填写+打印)'!C31,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0))))"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC1="""","""",IF('0-教学过程登记表（填写+打印)'!R[2]C[24]="""",""正常"",INDEX('0-教学过程登记表（填写+打印)'!C32,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0))))"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC1="""","""",IF('0-教学过程登记表（填写+打印)'!R[2]C[24]="""",""正考"",INDEX('0-教学过程登记表（填写+打印)'!C33,MATCH(RC1,'0-教学过程登记表（填写+打印)'!C2,0))))"
    Range("A4:I4").Select
    Range("I4").Activate
    Selection.AutoFill Destination:=Range("A4:I" & MaxLineCout), Type:=xlFillDefault
    Range("A4:I" & MaxLineCout).Select
    Columns("J:T").Select
    Selection.ClearContents
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
End Sub
Sub 实验成绩表公式()
On Error Resume Next
    Worksheets("课内实验成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("U3").Select
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "序号"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "学号"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "文本学号"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "姓名"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "实验成绩"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "人数"
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "学号"
    Range("U4").Select
    ActiveCell.FormulaR1C1 = "姓名"
    Range("U5").Select
    ActiveCell.FormulaR1C1 = "实验成绩"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""序号"",1,R[-1]C+1)"
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

Sub 评价环节比例设置公式()
On Error Resume Next
    Worksheets("评价环节比例设置").Activate
    Sheets("评价环节比例设置").Visible = 1
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-8]C2:R3C7,MATCH(R[-1]C&""*"",R2C2:R2C7,0))"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-8]C2:R3C7,MATCH(R[-1]C&""*"",R2C2:R2C7,0))"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "评价环节"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[4]"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[5]"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[5]"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[6]"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!RC[7]"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "折合平时"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "取消考试资格的最低平时成绩"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "比例"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-课程目标和综合分析（填写）'!R3C1,,MATCH(R2C,'2-课程目标和综合分析（填写）'!R2,0)-1,1,1)/100"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-课程目标和综合分析（填写）'!R3C1,,MATCH(R2C,'2-课程目标和综合分析（填写）'!R2,0)-1,1,1)/100"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-课程目标和综合分析（填写）'!R3C1,,MATCH(R2C,'2-课程目标和综合分析（填写）'!R2,0)-1,1,1)/100"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-课程目标和综合分析（填写）'!R3C1,,MATCH(R2C,'2-课程目标和综合分析（填写）'!R2,0)-1,1,1)/100"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-课程目标和综合分析（填写）'!R3C1,,MATCH(R2C,'2-课程目标和综合分析（填写）'!R2,0)-1,1,1)/100"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET('2-课程目标和综合分析（填写）'!R3C1,,MATCH(R2C,'2-课程目标和综合分析（填写）'!R2,0)-1,1,1)/100"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-1]C[-4]<>""期中成绩"",INDEX(RC[-6]:RC[-1],MATCH(R2C2,R[-1]C[-6]:R[-1]C[-1],0))+INDEX(RC[-6]:RC[-1],MATCH(R2C4,R[-1]C[-6]:R[-1]C[-1],0))+INDEX(RC[-6]:RC[-1],MATCH(R2C5,R[-1]C[-6]:R[-1]C[-1],0))+INDEX(RC[-6]:RC[-1],MATCH(R2C6,R[-1]C[-6]:R[-1]C[-1],0)),INDEX(RC[-6]:RC[-1],MATCH(R2C2,R[-1]C[-6]:R[-1]C[-1],0))+INDEX(RC[-6]:RC[-1],MATCH(R2C5,R[-1]C[-6]:R[-1]C[-1],0))+I" & _
        "NDEX(RC[-6]:RC[-1],MATCH(R2C6,R[-1]C[-6]:R[-1]C[-1],0)))" & _
        ""
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "30"
    Range("B7:E7").Select
    ActiveCell.FormulaR1C1 = "教学过程登记表最后一行成绩比例及签字栏"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "考核成绩"
    Range("B8:N8").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(R[-4]C,R[-3]C,R[-3]C[3],R[-3]C[4],R[-3]C[1],R[-3]C[2],R[-3]C[5],R[-2]C)"
    Range("B9:E9").Select
    ActiveCell.FormulaR1C1 = "成绩录入表表头成绩比例"
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "实验"
    Range("B11").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-8]C2:R3C7,MATCH(R[-1]C&""*"",R2C2:R2C7,0))"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "期中"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-8]C2:R3C7,MATCH(R[-1]C&""*"",R2C2:R2C7,0))"
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "平时"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "=评价环节比例设置!R[-8]C[4]*100"
    Range("E10").Select
    ActiveCell.FormulaR1C1 = "考核"
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "=评价环节比例设置!R[-8]C[2]*100"
    Range("B12").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=0,"""",R[-2]C&""*""&R[-1]C&""%+"")"
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=0,"""",R[-2]C&""*""&R[-1]C&""%+"")"
    Range("D12").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=0,"""",R[-2]C&""*""&R[-1]C&""%+"")"
    Range("E12").Select
    ActiveCell.FormulaR1C1 = "=R[-2]C&""*""&R[-1]C&""%）"""
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets("评价环节比例设置").Visible = 0
End Sub
Sub 设置质量分析报告格式()
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
    Worksheets("4-质量分析报告（填写+打印）").Activate
     '取消保护
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Rows("3:16").Select
    Range("P16").Activate
    Selection.RowHeight = 22
    i = 0
    Do
        Call 设置字号("D17:P17", ZihaoArray(i))
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
        Call 设置字号("D18:P18", ZihaoArray(i))
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
        Call 设置字号("D19:P19", ZihaoArray(i))
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
        Call 设置字号("D20:P20", ZihaoArray(i))
        LenCell = Len(Range("D20").Value)
        If (i < 4) Then
            i = i + 1
        End If
        If (LenCell <= LenArray2(i)) Or (i = 3) Then
            Exit Do
        End If
    Loop
    
    '保护工作表
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Application.ScreenUpdating = True
End Sub
Sub 设置字号(Cell As String, Zihao As Double)
  Range(Cell).Select
    With Selection.Font
        .Name = "宋体"
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
Sub 取消各行颜色()
Dim SumCount As Integer
    SumCount = Application.WorksheetFunction.Count(Range("A6:A185")) + 5
    Range("A6:AG185").FormatConditions.Delete
End Sub
Sub 设置各行颜色()
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
        "=AND(MOD($A6,2)=1,OR($AE6=""取消"",$AE6=""旷考"",$AE6<60))"
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
        "=AND(MOD($A6,2)=0,OR($AE6=""取消"",$AE6=""旷考"",$AE6<60))"
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

Sub 设置表格线(StartCell As String, EndCell As String, FontSize As Integer)
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
        .FontStyle = "常规"
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
Sub 删除表格线(StartRow As String, EndRow As String)

   '删除表格线
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
Sub 最后一行表格线(EndRow As String)
    '设置最后一行记录的下边格线
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
Sub 设置质量分析报告填写区域颜色()
    Worksheets("4-质量分析报告（填写+打印）").Activate
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
Sub 取消质量分析报告填写区域颜色()
    Worksheets("4-质量分析报告（填写+打印）").Activate
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
Sub 重新设置公式(EndCell As Integer)
    Dim Count As Integer
    Worksheets("0-教学过程登记表（填写+打印)").Activate
     '取消保护
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A2:AG2").Select
    ActiveCell.FormulaR1C1 = _
        "=专业矩阵状态!RC[1]&"" ""&'2-课程目标和综合分析（填写）'!R[2]C[1]&"" 课程(考试/考查/选修)教学过程登记表"""
    Range("AP4:AP5").Select
    ActiveCell.FormulaR1C1 = "=RC[-17]&""得分"""
    Range("B4:B5").Select
    ActiveCell.FormulaR1C1 = "学号"
    Range("C4:C5").Select
    ActiveCell.FormulaR1C1 = "姓名"
    Range("AA4:AA5").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[-2]C[-23]"
    Range("Y4:Y5").Select
    ActiveCell.FormulaR1C1 = "='2-课程目标和综合分析（填写）'!R[-2]C[-13]"
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
    '修正教学过程登记表标题栏“总评成绩”
    Range("AE4:AE5").Select
    ActiveCell.FormulaR1C1 = "总评成绩"
    'Count = Application.WorksheetFunction.Count(Range("A6:A300"))
    '学号
    Range("B6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(ISNA(MATCH('2-课程目标和综合分析（填写）'!R3C2&""-""&RC1,教学过程登记表!C2,0)),""请导入名单"",TEXT(INDEX(教学过程登记表!C[2],MATCH('2-课程目标和综合分析（填写）'!R3C2&""-""&RC1,教学过程登记表!C2,0)),""0000000000"")))"
    '姓名
    Range("C6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]="""",ISNA(MATCH('2-课程目标和综合分析（填写）'!R3C2&""-""&RC1,教学过程登记表!C2,0))),"""",INDEX(教学过程登记表!C[2],MATCH('2-课程目标和综合分析（填写）'!R3C2&""-""&RC1,教学过程登记表!C2,0)))"
    '作业成绩
    Range("Z6").Select
    ActiveCell.FormulaR1C1 = "=IF(RC32=""退课"","""",RC44)"
    
    '平时成绩
    Range("AB6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(ISNA(VLOOKUP(RC[-26],成绩表!C[-9]:C[-4],3,0)),"""",VLOOKUP(RC[-26],成绩表!C[-9]:C[-4],3,0)),IF(RC32=""退课"","""",RC[11]))"
    '实验成绩
    Range("AC6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC32=""退课"",RC2="""",INDEX(评价环节比例设置!R3,MATCH(R4C,评价环节比例设置!R2,0))=0),"""",IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(ISNA(VLOOKUP(RC[-27],成绩表!C[-10]:C[-5],4,0)),"""",VLOOKUP(RC[-27],成绩表!C[-10]:C[-5],4,0)),IF(ISNA(VLOOKUP(RC2,课内实验成绩表!C18:C20,MATCH(""实验成绩"",课内实验成绩表!R1C18:R1C20,0),0)),""缺"",VLOOKUP(RC2,课内实验成绩表!C18:C20,MATCH(""实验成绩"",课内实验成绩表!R1C18:R1C20,0),0))))"
    '考核成绩
    Range("AD6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(OR(RC32=""退课""),"""",IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(ISNA(VLOOKUP(RC[-28],成绩表!C[-11]:C[-6],5,0)),"""",VLOOKUP(RC[-28],成绩表!C[-11]:C[-6],5,0)),IF(ISNA(MATCH('0-教学过程登记表（填写+打印)'!RC2,'1-试卷成绩登记表（填写）'!C2,0)),"""",INDEX('1-试卷成绩登记表（填写）'!C14,MATCH('0-教学过程登记表（填写+打印)'!RC2,'1-试卷成绩登记表（填写）'!C2,0))))))"
    '总评成绩
    Range("AE6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF('2-课程目标和综合分析（填写）'!R3C17=""认证已提交成绩"",IF(ISNA(VLOOKUP(RC[-29],成绩表!C[-12]:C[-7],6,0)),"""",VLOOKUP(RC[-29],成绩表!C[-12]:C[-7],6,0)),IF(RC34="""",IF(RC[2]=""缓考"",RC[2],ROUND(SUM(RC[4]:RC[7]),0)),RC34)))"
    '类别
    Range("AH6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(RC32<>"""",RC32,IF(RC[-1]=""缓考"","""",IF(OR(RC29=""旷考"",RC29=""取消"",RC29=""缺"",RC26=""取消""),""取消"",""""))))"    '期中成绩
    Range("AI6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",RC[-8]*INDEX(评价环节比例设置!R3,MATCH(R4C27,评价环节比例设置!R2,0)))"
    '实验成绩
    Range("AJ6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(OR(RC[-7]="""",RC[-7]=""缺"",RC[-7]=""旷考"",RC[-7]=""取消""),0,RC[-7]*INDEX(评价环节比例设置!R3,MATCH(R4C36,评价环节比例设置!R2,0))))"
    '考核成绩
    Range("AK6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(RC[-7]="""",0,RC[-7]*INDEX(评价环节比例设置!R3,MATCH(R4C37,评价环节比例设置!R2,0))))"
    '折合平时
    Range("AL6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[1]="""","""",RC[1]*INDEX(评价环节比例设置!R3,MATCH(R4C38,评价环节比例设置!R2,0)))"
    '折合平时成绩
    Range("AM6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2=FALSE,评价环节比例设置!R3C8=0),"""",IF(OR(RC32=""取消""),0,IF(R4C27=""期中成绩"",ROUND(SUM(RC[2]:RC[4])/INDEX(评价环节比例设置!R3,MATCH(R4C38,评价环节比例设置!R2,0)),0),ROUND((SUM(RC[2]:RC[4])+RC[-12]*INDEX(评价环节比例设置!R3,MATCH(R4C27,评价环节比例设置!R2,0)))/INDEX(评价环节比例设置!R3,MATCH(R4C38,评价环节比例设置!R2,0)),0))))"
    '课堂测验平均分
    Range("AN6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(OR(RC32=""退课"",RC32=""取消""),0,INDEX(平时成绩表!C27,MATCH(RC2,平时成绩表!C2,0))))"
    '课堂测验
    Range("AO6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]="""","""",RC[-1]*INDEX(评价环节比例设置!R3,MATCH(R4C41,评价环节比例设置!R2,0)))"
    '课程报告
    Range("AP6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(RC[-17]="""",0,IF(ISNUMBER(RC[-17]),RC[-17]*INDEX(评价环节比例设置!R3,MATCH(MID(R4C42,1,4),评价环节比例设置!R2,0)),INDEX(C[7],MATCH(RC[-17],C48,0))*INDEX(评价环节比例设置!R3,MATCH(MID(R4C42,1,4),评价环节比例设置!R2,0)))))"
    '作业成绩
    Range("AQ6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[1]="""","""",IF(RC[1]=""取消"",0,RC[1]*INDEX(评价环节比例设置!R3,MATCH(R4C43,评价环节比例设置!R2,0))))"
    '作业成绩平均分
    Range("AR6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""请导入名单""),"""",IF(OR(RC32=""退课"",RC32=""取消""),0,IF(VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0)=0,""取消"",IF('2-课程目标和综合分析（填写）'!R5C[-42]<>""唐士杰"",VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0),IF(AND(RC[-23]<>"""",VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0)<>""取消""),IF(VL" & _
        "OOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0)+RC[-23]/0.2>100,100,VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0)+RC[-23]/0.2),VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0))))))" & _
        ""
    Range("AH6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC2="""",RC2=""请导入名单""),"""",IF(RC32<>"""",RC32,IF(RC[-1]=""缓考"","""",IF(OR(RC29=""旷考"",RC29=""取消"",RC29=""缺"",RC26=""取消""),""取消"",""""))))"
    Range("AR6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC2="""","""",IF(OR(RC32=""退课"",RC32=""取消""),0,IF(VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0)=0,""取消"",IF('2-课程目标和综合分析（填写）'!R5C[-42]<>""唐士杰"",VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0),IF(AND(RC[-23]<>"""",VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0)<>""取消""),IF(VLOOKUP(RC[-42],平时成绩" & _
        "表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0)+RC[-23]/0.2>100,100,VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0)+RC[-23]/0.2),VLOOKUP(RC[-42],平时成绩表!C[-42]:C[-18],MATCH(R4C44,平时成绩表!R4C2:R4C26,0),0))))))" & _
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
Sub 导入教学任务()
    Dim FileName As Variant
     '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim temp As String
    Dim JuzhenSheets As String
    Dim strCurPath As String    ' 应用程序的当前路径.
    Dim strWbkPath As String    ' 工作簿所在的路径.
    Application.ScreenUpdating = False
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    
    ThisWorkBookName = ThisWorkbook.Name
    FileName = ThisWorkbook.Path
    SourceWorkBook = "数据源-教学任务.xls"
    FileName = FileName & "\" & SourceWorkBook
    If MyFile.FileExists(FileName) = False Then
      Call MsgInfo(NoMsgBox, "数据源-教学任务.xls不存在，请将数据源-教学任务.xls复制到当前文件夹")
      FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
    '调用Windows打开文件对话框
    If (FileName = False) Then Exit Sub
    SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    On Error Resume Next
    If Sheets("教学任务") Is Nothing Then
        Worksheets("毕业要求-指标点数据表").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "教学任务"
    End If
    Worksheets("教学任务").Activate
    Worksheets("教学任务").Visible = True
    Call CopySheet(FileName, SourceWorkBook, "教学任务", "A:M", ThisWorkBookName, "教学任务", "A:M")
    Application.ScreenUpdating = False
End Sub '选择打开文件后并没有真实的把它打开
Sub 导入毕业要求矩阵(Major As String)
    Dim ThisWorksheet As String
    Dim FileName As Variant
     '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim temp As String
    Dim MatrixSheet As String
    Dim strCurPath As String    ' 应用程序的当前路径.
    Dim strWbkPath As String    ' 工作簿所在的路径.
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    ThisWorksheet = ActiveSheet.Name
    On Error Resume Next
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    
    ThisWorkBookName = ThisWorkbook.Name
    FileName = ThisWorkbook.Path
    SourceWorkBook = "数据源-" & Major & "-指标点数据矩阵.xls"
    FileName = FileName & "\" & SourceWorkBook
    If MyFile.FileExists(FileName) = False Then
      Call MsgInfo(NoMsgBox, "数据源-" & Major & "-指标点数据矩阵.xls不存在，请将数据源-" & Major & "-指标点数据矩阵.xls复制到当前文件夹")
      FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
    '调用Windows打开文件对话框
    If (FileName = False) Then Exit Sub
    SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    On Error Resume Next
    Worksheets("专业矩阵状态").Activate
    MatrixSheet = Application.Index(Range("C4:C" & MajorLastRow), Application.Match(Major, Range("B4:B" & MajorLastRow), 0))
    If Sheets(MatrixSheet) Is Nothing Then
        Worksheets("毕业要求-指标点数据表").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = MatrixSheet
    End If
    Worksheets(MatrixSheet).Activate
    Worksheets(MatrixSheet).Visible = True
    Call CopySheet(FileName, SourceWorkBook, "指标点课程矩阵", "A:AS", ThisWorkBookName, MatrixSheet, "A:AS")
    Application.ScreenUpdating = False
    Worksheets(ThisWorksheet).Activate
End Sub '选择打开文件后并没有真实的把它打开
Sub 提取毕业要求矩阵信息(Major As String)
    Dim ThisWorksheet As String
    Dim FileName As Variant
     '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim strCurPath As String    ' 应用程序的当前路径.
    Dim strWbkPath As String    ' 工作簿所在的路径.
    Dim CourseCount As Integer
    Dim PointCount As Integer
    
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    ThisWorksheet = ActiveSheet.Name
    On Error Resume Next
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    ThisWorkBookName = ThisWorkbook.Name
    FileName = ThisWorkbook.Path
    SourceWorkBook = "数据源-" & Major & "-指标点数据矩阵.xls"
    FileName = FileName & "\" & SourceWorkBook
    Sheets("专业矩阵状态").Visible = True
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Worksheets("专业矩阵状态").Activate
    If MyFile.FileExists(FileName) = False Then
        '指标点数据矩阵文件不存在
        Range("D" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = "不存在"
    Else
        Range("D" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = "存在"
    End If
    Workbooks.Open FileName
    Workbooks(SourceWorkBook).Activate
    Application.ScreenUpdating = False

    Worksheets("指标点课程矩阵").Activate
    CourseCount = Application.CountA(Range("D4:D200"))
    PointCount = Application.CountA(Range("E4:AS200"))
    Workbooks(ThisWorkBookName).Activate
    Worksheets("专业矩阵状态").Activate
    Range("E" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = CourseCount
    Range("F" & Application.Match(Major, Range("B1:B" & MajorLastRow), 0)).Value = PointCount
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Sheets("专业矩阵状态").Visible = False
    Workbooks(SourceWorkBook).Activate
    If ActiveWorkbook.Name = SourceWorkBook Then ActiveWorkbook.Close True
End Sub
Sub 导入学生名单()
    Dim FileName As Variant
     '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
    Dim ThisWorkBookName As String
    Dim Count As Integer
    Dim CourseNumber As String
    Dim SourceWorkBook As String
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Dim temp As String
    Dim strCurPath As String    ' 应用程序的当前路径.
    Dim strWbkPath As String    ' 工作簿所在的路径.
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    Worksheets("2-课程目标和综合分析（填写）").Activate
    
    ThisWorkBookName = ThisWorkbook.Name
    CourseNumber = Range("$B$3").Value
    FileName = ThisWorkbook.Path
    SourceWorkBook = "数据源-学生名单.xls"
    FileName = FileName & "\" & SourceWorkBook
    If MyFile.FileExists(FileName) = False Then
      Call MsgInfo(NoMsgBox, "数据源-学生名单.xls不存在，请手动指定")
      FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
    '调用Windows打开文件对话框
    If (FileName = False) Then Exit Sub
    SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:H", ThisWorkBookName, "学生名单", "A:H")
    Worksheets("学生名单").Visible = True
    Worksheets("学生名单").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Count = Application.CountA(Range("C:C"))
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=RC2&""-""&RC6"
    Range("G1:G1").Select
    Selection.AutoFill Destination:=Range("G1:G" & Count), Type:=xlFillDefault
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Worksheets("学生名单").Visible = False
    Application.ScreenUpdating = False
End Sub '选择打开文件后并没有真实的把它打开
Sub 导入成绩表()
     '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
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
    Dim strCurPath As String    ' 应用程序的当前路径.
    Dim strWbkPath As String    ' 工作簿所在的路径.
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
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    Worksheets("2-课程目标和综合分析（填写）").Activate
    CourseNumber = Range("$B$3").Value
    FileName = ThisWorkbook.Path
    SourceWorkBook = "成绩表-" & CourseNumber & ".xls"
    FileName = FileName & "\" & SourceWorkBook
    '成绩表-1720835
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    Worksheets("2-课程目标和综合分析（填写）").Activate
    If (Range("Q3").Value = "认证未提交成绩") Then
        Call MsgInfo(NoMsgBox, "【课程目标和综合分析工作表认证状态为未提交成绩，不需要导入成绩表】")
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
    If (ThisSheetName = "1-试卷成绩登记表（填写）") Then
        Call 试卷成绩登记表公式
    ElseIf (ThisSheetName = "成绩核对") Then
        Call 试卷成绩登记表公式
        Call 成绩核对表公式
    End If
    If MyFile.FileExists(FileName) = False Then
        Call MsgInfo(NoMsgBox, "成绩表-" & CourseNumber & ".xls" & "不存在，请手动指定")
        Msg = "已经在教务系统提交成绩，成绩表导入要求：" & vbCr
        Msg = Msg & "成绩表标题栏需要包含【学号、姓名、实验成绩、平时成绩、考核成绩、总评成绩、成绩类别、作业成绩、课堂测验、课程报告】等关键词；" & vbCr
        Call MsgInfo(NoMsgBox, Msg)
        FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
        '调用Windows打开文件对话框
        If (FileName = False) Then Exit Sub
        SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:N", ThisWorkBookName, "成绩表", "B:O")

    Application.ScreenUpdating = False
    Worksheets("成绩表").Visible = True
    Worksheets("成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Columns("A:A").Select
    Selection.ClearContents
    Worksheets("成绩表").Activate
    For i = 0 To 5
         Cells(1, i + 23).Value = ScoreType(i)
    Next i
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    Call 成绩表公式
    Worksheets("成绩表").Visible = True
    Worksheets("成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:O600").Select
    Selection.UnMerge

    Range(KeyRow & "3:" & Key1Col & "14").Select
    Selection.ClearContents
    Range(Key2Col & "3:" & Key2Col & "14").Select
    Selection.ClearContents
    NumExitFlag = False
    '确定成绩表的栏数

    For j = 1 To 10
        If Not (isError(Application.Match("*学*号*", Range("A" & j & ":" & "O" & j), 0))) Then
            NumRow = j
            CountNum = Application.WorksheetFunction.CountIf(Range("A" & j & ":" & "O" & j), "*学*号*")
            Num2Col = Application.Match("*总*评*", Range("A" & j & ":" & "O" & j), 0) + 2
        End If
        If Not (isError(Application.Match("*序*号*", Range("A" & j & ":" & "O" & j), 0))) Then
            NumExitFlag = True
        End If
    Next j
    '成绩表为学分制管理系统成绩表格式，没有序号列，需要增加序号列
    If (Not NumExitFlag) Then
        Range("A" & NumRow).Value = "序号"
        Range("A" & NumRow + 1).Select
        ActiveCell.FormulaR1C1 = "=IF(RC[1]="""","""",IF(R[-1]C[1]=""学号"",1,R[-1]C+1))"
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
    Worksheets("成绩表").Visible = False
    Worksheets("1 - 试卷成绩登记表(填写)").Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub '选择打开文件后并没有真实的把它打开。

Sub 导入教务系统成绩表()
    Dim FileName As Variant
     '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
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
    Dim strCurPath As String    ' 应用程序的当前路径.
    Dim strWbkPath As String    ' 工作簿所在的路径.
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
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    Worksheets("2-课程目标和综合分析（填写）").Activate
    CourseNumber = Range("$B$3").Value
    FileName = ThisWorkbook.Path
    SourceWorkBook = CourseNumber & "期末成绩.xls"
    FileName = FileName & "\" & SourceWorkBook

    For i = 0 To 5
        ScoreType(i) = Cells(2, 2 * i + 4).Value
    Next i
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    ThisWorkBookName = ThisWorkbook.Name
    If (ThisSheetName = "1-试卷成绩登记表（填写）") Then
        Call 试卷成绩登记表公式
    ElseIf (ThisSheetName = "成绩核对") Then
        Call 试卷成绩登记表公式
        Call 成绩核对表公式
    End If
    If MyFile.FileExists(FileName) = False Then
        Call MsgInfo(NoMsgBox, CourseNumber & "期末成绩.xls" & "不存在，请手动指定")
        Msg = "在学分制管理系统提交成绩前，请下载成绩表，在成绩核对工作表可直接导入，成绩表标题栏需要包含【学号、姓名、实验成绩、平时成绩、考核成绩、总评成绩、成绩类别】等关键词；" & vbCr
        Call MsgInfo(NoMsgBox, Msg)
        FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
        '调用Windows打开文件对话框
        If (FileName = False) Then Exit Sub
        SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    Worksheets("成绩表").Visible = True
    Worksheets("成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    Range("A1:O600").Select
    Selection.UnMerge
    Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:N", ThisWorkBookName, "成绩表", "B:O")
    Application.ScreenUpdating = False
    Worksheets("成绩表").Visible = True
    Worksheets("成绩表").Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    For i = 0 To 5
         Cells(1, i + 23).Value = ScoreType(i)
    Next i
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Call 成绩表公式
    Worksheets("成绩表").Visible = True
    Worksheets("成绩表").Activate
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
    '确定成绩表的栏数

    For j = 1 To 10
        If Not (isError(Application.Match("学号", Range("A" & j & ":" & "O" & j), 0))) Then
            NumRow = j
            CountNum = Application.WorksheetFunction.CountIf(Range("A" & j & ":" & "O" & j), "学号")
            Num2Col = Application.Match("总评*", Range("A" & j & ":" & "O" & j), 0) + 2
        End If
        If Not (isError(Application.Match("序号", Range("A" & j & ":" & "O" & j), 0))) Then
            NumExitFlag = True
        End If
    Next j
    '成绩表为学分制管理系统成绩表格式，没有序号列，需要增加序号列
    If (Not NumExitFlag) Then
        Range("A" & NumRow).Value = "序号"
        Range("A" & NumRow + 1).Select
        ActiveCell.FormulaR1C1 = "=IF(RC[1]="""","""",IF(R[-1]C[1]=""学号"",1,R[-1]C+1))"
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
    Worksheets("成绩表").Visible = False
    Worksheets(ThisSheetName).Activate
    Application.ScreenUpdating = True
End Sub '选择打开文件后并没有真实的把它打开。
Sub 导入实验成绩表()
    Dim FileName As Variant
     '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
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
    Dim strCurPath As String    ' 应用程序的当前路径.
    Dim strWbkPath As String    ' 工作簿所在的路径.

    Application.ScreenUpdating = False
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    Application.EnableEvents = False
    Worksheets("2-课程目标和综合分析（填写）").Activate
    
    ThisWorkBookName = ThisWorkbook.Name
    CourseNumber = Range("$B$3").Value
    FileName = ThisWorkbook.Path
    SourceWorkBook = "实验成绩表-" & CourseNumber & ".xls"
    FileName = FileName & "\" & SourceWorkBook
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    
    If MyFile.FileExists(FileName) = False Then
        Call MsgInfo(NoMsgBox, "实验成绩表-" & CourseNumber & ".xls" & "不存在，请手动指定")
        FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
        '调用Windows打开文件对话框
        If (FileName = False) Then Exit Sub
        SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
    End If
    
    Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:N", ThisWorkBookName, "课内实验成绩表", "A:N")
    Application.ScreenUpdating = False
    Worksheets("课内实验成绩表").Visible = True
    Worksheets("课内实验成绩表").Activate
    Call 实验成绩表公式
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
    Worksheets("课内实验成绩表").Visible = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub '选择打开文件后并没有真实的把它打开。
Function 导入教学过程登记表() As Boolean
    Dim FileName As Variant
     '打开文件对话框返回的文件名，是一个全路径文件名，其值也可能是False，因此类型为Variant
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
    Dim strCurPath As String    ' 应用程序的当前路径.
    Dim strWbkPath As String    ' 工作簿所在的路径.
    strCurPath = CurDir$
    strWbkPath = ThisWorkbook.Path
    On Error Resume Next
    ' /* 应用程序的当前路径不是工作簿所在的路径. */
    If StrComp(strCurPath, strWbkPath) <> 0 Then
        ChDrive strWbkPath
        ChDir strWbkPath
    End If
    Worksheets("2-课程目标和综合分析（填写）").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ThisWorkBookName = ThisWorkbook.Name
    Xueqi = Range("$AG$1").Value
    If Range("B4").Value = "请在数据源-教学任务中添加该课号的课程信息" Then
        Application.EnableEvents = True
        导入教学过程登记表 = False
        Exit Function
    Else
        SourceWorkBook = "数据源-" & Xueqi & "学期教学过程登记表.xls"
        FileName = ThisWorkbook.Path & "\" & SourceWorkBook
        If MyFile.FileExists(FileName) = False Then
            Call MsgInfo(NoMsgBox, FileName & "不存在，请手动指定" & Xueqi & "学期的教学过程登记表，支持导入学分管理系统下载成绩表名单，也可以从教务系统复制教学过程登记表名单到新EXCEL文档")
            FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
            If FileName = False Then
                Sheets("教学过程登记表").Visible = False
                Application.EnableEvents = True
                导入教学过程登记表 = False
                Exit Function
            End If
            '调用Windows打开文件对话框
            SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
        End If
        Sheets("教学过程登记表").Visible = 1
        Worksheets("教学过程登记表").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Range("A1:Z40000").Select
        Selection.ClearContents
        Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:C", ThisWorkBookName, "教学过程登记表", "C:E")
        Application.ScreenUpdating = False
        Sheets("教学过程登记表").Visible = 1
        Worksheets("教学过程登记表").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Range("A4").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(MID(R[-3]C[2],1,4)=""任课老师"",MID(R[-3]C3,FIND(""课程序号:"",R[-3]C[2],1)+5,LEN(R[-3]C3)-FIND(""课程序号:"",R[-3]C[2],1)),R[-1]C)"
        Range("B4").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(OR(RC[1]=""序号"",RC[1]="""",RC[3]=""""),"""",RC[-1]&""-""&RC[1])"
        Range("A4:B4").Select
        Selection.AutoFill Destination:=Range("A4:B50000"), Type:=xlFillDefault
        Range("A4:B50000").Select
        
        Worksheets("2-课程目标和综合分析（填写）").Activate
        If (Range("$T$1").Value = "无名单") Then
            Call MsgInfo(NoMsgBox, "导入的教学过程登记表中没有该课号的名单，请手动指定教学过程登记表文件，支持导入学分管理系统下载成绩表名单，也可以从教务系统复制教学过程登记表名单到新EXCEL文档")
            FileName = Application.GetOpenFilename("Microsoft Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "Get list")
            '调用Windows打开文件对话框
            If (FileName = False) Then
                Application.EnableEvents = True
                导入教学过程登记表 = False
                Exit Function
            End If
          SourceWorkBook = Mid(FileName, InStrRev(FileName, "\") + 1)
          Sheets("教学过程登记表").Visible = 1
          Worksheets("教学过程登记表").Activate
          Range("C1:Z40000").Select
          Selection.ClearContents
          Selection.UnMerge
          Call CopySheet(FileName, SourceWorkBook, "Sheet1", "A:C", ThisWorkBookName, "教学过程登记表", "C:E")
          Application.ScreenUpdating = False
          Worksheets("2-课程目标和综合分析（填写）").Activate
        End If
        CourseNum = Range("B3").Value

        Sheets("教学过程登记表").Visible = 1
        Worksheets("教学过程登记表").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        
        '导入的名单为教务系统下载的成绩表名单格式
        If Not (isError(Application.Match("*" & CourseNum & "*", Range("K:K"), 0))) Then
            Count = Application.WorksheetFunction.CountA(Range("C4:C200"))
            Range("C1:K200").Select
            Selection.UnMerge
            Columns("C:C").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("C3").Select
            ActiveCell.FormulaR1C1 = "序号"
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
            Call 重新设置公式(Count)
            
        End If
        Worksheets("2-课程目标和综合分析（填写）").Activate
        
        FirstRow = Range("$T$1").Value
        LastRow = Range("$W$1").Value
        Worksheets("教学过程登记表").Activate
        If (LastRow <> "无名单") Then
          Range("A" & LastRow & ":Z40000").Select
          Selection.ClearContents
        End If
        If (FirstRow <> 0) Then
            Range("A1:Z" & FirstRow).Select
            Selection.ClearContents
            Selection.Delete Shift:=xlUp
        End If
        Sheets("教学过程登记表").Visible = 0
    End If
    Worksheets("教学过程登记表").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Call 设置教学过程登记表
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Worksheets("0-教学过程登记表（填写+打印)").Activate
    导入教学过程登记表 = True
End Function '选择打开文件后并没有真实的把它打开。
Sub 新建专业矩阵状态工作表()
    Dim AllowEditCount As Integer
    On Error Resume Next
    Worksheets("专业矩阵状态").Activate
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "学院，专业及指标点数据矩阵设置及状态"
    Range("A2").Select
    Selection.FormulaR1C1 = "学院名称"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "序号"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "专业名称"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "矩阵工作表名"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "文件状态"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "课程门数"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "指标点数"
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
        .Name = "宋体"
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
        .Name = "宋体"
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
    ActiveCell.FormulaR1C1 = "=IF(RC[1]="""","""",IF(R[-1]C=""序号"",1,R[-1]C+1))"
    Range("A4").Select
    Selection.AutoFill Destination:=Range("A4:A12"), Type:=xlFillDefault
    Range("A4:A12").Select
    
    AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
    If (AllowEditCount <> 0) Then
        For i = 1 To AllowEditCount
            Sheets("专业矩阵状态").Protection.AllowEditRanges(1).Delete
        Next i
    End If
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "版本号"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "修订日期"
    Range("H1:H2").Select
    ActiveSheet.Protection.AllowEditRanges.Add Title:="专业", Range:=Range("B4:C12")
    ActiveSheet.Protection.AllowEditRanges.Add Title:="学院", Range:=Range("B2:D2")
    
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

'修订日期：2019年4月15日
''成绩导入表  修订日期：2019年1月25日
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
    Sheets("成绩录入").Visible = True
    Sheets("成绩录入").Select
    FolderName = ThisWorkbook.Path
    MobanWorkbookName = ThisWorkbook.Name
    
    Sheets("2-课程目标和综合分析（填写）").Select
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
    CourseNumber = Range("B3").Value
    Term = Range("AG2").Value
    CourseName = Range("B4").Value
    ProportionNomal = Range("T4").Value
    ProportionExperiment = Range("U4").Value
    ProportionMidterm = Range("W4").Value
    ProportionExamine = Range("V4").Value
    Count = Range("S3").Value
    Sheets("2-课程目标和综合分析（填写）").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    
    CourseFileName = CourseNumber & "成绩录入表.xls"
    
    Call CreateNewWorkbook(FolderName, CourseFileName)
    CourseWorkbookName = ThisWorkbook.Name
    Range("1:200").Select
    Selection.UnMerge
    
    Workbooks(MobanWorkbookName).Activate
    Sheets("成绩录入").Select
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
    If (ScoreCategory = "退课") Then
      Rows(i & ":" & i).Select
      Selection.Delete
    End If
    ExamineCategory = Range("I" & i).Value
    If (ScoreCategory = "缓考") Then
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
        .Name = "宋体"
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
    Sheets("成绩录入").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
    Sheets("成绩录入").Visible = False
    
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
  '如果工作簿存在,则返回True
  FileExists = Len(Dir(FullFileName)) > 0
End Function
Sub CreateNewSheet(SheetName As String)
  On Error Resume Next
  '若SheetName的工作表不存在,则新建一个工作表
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
Sub 课程目标和综合分析(Mode As String, TargetValue As String, TargetRow As String, TargetColumn As String)
    Dim ImportStatus As Boolean
    Dim TempMsgBox As Boolean
    Select Case Mode
        Case "认证状态"
            If TargetValue = "非认证" Then
                Call 非认证
            ElseIf TargetValue = "认证未提交成绩" Then
                Call 认证未提交成绩
            ElseIf TargetValue = "认证已提交成绩" Then
                Call 认证已提交成绩
            End If
        Case "课程序号"
            If TargetValue <> "" Then
                If Range("B4").Value = "请在数据源-教学任务中添加该课号的课程信息" Then
                    Call 导入教学任务
                    Worksheets("2-课程目标和综合分析（填写）").Activate
                End If
                If Range("Q3").Value = "认证未提交成绩" Or Range("Q3").Value = "非认证" Then
                    ImportStatus = 导入教学过程登记表
                    Worksheets("0-教学过程登记表（填写+打印)").Activate
                    SumCount = Range("AM1").Value
                    If (Application.WorksheetFunction.Count(Range("A6:A200")) <> SumCount) And SumCount <> 0 Then
                        Worksheets("0-教学过程登记表（填写+打印)").Activate
                        Call 设置教学过程登记表
                    End If
                ElseIf Range("Q3").Value = "认证已提交成绩" Then
                    TempMsgBox = NoMsgBox
                    NoMsgBox = True
                    ImportStatus = 导入教学过程登记表
                    NoMsgBox = TempMsgBox
                End If
                Worksheets("2-课程目标和综合分析（填写）").Activate
            End If
        Case "认证专业"
            If TargetValue <> "" Then
                Sheets("专业矩阵状态").Visible = True
                Worksheets("专业矩阵状态").Activate
                If Not isError(Application.Match(TargetValue, Range("B4:B" & MajorLastRow), 0)) Then
                    Sheets("专业矩阵状态").Visible = False
                    Worksheets("2-课程目标和综合分析（填写）").Activate
                    If (Range("B6").Value = 0) Then
                        Application.ScreenUpdating = False
                        Application.EnableEvents = False
                        Call 导入学生名单
                        Application.ScreenUpdating = True
                        Application.EnableEvents = True
                        Worksheets("2-课程目标和综合分析（填写）").Activate
                    End If
                    Application.ScreenUpdating = False
                    Call 导入矩阵
                    Call 指标点数据表公式
                    Call 毕业要求数据表公式
                    Application.ScreenUpdating = True
                    Worksheets("2-课程目标和综合分析（填写）").Activate
                End If
            End If
    End Select
End Sub
Sub 非认证()
    Dim WorkBookName As String
      '不需要进行达成度评价的课程，只需要填写教学过程登记表和2-课程目标和综合分析（填写）表中的课号，评价环节比例等信息，打印教学过程登记表和质量分析报告。
        Application.ScreenUpdating = False
        Call 设置表格主题
        WorkBookName = ThisWorkbook.Name
        Workbooks(WorkBookName).Activate
        Worksheets("2-课程目标和综合分析（填写）").Activate
        '隐藏"2-课程目标和综合分析（填写）"达成度评价部分表格
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("2-课程目标和综合分析（填写）").Rows("10:28").Select
        Selection.EntireRow.Hidden = True
        Call 课程目标允许编辑区域
        Sheets("1-课程目标达成度评价（打印）").Visible = True
        Worksheets("1-课程目标达成度评价（打印）").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("1-课程目标达成度评价（打印）").Rows("11:20").Select
        Selection.EntireRow.Hidden = False
        SumRow = Application.WorksheetFunction.CountA(Range("B11:B20"))
        Sheets("1-课程目标达成度评价（打印）").Rows(SumRow + 11 & ":20").Select
        Selection.EntireRow.Hidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        '隐藏3-毕业要求数据表（填写），1-课程目标达成度评价（打印），2-毕业要求达成度评价（打印）和3-综合分析（打印）工作表
        Sheets("0-教学过程登记表（填写+打印)").Visible = True
        Sheets("1-试卷成绩登记表（填写）").Visible = True
        Sheets("4-质量分析报告（填写+打印）").Visible = True
        Sheets("成绩核对").Visible = True
        Sheets("3-毕业要求数据表（填写）").Visible = False
        Sheets("1-课程目标达成度评价（打印）").Visible = False
        Sheets("2-毕业要求达成度评价（打印）").Visible = False
        Sheets("3-综合分析（打印）").Visible = False
        
        Sheets("成绩录入").Visible = False
        Sheets("教学过程登记表").Visible = False
        Sheets("学生名单").Visible = False
        Sheets("课内实验成绩表").Visible = False
        Sheets("平时成绩表").Visible = False
        Sheets("成绩表").Visible = False
        Sheets("评价环节比例设置").Visible = False
        Sheets("毕业要求-指标点数据表").Visible = False
        Sheets("课程目标达成度汇总用数据").Visible = False
        Sheets("毕业要求达成度汇总用数据").Visible = False
        '1-试卷成绩登记表（填写）工作表中的其中成绩，平时成绩，实验成绩，课堂测验，课程报告列隐藏
        '1-试卷成绩登记表（填写）工作表的考核成绩列允许编辑，但不删除公式
        
        Workbooks(WorkBookName).Activate
        Worksheets("1-试卷成绩登记表（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("1-试卷成绩登记表（填写）").Range("R1:X1").Select
        Selection.UnMerge
        Sheets("1-试卷成绩登记表（填写）").Columns("V:Y").Select
        Selection.EntireColumn.Hidden = True
        Sheets("1-试卷成绩登记表（填写）").Range("R1:X1").Select
        Selection.Merge
        Sheets("1-试卷成绩登记表（填写）").Range("N4:N" & MaxLineCout).Select
        Selection.Locked = True
        Selection.FormulaHidden = False
        With Selection.Font
          .Name = "Calibri"
          .FontStyle = "加粗"
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
        '统计允许编辑区域个数，并全部删除
        AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
        If (AllowEditCount <> 0) Then
            For i = 1 To AllowEditCount
                Sheets("1-试卷成绩登记表（填写）").Protection.AllowEditRanges(1).Delete
            Next i
        End If

        On Error Resume Next
        ActiveSheet.Protection.AllowEditRanges.Add Title:="大题名称", Range:=Range("E2:M2")
        ActiveSheet.Protection.AllowEditRanges.Add Title:="考试大题分数", Range:=Range("E3:M" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="考核成绩", Range:=Range("N4:N" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="其他成绩", Range:=Range("O4:R" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="认证方式", Range:=Range("Z2")
        Worksheets("1-试卷成绩登记表（填写）").Activate
        Sheets("1-试卷成绩登记表（填写）").Range("N4").Select
        ActiveCell.FormulaR1C1 = _
          "=IF(OR(SUM(RC[-9]:RC[-1])=0,RC2=""""),"""",SUM(RC[-9]:RC[-1]))"
        Sheets("1-试卷成绩登记表（填写）").Range("N4").Select
        Selection.AutoFill Destination:=Sheets("1-试卷成绩登记表（填写）").Range("N4:N" & MaxLineCout), Type:=xlFillDefault
        Sheets("1-试卷成绩登记表（填写）").Range("N4:N" & MaxLineCout).Select
        Call 试卷成绩登记表核心公式
        
        Worksheets("1-试卷成绩登记表（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("2-课程目标和综合分析（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Application.ScreenUpdating = True
End Sub

Sub 认证未提交成绩()
        Dim WorkBookName As String
        Dim Msg As String
        Dim SchoolName As String
        Application.ScreenUpdating = False
        Worksheets("专业矩阵状态").Visible = True
        Worksheets("专业矩阵状态").Activate
        SchoolName = Range("B2").Value
        Worksheets("专业矩阵状态").Visible = False
        Call 设置表格主题
        Sheets("2-课程目标和综合分析（填写）").Select
        '恢复"2-课程目标和综合分析（填写）"达成度评价部分表格
        Worksheets("2-课程目标和综合分析（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("2-课程目标和综合分析（填写）").Rows("10:28").Select
        Selection.EntireRow.Hidden = False
        Call 课程目标允许编辑区域
        Sheets("1-课程目标达成度评价（打印）").Visible = True
        Worksheets("1-课程目标达成度评价（打印）").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("1-课程目标达成度评价（打印）").Rows("11:20").Select
        Selection.EntireRow.Hidden = False
        SumRow = Application.WorksheetFunction.CountA(Range("B11:B20"))
        Sheets("1-课程目标达成度评价（打印）").Rows(SumRow + 11 & ":20").Select
        Selection.EntireRow.Hidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        '恢复3-毕业要求数据表（填写），1-课程目标达成度评价（打印），2-毕业要求达成度评价（打印）和3-综合分析（打印）工作表

        Sheets("0-教学过程登记表（填写+打印)").Visible = True
        Sheets("1-试卷成绩登记表（填写）").Visible = True
        Sheets("4-质量分析报告（填写+打印）").Visible = True
        Sheets("成绩核对").Visible = True
        If SchoolName = "计算机信息与安全学院" Then
            Sheets("3-毕业要求数据表（填写）").Visible = False
        Else
            Sheets("3-毕业要求数据表（填写）").Visible = True
        End If
        Sheets("1-课程目标达成度评价（打印）").Visible = False
        Sheets("2-毕业要求达成度评价（打印）").Visible = False
        Sheets("3-综合分析（打印）").Visible = False
        
        Sheets("成绩录入").Visible = False
        Sheets("教学过程登记表").Visible = False
        Sheets("学生名单").Visible = False
        Sheets("课内实验成绩表").Visible = False
        Sheets("平时成绩表").Visible = False
        Sheets("成绩表").Visible = False
        Sheets("评价环节比例设置").Visible = False
        Sheets("毕业要求-指标点数据表").Visible = False
        Sheets("课程目标达成度汇总用数据").Visible = False
        Sheets("毕业要求达成度汇总用数据").Visible = False
        
        Worksheets("1-试卷成绩登记表（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password

        '1-试卷成绩登记表（填写）工作表的考核成绩列不允许编辑
        Sheets("1-试卷成绩登记表（填写）").Range("N4:N" & MaxLineCout).Select
        Selection.Locked = True
        Selection.FormulaHidden = True
        '统计允许编辑区域个数，并全部删除
        AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
        If (AllowEditCount <> 0) Then
            For i = 1 To AllowEditCount
                Sheets("1-试卷成绩登记表（填写）").Protection.AllowEditRanges(1).Delete
            Next i
        End If

        On Error Resume Next
        ActiveSheet.Protection.AllowEditRanges.Add Title:="大题名称", Range:=Range("E2:M2")
        ActiveSheet.Protection.AllowEditRanges.Add Title:="考试大题分数", Range:=Range("E3:M" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="认证方式", Range:=Range("Z2")
        With Selection.Font
          .Name = "Calibri"
          .FontStyle = "加粗"
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
        Sheets("1-试卷成绩登记表（填写）").Columns("O:S").Select
        Selection.EntireColumn.Hidden = False
        Call 试卷成绩登记表核心公式
        Worksheets("1-试卷成绩登记表（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("2-课程目标和综合分析（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Application.ScreenUpdating = True
End Sub

Sub 认证已提交成绩()
    Dim WorkBookName As String
    Dim Msg As String
    Dim SchoolName As String
        Application.ScreenUpdating = False
        Worksheets("专业矩阵状态").Visible = True
        Worksheets("专业矩阵状态").Activate
        SchoolName = Range("B2").Value
        Worksheets("专业矩阵状态").Visible = False
        Call 设置表格主题
        Sheets("2-课程目标和综合分析（填写）").Select
        '恢复"2-课程目标和综合分析（填写）"达成度评价部分表格
        Worksheets("2-课程目标和综合分析（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("2-课程目标和综合分析（填写）").Rows("10:28").Select
        Selection.EntireRow.Hidden = False
        Call 课程目标允许编辑区域
        Sheets("1-课程目标达成度评价（打印）").Visible = True
        Worksheets("1-课程目标达成度评价（打印）").Activate
        ActiveSheet.PageSetup.CenterFooter = ""
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password
        Sheets("1-课程目标达成度评价（打印）").Rows("11:20").Select
        Selection.EntireRow.Hidden = False
        SumRow = Application.WorksheetFunction.CountA(Range("B11:B20"))
        Sheets("1-课程目标达成度评价（打印）").Rows(SumRow + 11 & ":20").Select
        Selection.EntireRow.Hidden = True
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        
        '恢复3-毕业要求数据表（填写），1-课程目标达成度评价（打印），2-毕业要求达成度评价（打印）和3-综合分析（打印）工作表
        Sheets("0-教学过程登记表（填写+打印)").Visible = False
        Sheets("1-试卷成绩登记表（填写）").Visible = True
        Sheets("4-质量分析报告（填写+打印）").Visible = False
        Sheets("成绩核对").Visible = False
        If SchoolName = "计算机信息与安全学院" Then
            Sheets("3-毕业要求数据表（填写）").Visible = False
        Else
            Sheets("3-毕业要求数据表（填写）").Visible = True
        End If
        Sheets("1-课程目标达成度评价（打印）").Visible = False
        Sheets("2-毕业要求达成度评价（打印）").Visible = False
        Sheets("3-综合分析（打印）").Visible = False
        
        Sheets("成绩录入").Visible = False
        Sheets("教学过程登记表").Visible = False
        Sheets("学生名单").Visible = False
        Sheets("课内实验成绩表").Visible = False
        Sheets("平时成绩表").Visible = False
        Sheets("成绩表").Visible = False
        Sheets("评价环节比例设置").Visible = False
        Sheets("毕业要求-指标点数据表").Visible = False
        Sheets("课程目标达成度汇总用数据").Visible = False
        Sheets("毕业要求达成度汇总用数据").Visible = False
        Worksheets("1-试卷成绩登记表（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False, Password:=Password

        Sheets("1-试卷成绩登记表（填写）").Columns("O:S").Select
        Selection.EntireColumn.Hidden = False
        
        '1-试卷成绩登记表（填写）工作表的考核成绩列不允许编辑
        Sheets("1-试卷成绩登记表（填写）").Range("N4:N" & MaxLineCout).Select
        Selection.Locked = True
        Selection.FormulaHidden = True
        '统计允许编辑区域个数，并全部删除
        AllowEditCount = ActiveSheet.Protection.AllowEditRanges.Count
        If (AllowEditCount <> 0) Then
            For i = 1 To AllowEditCount
                Sheets("1-试卷成绩登记表（填写）").Protection.AllowEditRanges(1).Delete
            Next i
        End If

        On Error Resume Next
        ActiveSheet.Protection.AllowEditRanges.Add Title:="大题名称", Range:=Range("E2:M2")
        ActiveSheet.Protection.AllowEditRanges.Add Title:="考试大题分数", Range:=Range("E3:M" & MaxLineCout)
        ActiveSheet.Protection.AllowEditRanges.Add Title:="认证方式", Range:=Range("Z2")
        With Selection.Font
          .Name = "Calibri"
          .FontStyle = "加粗"
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
        Sheets("1-试卷成绩登记表（填写）").Columns("O:S").Select
        Selection.EntireColumn.Hidden = False
        Call 试卷成绩登记表核心公式
        
        Worksheets("1-试卷成绩登记表（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Worksheets("2-课程目标和综合分析（填写）").Activate
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=Password
        Application.ScreenUpdating = True
End Sub

Sub 教学过程登记表()
Dim SumCount As Integer
Dim Xueqi As String
Dim AllowEditCount As Integer
Dim ImportStatus As Boolean
    Application.ScreenUpdating = False
    Worksheets("0-教学过程登记表（填写+打印)").Activate
    Xueqi = Range("AN1").Value
    SumCount = Range("AM1").Value
    If (Application.WorksheetFunction.Count(Range("A6:A200")) <> SumCount) And SumCount <> 0 Then
        Worksheets("0-教学过程登记表（填写+打印)").Activate
        Call 设置教学过程登记表
    ElseIf (SumCount = 0) Then
        If (Xueqi = "") Then
           Exit Sub
        Else
            ImportStatus = 导入教学过程登记表
            If ImportStatus = False Then
                Application.ScreenUpdating = True
                Exit Sub
            End If
            Worksheets("0-教学过程登记表（填写+打印)").Activate
            If (Range("AM1").Value <> 0) Then
                Call 设置教学过程登记表
            End If
        End If
    End If
    Application.ScreenUpdating = True
End
Sub 设置区域颜色(SetSheetName As String, SetRange As String, SetColor As String)
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
'[版本号]V5.06.34


