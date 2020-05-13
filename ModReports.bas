Attribute VB_Name = "ModReports"
'===============================================================
' Module ModReports
' Imports data from MS Project File
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 May 20
'===============================================================
Option Explicit


' ===============================================================
' LookAheadRep
' Opens project file from passed filepath, extracts data and sends
' to display procedure
' ---------------------------------------------------------------
Public Sub LookAheadRep(ProjectPath As String)
    Dim ObjMSProject As Object
    Dim ProgName As String
    Dim ProjName As String
    Dim AryTasks() As String
    Dim PLevel As Integer
    Dim LookAhead As Integer
    Dim Tsk As Task
    Dim i As Integer
    
    On Error GoTo ErrorHandler:
    
    Set ObjMSProject = CreateObject("MSProject.Application")
    
    ModLibrary.PerfSettingsOn
    
    Worksheets("Look Ahead Report").Visible = True
    DeleteSheets
        
    If ObjMSProject Is Nothing Then
      MsgBox "Project is not installed"
      Exit Sub
    End If
    
    With ShtMain
        .Unprotect
        .Range("NO_PROJS") = 0
        .Range("U:U").ClearContents
        .Protect
    End With
    
    With ObjMSProject
        .Visible = False
        .DisplayAlerts = False
        .FileOpen Name:=ProjectPath, ReadOnly:=True
        .OptionsViewEx DisplaySummaryTasks:=True
        .OutlineShowAllTasks
        .FilterApply Name:="All Tasks"
        .AutoFilter
        .AutoFilter
        .Application.CalculateProject
        
    End With
    
    LookAhead = ShtMain.Range("LA_PERIOD")
    PLevel = ShtMain.Range("LEVEL")
    
    'cycle through tasks in plan and add to tasks array
    ReDim AryTasks(1 To ObjMSProject.ActiveProject.Tasks.Count, 1 To 12)
    
    i = 1
    For Each Tsk In ObjMSProject.ActiveProject.Tasks
        If Not Tsk Is Nothing And Tsk.Summary = False Then
            If Tsk.Number1 <= PLevel And Tsk.BaselineFinish < DateAdd("ww", LookAhead, Now) Then
                AryTasks(i, enRef) = tsk.UniqueID
                AryTasks(i, enlevel) = Tsk.Number1
                AryTasks(i, enMileName) = Tsk.Name
                AryTasks(i, enBaseFinish) = Format(Tsk.BaselineFinish, "dd mmm yy")
                AryTasks(i, enForeFinish) = Format(Tsk.Finish, "dd mmm yy")
                AryTasks(i, enDTI) = Tsk.Number13
                AryTasks(i, enRAG) = Tsk.Text22
                AryTasks(i, enLocalRAG) = tsk.Text10
                AryTasks(i, enIssue) = Tsk.Text14
                AryTasks(i, enImpact) = Tsk.Text15
                AryTasks(i, enAction) = Tsk.Text16
                AryTasks(i, enProject) = Tsk.Text8
                AddProjSheets Tsk.Text8
                i = i + 1
            End If
        End If
    Next Tsk
        
    DisplayTasks AryTasks
    
    Worksheets("Look Ahead Report").Visible = False
    
    ModLibrary.PerfSettingsOff
    
    MsgBox "Import Complete", vbOKOnly + vbInformation
    
    ObjMSProject.FileClose (False)
    Set ObjMSProject = Nothing
Exit Sub

ErrorHandler:
    Debug.Print Err.Number & " - " & Err.Description
    ModLibrary.PerfSettingsOff
    ObjMSProject.FileClose (False)
    Set ObjMSProject = Nothing
End Sub

' ===============================================================
' DependancyRep
' Opens project file from passed filepath, extracts data and sends
' to display procedure
' ---------------------------------------------------------------
Public Sub DependancyRep()
    Dim ObjMSProject As Object
    Dim ProgName As String
    Dim ProjName As String
    Dim AryTasks() As Variant
    Dim PLevel As Integer
    Dim LookAhead As Integer
    Dim tsk As Task
    Dim i As Integer
    
    On Error GoTo ErrorHandler:
    
    Set ObjMSProject = CreateObject("MSProject.Application")
    
    ModLibrary.PerfSettingsOn
    
    If ObjMSProject Is Nothing Then
      MsgBox "Project is not installed"
      Exit Sub
    End If
    
    With ObjMSProject
        .Visible = False
        .DisplayAlerts = False
        .FileOpen Name:=[mpp_filepath], ReadOnly:=True
        .OptionsViewEx DisplaySummaryTasks:=True
        .OutlineShowAllTasks
        .FilterApply Name:="All Tasks"
        .AutoFilter
        .AutoFilter
        .Application.CalculateProject
        
    End With
    
    LookAhead = ShtMain.Range("LA_PERIOD")
    PLevel = ShtMain.Range("LEVEL")
    
    'cycle through tasks in plan and add to tasks array
    ReDim AryTasks(1 To ObjMSProject.ActiveProject.Tasks.Count, 1 To 15)
    
    i = 1
    For Each tsk In ObjMSProject.ActiveProject.Tasks
        If Not tsk Is Nothing And tsk.Summary = False Then
            If tsk.BaselineFinish < DateAdd("ww", LookAhead, Now) Then
                Select Case tsk.number1
                    Case 10, 5, 11
                        AryTasks(i, enDLRef) = tsk.UniqueID
                        AryTasks(i, enDLMileName) = tsk.Name
                        AryTasks(i, enDLLevel) = tsk.number1
                        AryTasks(i, enDLBenef) = tsk.Text20
                        AryTasks(i, enDLDonor) = tsk.Text28
                        AryTasks(i, enDLBaseFinish) = Format(tsk.BaselineFinish, "dd mmm yy")
                        AryTasks(i, enDLForeFinish) = Format(tsk.finish, "dd mmm yy")
                        AryTasks(i, enDLRAG) = tsk.Text22
                        AryTasks(i, enDLLocalRAG) = tsk.Text10
                        AryTasks(i, enDLIssue) = tsk.Text14
                        AryTasks(i, enDLImpact) = tsk.Text15
                        AryTasks(i, enDLAction) = tsk.Text16
                        AryTasks(i, enDLProject) = tsk.Text8
                        
                        If tsk.flag18 = True Then AryTasks(i, enDLDepIn) = 1 Else AryTasks(i, enDLDepIn) = 0
                        If tsk.flag19 = True Then AryTasks(i, enDLDepOut) = 1 Else AryTasks(i, enDLDepOut) = 0
                        i = i + 1
                End Select
            End If
        End If
    Next tsk
        
    ShtDepLog.DisplayTasks AryTasks
    
    ModLibrary.PerfSettingsOff
    
    MsgBox "Dependancy Report Complete", vbOKOnly + vbInformation
    
    ObjMSProject.fileclose (False)
    Set ObjMSProject = Nothing
Exit Sub

ErrorHandler:
    Debug.Print Err.Number & " - " & Err.Description
    ModLibrary.PerfSettingsOff
    ObjMSProject.fileclose (False)
    Set ObjMSProject = Nothing
End Sub

' ===============================================================
' AddProjSheets
' Checks to see whether a project sheet exists, if not it creates
' one
' ---------------------------------------------------------------
Public Sub AddProjSheets(ProjName As String)
    Dim Wkst As Worksheet
    Dim SheetExists As Boolean
    
    If ProjName = "" Then Exit Sub
    
    For Each Wkst In Worksheets
        If Wkst.Name = ProjName Then
            SheetExists = True
            Exit For
        End If
    Next
    
    If Not SheetExists Then
        Worksheets("Look Ahead Report").Copy After:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = ProjName
'        ActiveSheet
        With Worksheets(ProjName)
            .Range("A1") = "Milestone Report - " & ProjName
            .Range("B2") = "Ref"
            .Range("C2") = "Level"
            .Range("D2") = "Milestone Name"
            .Range("E2") = "Baseline Finish"
            .Range("F2") = "Forecast Finish"
            .Range("G2") = "DTI"
            .Range("H2") = "RAG"
            .Range("I2") = "Local RAG"
            .Range("J2") = "Issue"
            .Range("K2") = "Impact"
            .Range("L2") = "Action"
            
            .Columns(1).ColumnWidth = 10
            .Columns(2).ColumnWidth = 5
            .Columns(3).ColumnWidth = 5
            .Columns(4).ColumnWidth = 40
            .Columns(5).ColumnWidth = 15
            .Columns(6).ColumnWidth = 15
            .Columns(7).ColumnWidth = 15
            .Columns(8).ColumnWidth = 10
            .Columns(9).ColumnWidth = 10
            .Columns(10).ColumnWidth = 10
            .Columns(11).ColumnWidth = 10
            .Columns(12).ColumnWidth = 10
            .Columns(13).Hidden = True
            With ShtMain
                ShtMain.Unprotect
                .Range("NO_PROJS") = .Range("NO_PROJS") + 1
                .Range("Proj_IND").Offset(.Range("NO_PROJS"), 0) = ProjName
                ShtMain.Protect
            End With
            
        End With
    End If
End Sub

' ===============================================================
' DeleteSheets
' Deletes all project sheets
' ---------------------------------------------------------------
Public Sub DeleteSheets()
    Dim i As Integer
    Dim WSheet As Worksheet
    
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    
    With ShtMain
        .Unprotect
        .Range("NO_PROJS") = 0
        .Range("U:U").ClearContents
        .Protect
    End With
    
    For Each WSheet In Worksheets
        Select Case WSheet.Name
            Case Is = ShtExceptRep.Name
            Case Is = ShtMain.Name
            Case Is = ShtPlanData.Name
            Case Is = ShtTaskView.Name
            Case Is = ShtDepLog.Name
            Case Else
                WSheet.Delete
        End Select
    Next
    
    Application.DisplayAlerts = True
Exit Sub
ErrorHandler:

    Application.DisplayAlerts = True
End Sub

' ===============================================================
' DisplayTasks
' completes reports for each project
' ---------------------------------------------------------------
Public Sub DisplayTasks(AryTasks() As String)
    Dim Task As Integer
    Dim DataItem As enDataCols
    Dim ProjName As String
    Dim AryTask(1 To 12) As Variant
    Dim SheetName As String
    Dim i As Integer
    
    For Task = LBound(AryTasks, 1) To UBound(AryTasks, 1)
        ProjName = AryTasks(Task, enProject)
        
        For DataItem = LBound(AryTasks, 2) To UBound(AryTasks, 2)
            AryTask(DataItem) = AryTasks(Task, DataItem)
        Next
        WriteTask AryTask, ProjName
    Next Task
End Sub

' ===============================================================
' WriteTask
' Writes a task to the specified sheet
' ---------------------------------------------------------------
Public Sub WriteTask(AryTask() As Variant, ProjName As String)
    Dim WSheet As Worksheet
    Dim x As Integer
    Dim y As Integer
    
    If ProjName = "" Then Exit Sub
    
    Set WSheet = Worksheets(ProjName)
    
    x = Application.WorksheetFunction.CountA(WSheet.Range("B:B")) + 1
    
    For y = LBound(AryTask) To UBound(AryTask)
        WSheet.Range("A1").Offset(x, y) = AryTask(y)
    Next y
        
    Set WSheet = Nothing
End Sub

' ===============================================================
' DataImport
' Imports data into tab
' ---------------------------------------------------------------
Public Sub DataImport()
    Dim ObjMSProject As Object
    Dim Fldr As FileDialog
    Dim FilePath As String
    Dim Tsk As Task
    Dim i As Integer
    Dim ProjName As String
    Dim AryTasks() As Variant
    
    On Error GoTo ErrorHandler:
    
    Set ObjMSProject = CreateObject("MSProject.Application")
    
    Set Fldr = Application.FileDialog(msoFileDialogFilePicker)
    With Fldr
        .Title = "Select a File"
        .Filters.Clear
        .Filters.Add "Microft Project Files", "*.MPP", 1
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .InitialFileName = Application.DefaultFilePath
        
        If .Show <> -1 Then Exit Sub
        
        ShtMain.Unprotect
        [mpp_filepath] = .SelectedItems(1)
        ShtMain.Protect
        
    End With
       
    ModLibrary.PerfSettingsOn
    
    DeleteSheets
    ShtExceptRep.ClearData
    ShtPlanData.ClearTasks
    
    With ShtPlanData
        .Visible = True
        .UsedRange.ClearContents
    End With
    
    With ShtMain
        .Unprotect
        .Range("NO_PROJS") = 0
        .Range("U:U").ClearContents
        .Protect
    End With
    
    If ObjMSProject Is Nothing Then
      MsgBox "Project is not installed"
      Exit Sub
    End If
    
    With ObjMSProject
        .Visible = False
        .DisplayAlerts = False
        .FileOpen Name:=[mpp_filepath], ReadOnly:=True
        .OptionsViewEx DisplaySummaryTasks:=True
        .OutlineShowAllTasks
        .FilterApply Name:="All Tasks"
        .AutoFilter
        .AutoFilter
        .Application.CalculateProject
        
    End With
    
    ReDim AryTasks(1 To ObjMSProject.ActiveProject.Tasks.Count, 1 To 12)
    
    i = 1
    For Each Tsk In ObjMSProject.ActiveProject.Tasks
        AryTasks(i, enRef) = tsk.UniqueID
        AryTasks(i, enlevel) = Tsk.Number1
        AryTasks(i, enMileName) = Tsk.Name
        AryTasks(i, enBaseFinish) = Format(Tsk.BaselineFinish, "dd mmm yy")
        AryTasks(i, enForeFinish) = Format(Tsk.Finish, "dd mmm yy")
        AryTasks(i, enDTI) = Tsk.Number13
        AryTasks(i, enRAG) = Tsk.Text22
        AryTasks(i, enIssue) = Tsk.Text14
        AryTasks(i, enImpact) = Tsk.Text15
        AryTasks(i, enAction) = Tsk.Text16
        AryTasks(i, enProject) = Tsk.Text8
        ProjName = AryTasks(i, enProject)
        
        i = i + 1
    Next Tsk
        
    ShtPlanData.DisplayTasks AryTasks
        
    ModLibrary.PerfSettingsOff
    
    MsgBox "Import Complete", vbOKOnly + vbInformation
    
    ObjMSProject.FileClose (False)
    Set ObjMSProject = Nothing
Exit Sub

ErrorHandler:
    Debug.Print Err.Number & " - " & Err.Description
    ModLibrary.PerfSettingsOff
    ObjMSProject.FileClose (False)
    Set ObjMSProject = Nothing
End Sub

' ===============================================================
' ExceptionReport
' creates exception report from MPP
' ---------------------------------------------------------------
Public Sub ExceptionReport()
    Dim ObjMSProject As Object
    Dim Fldr As FileDialog
    Dim FilePath As String
    Dim Tsk As Task
    Dim AryTasks(1 To 12) As Variant
    Dim MileName As String
    Dim BLFinish As Date
    Dim FCFinish As Date
    Dim ActFinish As Date
    Dim LookAhead As Integer
    Dim TaskComplete As Boolean
    Dim PLevel As Integer
    Dim DTI As Double
    Dim LocalRAG As String
    Dim CalcRAG As String
    Dim PCComplete As Integer
    
    Dim i As Integer
    
    On Error GoTo ErrorHandler:
    
    Set ObjMSProject = CreateObject("MSProject.Application")
       
    ModLibrary.PerfSettingsOn
    
    With ShtExceptRep
        .Visible = True
        .ClearData
    End With
        
    If ObjMSProject Is Nothing Then
      MsgBox "Project is not installed"
      Exit Sub
    End If
    
    With ObjMSProject
        .Visible = False
        .DisplayAlerts = False
        .FileOpen Name:=[mpp_filepath], ReadOnly:=True
        .OptionsViewEx DisplaySummaryTasks:=True
        .OutlineShowAllTasks
        .FilterApply Name:="All Tasks"
        .AutoFilter
        .AutoFilter
        .Application.CalculateProject
        
    End With
    
    LookAhead = ShtMain.Range("LA_PERIOD")
    PLevel = ShtMain.Range("LEVEL")
    
    i = 1
    For Each Tsk In ObjMSProject.ActiveProject.Tasks
        If Not tsk Is Nothing And tsk.Summary = False Then
            If tsk.number1 <= PLevel And tsk.BaselineFinish < DateAdd("ww", LookAhead, Now) Then
        
                AryTasks(enRef) = tsk.UniqueID
                AryTasks(enlevel) = tsk.number1
                AryTasks(enMileName) = tsk.Name
                AryTasks(enBaseFinish) = Format(tsk.BaselineFinish, "dd mmm yy")
                AryTasks(enForeFinish) = Format(tsk.finish, "dd mmm yy")
                AryTasks(enDTI) = tsk.Number13
                AryTasks(enRAG) = tsk.Text22
                AryTasks(enLocalRAG) = tsk.Text10
                AryTasks(enIssue) = tsk.Text14
                AryTasks(enImpact) = tsk.Text15
                AryTasks(enAction) = tsk.Text16
                AryTasks(enProject) = tsk.Text8
                
                If tsk.Summary = False Then
                    MileName = tsk.Text1
                    BLFinish = Format(tsk.BaselineFinish, "dd mmm yy")
                    FCFinish = Format(tsk.finish, "dd mmm yy")
                    DTI = tsk.Number13
                    LocalRAG = tsk.Text10
                    CalcRAG = tsk.Text22
                    PCComplete = tsk.PercentComplete
                    
                    With ShtExceptRep
                        If PCComplete Then
                            .EnterData AryTasks, Completed
                        
                        ElseIf LocalRAG = "AMBER" Then
                            .EnterData AryTasks, Amber
                        
                        ElseIf LocalRAG = "RED" And BLFinish < Now Then
                            .EnterData AryTasks, MissedRed
                        
                        ElseIf LocalRAG = "RED" And BLFinish >= Now Then
                            .EnterData AryTasks, FutureRed
                        End If
                    End With
                End If
                i = i + 1
                Debug.Print i
            End If
        End If
    Next Tsk
    ModLibrary.PerfSettingsOff
    
    MsgBox "Exception Report Created", vbOKOnly + vbInformation
    
    ObjMSProject.FileClose (False)
    Set ObjMSProject = Nothing
Exit Sub

ErrorHandler:
    Debug.Print Err.Number & " - " & Err.Description
    Stop
    Resume Next
    ModLibrary.PerfSettingsOff
    ObjMSProject.FileClose (False)
    Set ObjMSProject = Nothing
End Sub


