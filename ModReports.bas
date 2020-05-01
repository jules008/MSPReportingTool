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
' Date - 1 May 20
'===============================================================
Option Explicit

' ===============================================================
' SelectMPP
' Selects MS Project File for importing
' ---------------------------------------------------------------
Public Sub SelectMPP()
    Dim Fldr As FileDialog
    Dim FilePath As String
    
    Set Fldr = Application.FileDialog(msoFileDialogFilePicker)
    With Fldr
        .Title = "Select a File"
        .Filters.Clear
        .Filters.Add "Microft Project Files", "*.MPP", 1
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .InitialFileName = Application.DefaultFilePath
        
        If .Show <> -1 Then Exit Sub
        [MPP_FILEPATH] = .SelectedItems(1)
    End With
End Sub

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
        .FileOpen Name:=ProjectPath, ReadOnly:=True
        .OptionsViewEx DisplaySummaryTasks:=True
        .OutlineShowAllTasks
        .FilterApply Name:="All Tasks"
        .AutoFilter
        .AutoFilter
        .Application.CalculateProject
        
        ProgName = .ActiveProject.BuiltinDocumentProperties("Subject")
        ProjName = .ActiveProject.BuiltinDocumentProperties("Company")
    End With
    
    LookAhead = ShtMain.Range("LA_PERIOD")
    PLevel = ShtMain.Range("LEVEL")
    
    'cycle through tasks in plan and add to tasks array
    ReDim AryTasks(1 To ObjMSProject.ActiveProject.Tasks.Count, 1 To 12)
    
    i = 1
    For Each Tsk In ObjMSProject.ActiveProject.Tasks
        If Not Tsk Is Nothing And Tsk.Summary = False Then
            If Tsk.Number1 <= PLevel And Tsk.BaselineFinish < DateAdd("ww", LookAhead, Now) Then
                AryTasks(i, enRef) = Tsk.Text1
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
                
                AddProjSheets AryTasks(i, enProject)
                i = i + 1
            End If
        End If
        
        If Not Tsk Is Nothing And Tsk.Summary = True Then
            AryTasks(i, enMileName) = Tsk.Name
            AryTasks(i, enProject) = "All"
            i = i + 1
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
' AddProjSheets
' Checks to see whether a project sheet exists, if not it creates
' one
' ---------------------------------------------------------------
Public Sub AddProjSheets(ProjName As String)
    Dim Wkst As Worksheet
    Dim SheetExists As Boolean
    
    For Each Wkst In Worksheets
        If Wkst.Name = ProjName Then
            SheetExists = True
            Exit For
        End If
    Next
    
    If Not SheetExists Then
        Worksheets("Look Ahead Report").Copy After:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = ProjName
        With Worksheets(ProjName)
            .Range("C2") = "Ref"
            .Range("D2") = "Level"
            .Range("E2") = "Milestone Name"
            .Range("F2") = "Baseline Finish"
            .Range("G2") = "Forecast Finish"
            .Range("H2") = "DTI"
            .Range("I2") = "Last RAG"
            .Range("J2") = "RAG"
            .Range("K2") = "Issue"
            .Range("L2") = "Impact"
            .Range("M2") = "Action"
                
            .Range("A:M").Columns.AutoFit
            .Columns("N").Hidden = True
            
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
    
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    
    For i = Worksheets.Count To 3 Step -1
        Worksheets(i).Activate
        Worksheets(i).Delete
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
            
        If ProjName = "All" Then
            For i = 3 To Worksheets.Count
                WriteTask AryTask, Worksheets(i).Name
            Next
        Else
            WriteTask AryTask, ProjName
        End If
    Next Task
End Sub

' ===============================================================
' WriteTask
' Writes a task to the specified sheet
' ---------------------------------------------------------------
Public Sub WriteTask(AryTask() As Variant, ProjName As String)
    Dim Wsheet As Worksheet
    Static PrevTaskSummary As Boolean
    Dim x As Integer
    Dim y As Integer
    
    If ProjName = "" Then Exit Sub
    
    Set Wsheet = Worksheets(ProjName)
    
    x = Application.WorksheetFunction.CountA(Wsheet.Range("E:E"))
    
    If Wsheet.Range("B2").Offset(x - 1, enlevel) = "" And x > 1 Then
        PrevTaskSummary = True
    Else
        PrevTaskSummary = False
    End If
    
    If PrevTaskSummary And AryTask(enlevel) = "" Then
        Wsheet.Range("B2").Offset(x, enlevel) = 0
        Wsheet.Range("B2").Offset(x, enMileName) = "No Tasks"
        x = x + 1
    End If
    
    For y = LBound(AryTask) To UBound(AryTask)
        Wsheet.Range("B2").Offset(x, y) = AryTask(y)
    Next y
        
    Wsheet.Range("A:M").Columns.AutoFit
    Set Wsheet = Nothing
End Sub
