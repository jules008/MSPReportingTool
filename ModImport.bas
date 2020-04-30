Attribute VB_Name = "ModImport"
'===============================================================
' Module ModImport
' Imports data from MS Project File
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 28 Apr 20
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
    
    LookAhead = Shtmain.Range("LA_PERIOD")
    PLevel = Shtmain.Range("LEVEL")
    
    'cycle through tasks in plan and add to tasks array
    ReDim AryTasks(ObjMSProject.ActiveProject.Tasks.Count, 11)
    
    i = 1
    For Each Tsk In ObjMSProject.ActiveProject.Tasks
        If Not Tsk Is Nothing And Tsk.Summary = False Then
            If Tsk.Number1 <= PLevel and tsk.baselinefinish < dateadd(  Then
                AryTasks(i, enRef) = Tsk.Text1
                AryTasks(i, enLevel) = Tsk.Number1
                AryTasks(i, enMileName) = Tsk.Name
                AryTasks(i, enBaseFinish) = Format(Tsk.BaselineFinish, "dd mmm yy")
                AryTasks(i, enForeFinish) = Format(Tsk.Finish, "dd mmm yy")
                AryTasks(i, enDTI) = Tsk.Number13
                AryTasks(i, enRAG) = Tsk.Text22
                AryTasks(i, enIssue) = Tsk.Text14
                AryTasks(i, enImpact) = Tsk.Text15
                AryTasks(i, enAction) = Tsk.Text16
                i = i + 1
            End If
        End If
        
        If Not Tsk Is Nothing And Tsk.Summary = True Then
            AryTasks(Tsk.ID, enMileName) = Tsk.Name
            i = i + 1
        End If
    Next Tsk
    
    ShtTaskView.DisplayTasks AryTasks
    
    ModLibrary.PerfSettingsOff
    ShtTaskView.Activate
    
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
