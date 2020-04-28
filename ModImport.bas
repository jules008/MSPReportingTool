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
        FilePath = .SelectedItems(1)
    End With
    ImportData FilePath
End Sub

' ===============================================================
' ImportData
' Opens project file from passed filepath, extracts data and sends
' to display procedure
' ---------------------------------------------------------------
Private Sub ImportData(ProjectPath As String)
    Dim ObjMSProject As Object
    Dim ProgName As String
    Dim ProjName As String
    Dim AryTasks() As String
    Dim Tsk As Task
    
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
    
    'cycle through tasks in plan and add to tasks array
    ReDim AryTasks(ObjMSProject.ActiveProject.Tasks.Count, 11)
    
    For Each Tsk In ObjMSProject.ActiveProject.Tasks
        If Not Tsk Is Nothing And Tsk.Summary = False Then
            AryTasks(Tsk.ID, enRef) = Tsk.Text1
            AryTasks(Tsk.ID, enLevel) = Tsk.Number1
            AryTasks(Tsk.ID, enMileName) = Tsk.Name
            AryTasks(Tsk.ID, enBaseFinish) = Format(Tsk.BaselineFinish, "dd mmm yy")
            AryTasks(Tsk.ID, enForeFinish) = Format(Tsk.Finish, "dd mmm yy")
            AryTasks(Tsk.ID, enDTI) = "DTI"
            AryTasks(Tsk.ID, enLastRAG) = Tsk.Text21
            AryTasks(Tsk.ID, enRAG) = Tsk.Text22
            AryTasks(Tsk.ID, enIssue) = Tsk.Text14
            AryTasks(Tsk.ID, enImpact) = Tsk.Text15
            AryTasks(Tsk.ID, enAction) = Tsk.Text16
        End If
        
        If Not Tsk Is Nothing And Tsk.Summary = True Then
            AryTasks(Tsk.ID, enMileName) = Tsk.Name
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
