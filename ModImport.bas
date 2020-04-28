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
    ReDim AryTasks(ObjMSProject.ActiveProject.Tasks.Count, 7)
    
    For Each Tsk In ObjMSProject.ActiveProject.Tasks
        If Not Tsk Is Nothing And Tsk.Summary = False Then
            AryTasks(Tsk.ID, 1) = Tsk.ID
            AryTasks(Tsk.ID, 2) = Tsk.Name
            AryTasks(Tsk.ID, 3) = DurationFormat(Tsk.Duration, pjDays)
            AryTasks(Tsk.ID, 4) = Format(Tsk.BaselineFinish, "dd mmm yy")
            AryTasks(Tsk.ID, 5) = Format(Tsk.ActualFinish, "dd mmm yy")
            AryTasks(Tsk.ID, 6) = Tsk.Summary
            AryTasks(Tsk.ID, 7) = Format(Tsk.ScheduledFinish, "dd mmm yy")
        End If
        
        If Not Tsk Is Nothing And Tsk.Summary = True Then
            AryTasks(Tsk.ID, 2) = Tsk.Name
            AryTasks(Tsk.ID, 6) = Tsk.Summary
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
    ModLibrary.PerfSettingsOff
    ObjMSProject.FileClose (False)
    Set ObjMSProject = Nothing
End Sub
