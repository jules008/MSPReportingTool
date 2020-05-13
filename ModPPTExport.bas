Attribute VB_Name = "ModPPTExport"
'===============================================================
' Module ModPPTExport
' Exports data to presentation
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

Private PowerPointApp As PowerPoint.Application
Private MyPPT As PowerPoint.Presentation

' ===============================================================
' PPTExport
' master routine to manage PPT Export
' ---------------------------------------------------------------
Public Sub PPTExport()
    
    On Error Resume Next
    
    Set PowerPointApp = GetObject(Class:="PowerPoint.Application")
    
    Err.Clear

    If PowerPointApp Is Nothing Then Set PowerPointApp = CreateObject(Class:="PowerPoint.Application")
    
    If Err.Number = 429 Then
        MsgBox "PowerPoint could not be found, aborting."
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler

    PowerPointApp.Presentations.Add
    Set MyPPT = PowerPointApp.ActivePresentation
    
    PerfSettingsOn
    
    GetSlideRange
        
    PerfSettingsOff
    Set PowerPointApp = Nothing
    Set MyPPT = Nothing
    
Exit Sub

ErrorHandler:


    PerfSettingsOff
    Set PowerPointApp = Nothing
    Set MyPPT = Nothing

End Sub

' ===============================================================
' GetSlideRange
' Takes each range of data for export to slide
' ---------------------------------------------------------------
Public Sub GetSlideRange()
    Dim RepSheet As Worksheet
    Dim RngReport As Range
    Dim Title As String
    
    On Error GoTo ErrorHandler
    
    For Each RepSheet In Worksheets
        Select Case RepSheet.Name
            
            Case ShtMain.Name, ShtTaskView.Name, ShtPlanData.Name
            
            Case Else
        
                Set RngReport = RepSheet.UsedRange
                RngReport.Copy
                Title = RepSheet.Range("A1")
                CreatePPTSlide RngReport, Title
        End Select
    
    Next
    Set RepSheet = Nothing
    Set RngReport = Nothing
Exit Sub

ErrorHandler:
    
    Set RepSheet = Nothing
    Set RngReport = Nothing
End Sub

' ===============================================================
' CreatePPTSlide
' Takes range and pastes into new slide of powerpoint
' ---------------------------------------------------------------
Sub CreatePPTSlide(RngPaste As Range, Title As String)
    Dim mySlide As PowerPoint.Slide
    Dim myShape As Object
  
    Set mySlide = MyPPT.Slides.Add(1, 11) '11 = ppLayoutTitleOnly
    
    mySlide.Shapes.Title.TextFrame.TextRange = Title
    mySlide.Shapes.PasteSpecial DataType:=2  '2 = ppPasteEnhancedMetafile
    Set myShape = mySlide.Shapes(mySlide.Shapes.Count)
  
    myShape.Left = 40
    myShape.Top = 100
    myShape.Width = 900

    PowerPointApp.Visible = True
    PowerPointApp.Activate

    Application.CutCopyMode = False
  
End Sub



