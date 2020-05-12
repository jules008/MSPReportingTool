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
' Date - 12 May 20
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
    
    Set PowerPointApp = New PowerPoint.Application
        
    On Error GoTo ErrorHandler

    PowerPointApp.Presentations.Add
    Set MyPPT = PowerPointApp.ActivePresentation
    
    PerfSettingsOn
    
    GetSlideRange
    
    
    MsgBox "Powerpoint Created"
    
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
    
    For Each RepSheet In Workbooks
        Select Case RepSheet.Name
            
            Case "ShtMain", "ShtTaskView", "ShtDepLog", "ShtExceptRep"
            
                
        Set RngReport = RepSheet.UsedRange
        
        Title = RepSheet.Range("A1")
        
        CreatePPTSlide RngReport, Title
    
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


'Add a slide to the Presentation
  Set mySlide = MyPPT.Slides.Add(1, 11) '11 = ppLayoutTitleOnly

'Copy Excel Range
  rng.Copy

'Paste to PowerPoint and position
  mySlide.Shapes.PasteSpecial DataType:=2  '2 = ppPasteEnhancedMetafile
  Set myShape = mySlide.Shapes(mySlide.Shapes.Count)
  
    'Set position:
      myShape.Left = 66
      myShape.Top = 152

'Make PowerPoint Visible and Active
  PowerPointApp.Visible = True
  PowerPointApp.Activate

'Clear The Clipboard
  Application.CutCopyMode = False
  
End Sub



