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
' Date - 26 May 20
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
        
    ReverseSlideOrder MyPPT
    
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
    Dim RngPrint As Range
    Dim RngReport As Range
    Dim RngHeading As Range
    Dim Title As String
    Dim x As Integer
    Dim NoCols As Integer
    Dim ShtTemp As Worksheet
    
    On Error GoTo ErrorHandler
    
    
    For Each RepSheet In Worksheets
    
        Select Case RepSheet.Name
                  
            Case ShtMain.Name, ShtTaskView.Name, ShtPlanData.Name
            
            Case Else
    
                RepSheet.Copy After:=Worksheets(Worksheets.Count)
                ActiveSheet.Name = "TempSht"
                
                Set RngReport = ActiveSheet.UsedRange
                
                NoCols = RngReport.Columns.Count
                
                x = 1
                Do While Application.WorksheetFunction.CountA(RngReport.Range("A:A")) > 2
                    With RngReport
                        Set RngPrint = .Range(.Cells(2, 1), .Cells(16, NoCols))
                    End With
                    Debug.Print RngPrint.Address
                    
                    Title = RepSheet.Range("A1")
                    RngPrint.Copy
                    
                    If Application.WorksheetFunction.CountA(RngPrint) > 0 Then
                        CreatePPTSlide RngHeading, Title
                    End If
                    
                    With RngReport
                        .Range(.Cells(3, 1), .Cells(3 + 13, NoCols)).Delete Shift:=xlShiftUp
                    End With
                    
                    
                    x = x + 15
                    
                Loop
                Application.DisplayAlerts = False
                Worksheets("TempSht").Delete
                Application.DisplayAlerts = True
        End Select
        
                                
    Next
    Set ShtTemp = Nothing
    Set RepSheet = Nothing
    Set RngReport = Nothing
    Set RngPrint = Nothing
Exit Sub

ErrorHandler:
    
    Application.DisplayAlerts = True
    
    Set RepSheet = Nothing
    Set RngReport = Nothing
    Set RngPrint = Nothing
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

' ===============================================================
' ReverseSlideOrder
' Reverses the order of the powerpoint slides
' ---------------------------------------------------------------
Sub ReverseSlideOrder(MyPPT As PowerPoint.Presentation)
   Dim NoSlides As Long
   Dim x As Long

   NoSlides = MyPPT.Slides.Count

      For x = 1 To NoSlides - 1

         MyPPT.Slides(NoSlides).Cut
         MyPPT.Slides.Paste x
      Next x

End Sub


