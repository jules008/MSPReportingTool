VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtExceptRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtEceptRep
' Displays the exception report
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 3 May 20
'===============================================================
Option Explicit


' ===============================================================
' ClearData
' Clears all data from sheet apart from headings
' ---------------------------------------------------------------
Public Sub ClearData()
    Dim Rw As Row
    Dim NoRows As Integer
    Dim i As Integer
    Dim x As Integer
    
    NoRows = Me.UsedRange.Rows.Count
    
    x = 3
    For i = 1 To NoRows
        If InStr(1, Cells(x, 1), "Milestone", vbTextCompare) Then
            Me.Rows(x + 1).ClearContents
            x = x + 2
        Else
            Me.Rows(x).Delete
        End If
    Next
End Sub

' ===============================================================
' EnterData
' Enters milestones into report
' ---------------------------------------------------------------
Public Sub EnterData(AryInput() As Variant, Section As enExcepRep)
    Dim RngResult As Range
    Dim RowNo As Integer
    Dim RngInsert As Range
    
    Select Case Section
        Case Completed
            Set RngResult = Range("A:A").Find("Completed Milestones")
        Case Amber
            Set RngResult = Range("A:A").Find("Amber Milestones")
        Case FutureRed
            Set RngResult = Range("A:A").Find("Red Milestones in Future")
        Case MissedRed
            Set RngResult = Range("A:A").Find("Missed Red Milestones")
    End Select
    
    If RngResult Is Nothing Then Exit Sub
    
    RowNo = RngResult.Row
    
    Set RngResult = RngResult.Resize(1, 12)
        
    With RngResult
        .Offset(1, 0).Insert Shift:=xlDown
        .Offset(1, 0).ClearFormats
        .Offset(1, 0) = AryInput
    End With
    
    RngResult.Columns.AutoFit
    
    Set RngInsert = Nothing
    Set RngResult = Nothing
End Sub

' ===============================================================
' EnterAmberMLs
' Enters amber milestones into report
' ---------------------------------------------------------------
Public Sub EnterAmberMLs(AryInput() As Variant)
    Dim RngResult As Range
    Dim RowNo As Integer
    Dim RngInsert As Range
    
    Set RngResult = Range("A:A").Find("Completed Milestones")
    
    If RngResult Is Nothing Then Exit Sub
    
    RowNo = RngResult.Row
    
    Set RngResult = RngResult.Resize(1, 12)
        
    With RngResult
        .Offset(2, 0).Insert Shift:=xlDown
        
        RngResult.Offset(1, 0) = AryInput
    End With
    
    RngResult.Columns.AutoFit
    
    Set RngInsert = Nothing
    Set RngResult = Nothing

End Sub

' ===============================================================
' EnterRedFutMls
' Enters future red milestones into report
' ---------------------------------------------------------------
Public Sub EnterRedFutMls(AryInput() As Variant)

End Sub

' ===============================================================
' EnterRedMissMls
' Enters missed red milestones into report
' ---------------------------------------------------------------
Public Sub EnterRedMissMls(AryInput() As Variant)

End Sub

