VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtTaskView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtTaskView
' Displays tasks from Plan
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Feb 20
'===============================================================
Option Explicit

' ===============================================================
' DisplayTasks
' Displays tasks sent from Project Plan
' ---------------------------------------------------------------
Public Sub DisplayTasks(AryTasks() As String)
    Dim x As Integer
    Dim y As Integer
    
    Range("A1:AA1000").ClearContents
    
    For x = LBound(AryTasks, 1) To UBound(AryTasks, 1)
        For y = LBound(AryTasks, 2) To UBound(AryTasks, 2)
            Me.Range("B2").Offset(x, y) = AryTasks(x, y)
        Next y
    Next x
    
    Range("C2") = "ID"
    Range("D2") = "Name"
    Range("E2") = "Duration"
    Range("F2") = "Baseline End"
    Range("G2") = "Actual End"
    Range("H2") = "Summary"
    Range("I2") = "Scheduled End"
    Range("A:I").Columns.AutoFit
    Columns("H").Hidden = True
End Sub
