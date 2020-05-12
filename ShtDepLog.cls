VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtDepLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtDepLog
' Displays tasks from Plan
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 12 May 20
'===============================================================
Option Explicit

' ===============================================================
' DisplayTasks
' Displays tasks sent from Project Plan
' ---------------------------------------------------------------
Public Sub DisplayTasks(AryTasks() As Variant)
    Dim x As Integer
    Dim y As Integer
    
    Range("A1:AA10000").ClearContents
    
    For x = LBound(AryTasks, 1) To UBound(AryTasks, 1)
        For y = LBound(AryTasks, 2) To UBound(AryTasks, 2)
            Me.Range("A2").Offset(x, y) = AryTasks(x, y)
        Next y
    Next x
    
    Range("B2") = "Ref"
    Range("C2") = "Milestone Name"
    Range("D2") = "Level"
    Range("D2") = "Dep In"
    Range("D2") = "Dep Out"
    Range("D2") = "Beneficiary"
    Range("D2") = "Donor"
    Range("E2") = "Baseline Finish"
    Range("F2") = "Forecast Finish"
    Range("G2") = "RAG"
    Range("H2") = "Last RAG"
    Range("I2") = "Issue"
    Range("J2") = "Impact"
    Range("K2") = "Action"
    Range("L2") = "Project"
    
    Range("A:M").Columns.AutoFit
End Sub

' ===============================================================
' ClearTasks
' Clears page
' ---------------------------------------------------------------
Public Sub ClearTasks()
    Me.UsedRange.ClearContents
End Sub


