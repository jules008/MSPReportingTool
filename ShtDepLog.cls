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
    
    ClearTasks
    
    For x = LBound(AryTasks, 1) To UBound(AryTasks, 1)
        For y = LBound(AryTasks, 2) To UBound(AryTasks, 2)
            Me.Range("A2").Offset(x, y) = AryTasks(x, y)
        Next y
    Next x
    
    Range("A1") = "Dependency Log"
    Range("B2") = "Ref"
    Range("C2") = "Milestone Name"
    Range("D2") = "Level"
    Range("E2") = "Dep In"
    Range("F2") = "Dep Out"
    Range("G2") = "Beneficiary"
    Range("H2") = "Donor"
    Range("I2") = "Baseline Finish"
    Range("J2") = "Forecast Finish"
    Range("K2") = "RAG"
    Range("L2") = "Last RAG"
    Range("M2") = "Issue"
    Range("N2") = "Impact"
    Range("O2") = "Action"
    Range("P2") = "Project"
    
    Range("A:M").Columns.AutoFit
End Sub

' ===============================================================
' ClearTasks
' Clears page
' ---------------------------------------------------------------
Public Sub ClearTasks()
    Me.UsedRange.ClearContents
End Sub


