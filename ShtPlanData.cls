VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtPlanData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtPlanData
' Displays tasks from Plan
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 11 May 20
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
    Range("C2") = "Level"
    Range("D2") = "Milestone Name"
    Range("E2") = "Baseline Finish"
    Range("F2") = "Forecast Finish"
    Range("G2") = "DTI"
    Range("H2") = "Last RAG"
    Range("I2") = "RAG"
    Range("J2") = "Issue"
    Range("K2") = "Impact"
    Range("L2") = "Action"
    
    Range("A:M").Columns.AutoFit
End Sub

' ===============================================================
' ClearTasks
' Clears page
' ---------------------------------------------------------------
Public Sub ClearTasks()
    Me.UsedRange.ClearContents
End Sub

