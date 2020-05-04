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
' Date - 03 May 20
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
            Me.Range("B2").Offset(x, y) = AryTasks(x, y)
        Next y
    Next x
    
    Range("C2") = "Ref"
    Range("D2") = "Level"
    Range("E2") = "Milestone Name"
    Range("F2") = "Baseline Finish"
    Range("G2") = "Forecast Finish"
    Range("H2") = "DTI"
    Range("I2") = "Last RAG"
    Range("J2") = "RAG"
    Range("K2") = "Issue"
    Range("L2") = "Impact"
    Range("M2") = "Action"
    
    Range("A:M").Columns.AutoFit
End Sub

' ===============================================================
' ClearTasks
' Clears page
' ---------------------------------------------------------------
Public Sub ClearTasks()
    Me.UsedRange.ClearContents
End Sub

