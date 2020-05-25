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
' Date - 25 May 20
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
            
            With Me.Range("A2").Offset(x, y).Borders
                .Weight = 2
                .ColorIndex = 1
                .LineStyle = xlContinuous
            End With
       Next y
    Next x
    
    Range("A1") = "Dependency Log"
    Range("B2") = "Project"
    Range("C2") = "Ref"
    Range("D2") = "Milestone Name"
    Range("E2") = "Level"
    Range("F2") = "Beneficiary"
    Range("G2") = "Donor"
    Range("H2") = "Baseline Finish"
    Range("I2") = "Forecast Finish"
    Range("J2") = "RAG"
    Range("K2") = "Last RAG"
    Range("L2") = "Issue"
    Range("M2") = "Impact"
    Range("N2") = "Action"
    
End Sub

' ===============================================================
' ClearTasks
' Clears page
' ---------------------------------------------------------------
Public Sub ClearTasks()
    With Range("B3:N10000")
        .ClearContents
        .Borders.LineStyle = xlNone
    End With
End Sub


