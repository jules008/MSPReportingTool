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
' Date - 26 May 20
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
            Me.Range("A2").Offset(x, y - 1) = AryTasks(x, y)
        
            With Me.Range("A2").Offset(x, y - 1).Borders
                .Weight = 2
                .ColorIndex = 1
                .LineStyle = xlContinuous
            End With
       
       Next y
    Next x
    
    Range("A1") = "Dependency Log"
    Range("A2") = "Ref"
    Range("B2") = "Project"
    Range("C2") = "Milestone Name"
    Range("D2") = "Level"
    Range("E2") = "Beneficiary"
    Range("F2") = "Donor"
    Range("G2") = "Baseline Finish"
    Range("H2") = "Forecast Finish"
    Range("I2") = "RAG"
    Range("J2") = "Last RAG"
    Range("K2") = "Issue"
    Range("L2") = "Impact"
    Range("M2") = "Action"
    
End Sub

' ===============================================================
' ClearTasks
' Clears page
' ---------------------------------------------------------------
Public Sub ClearTasks()
    With Range("A3:N10000")
        .ClearContents
        .Borders.LineStyle = xlNone
    End With
End Sub


