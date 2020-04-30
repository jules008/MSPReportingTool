Attribute VB_Name = "ModLibrary"
'===============================================================
' Module ModLibrary
' Library Procedure settings
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 28 Apr 20
'===============================================================
Option Explicit

' ===============================================================
' PerfSettingsOn
' turns off system functions to increase performance
' ---------------------------------------------------------------
Public Sub PerfSettingsOn()
    On Error Resume Next
    
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

' ===============================================================
' PerfSettingsOff
' turns system functions back to normal
' ---------------------------------------------------------------
Public Sub PerfSettingsOff()
    On Error Resume Next
        
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ===============================================================
' ColourConvert
' Converts RGB colour to long
' ---------------------------------------------------------------

Public Sub ColourConvert(R As Integer, A As Integer, G As Integer)
     Dim Colour1 As Long
     Colour1 = RGB(R, A, G)
     
     Debug.Print Colour1

End Sub
