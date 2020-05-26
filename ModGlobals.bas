Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 25 May 20
'===============================================================
Option Explicit

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
Public Const PROJECT_FILE_NAME As String = "MSP Reporting Tool"
Public Const APP_NAME As String = "MSP Reporting Tool"
Public Const EXPORT_FILE_PATH As String = "G:\Development Areas\MSP Reporting Tool\Library\"
Public Const IMPORT_FILE_PATH As String = "G:\MSPReportingTool\"
Public Const VERSION = "V0.0.0"
Public Const VER_DATE = "28 Apr 20"

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ===============================================================
' Colours
' ---------------------------------------------------------------
Public Const COLOUR_1 As Long = 12239288
Public Const COLOUR_2 As Long = 3681498
Public Const COLOUR_3 As Long = 1012796
Public Const COLOUR_4 As Long = 7170677
Public Const COLOUR_5 As Long = 16548170
Public Const COLOUR_6 As Long = 9868187
Public Const COLOUR_7 As Long = 15263976
Public Const COLOUR_8 As Long = 8486538

' ===============================================================
' Enums
' ---------------------------------------------------------------
Enum enDataCols
    enRef = 1
    enLevel
    enMileName
    enBaseFinish
    enForeFinish
    enDTI
    enLocalRAG
    enRAG
    enIssue
    enImpact
    enAction
    enProject
End Enum

Enum enDLDataCols
    enDLRef = 1
    enDLProject
    enDLMileName
    enDLLevel
    enDLBenef
    enDLDonor
    enDLBaseFinish
    enDLForeFinish
    enDLRAG
    enDLLocalRAG
    enDLIssue
    enDLImpact
    enDLAction
End Enum

Enum enExcepRep
    MissedRed = 1
    FutureRed
    Amber
    Completed
End Enum
