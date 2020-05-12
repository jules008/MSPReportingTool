VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Public Sub ClearData()
    ShtPlanData.ClearTasks
    ModReports.DeleteSheets
    ShtExceptRep.ClearData
    ShtMain.Activate
    ShtDepLog.ClearTasks
End Sub
