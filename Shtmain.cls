VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' ShtMain
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 1 May 20
'===============================================================
Option Explicit

' ===============================================================
' BtnImportClick
' ---------------------------------------------------------------
Public Sub BtnImportClick()
    SelectMPP
End Sub

' ===============================================================
' BtnPPTExpClick
' ---------------------------------------------------------------
Public Sub BtnPPTExpClick()
    ActionButtonClick Me.Shapes("BtnPPTExp")
    CreatePowerPoint.CreatePowerPoint
End Sub

' ===============================================================
' BtnLARGenClick
' ---------------------------------------------------------------
Public Sub BtnLARGenClick()
    ActionButtonClick Me.Shapes("BtnLARGen")
    LookAheadRep [MPP_FILEPATH]
End Sub

' ===============================================================
' ActionButtonClick
' ---------------------------------------------------------------
Public Sub ActionButtonClick(Button As Shape)
    Dim DTime As Double
    
    On Error GoTo ErrorHandler
    
    If [MPP_FILEPATH] = "" Then
        MsgBox "Please import a Project File", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    FormatButton Button, True
    
    DTime = Time
    Do While Time < DTime + 1 / 24 / 60 / 60 / 2
        DoEvents
    Loop
    
    FormatButton Button, False
        
    
Exit Sub
ErrorHandler:
    MsgBox "An error has occured, unable to open the Project File", vbOKCancel + vbCritical
    
End Sub

' ===============================================================
' BtnLvl1Click
' ---------------------------------------------------------------
Public Sub BtnLvl1Click()
    Me.Unprotect
    [Level] = 1
    FormatButton Me.Shapes("BtnLvl1"), True
    FormatButton Me.Shapes("BtnLvl2"), False
    FormatButton Me.Shapes("BtnLvl3"), False
    FormatButton Me.Shapes("BtnLvl4"), False
    Me.Protect
End Sub

' ===============================================================
' BtnLvl2Click
' ---------------------------------------------------------------
Public Sub BtnLvl2Click()
    Me.Unprotect
    [Level] = 2
    FormatButton Me.Shapes("BtnLvl1"), False
    FormatButton Me.Shapes("BtnLvl2"), True
    FormatButton Me.Shapes("BtnLvl3"), False
    FormatButton Me.Shapes("BtnLvl4"), False
    Me.Protect
End Sub

' ===============================================================
' BtnLvl3Click
' ---------------------------------------------------------------
Public Sub BtnLvl3Click()
    Me.Unprotect
    [Level] = 3
    FormatButton Me.Shapes("BtnLvl1"), False
    FormatButton Me.Shapes("BtnLvl2"), False
    FormatButton Me.Shapes("BtnLvl3"), True
    FormatButton Me.Shapes("BtnLvl4"), False
    Me.Protect
End Sub

' ===============================================================
' BtnLvl4Click
' ---------------------------------------------------------------
Public Sub BtnLvl4Click()
    Me.Unprotect
    
    [Level] = 4
    FormatButton Me.Shapes("BtnLvl1"), False
    FormatButton Me.Shapes("BtnLvl2"), False
    FormatButton Me.Shapes("BtnLvl3"), False
    FormatButton Me.Shapes("BtnLvl4"), True
    
    Me.Protect
End Sub

' ===============================================================
' BtnLAR4wClick
' ---------------------------------------------------------------
Public Sub BtnLAR4wClick()
    Me.Unprotect
    [LA_PERIOD] = 4
    FormatButton Me.Shapes("BtnLAR12w"), False
    FormatButton Me.Shapes("BtnLAR4w"), True
    Me.Protect
End Sub

' ===============================================================
' BtnLAR12wClick
' ---------------------------------------------------------------
Public Sub BtnLAR12wClick()
    Me.Unprotect
    [LA_PERIOD] = 12
    FormatButton Me.Shapes("BtnLAR12w"), True
    FormatButton Me.Shapes("BtnLAR4w"), False
    Me.Protect
End Sub

' ===============================================================
' FormatButton
' Formats the control to be selected or not
' ---------------------------------------------------------------
Private Sub FormatButton(Btn As Shape, OnState As Boolean)
    Me.Unprotect
    With Btn
        If OnState Then
            With .Fill
                .TwoColorGradient msoGradientHorizontal, 1
                .ForeColor.RGB = COLOUR_2
                .BackColor.RGB = COLOUR_2
            End With
            
            With .Line
                .ForeColor.RGB = COLOUR_4
                .Weight = 0
                If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
            End With
            
            With .TextFrame
                .Characters.Font.Bold = False
                .Characters.Font.Name = "Abadi"
                .Characters.Font.Color = COLOUR_7
                .Characters.Font.Size = 10
                .HorizontalAlignment = xlHAlignCenter
            End With
            
            With .Shadow
                .Visible = msoTrue
                .Type = msoShadow30
            End With
            
        Else
            With .Fill
                .TwoColorGradient msoGradientHorizontal, 1
                .ForeColor.RGB = COLOUR_8
                .BackColor.RGB = COLOUR_8
            End With
            
            With .Line
                .ForeColor.RGB = COLOUR_2
                .Weight = 0
                If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
            End With
            
            With .TextFrame
                .Characters.Font.Bold = False
                .Characters.Font.Name = "Abadi"
                .Characters.Font.Color = COLOUR_7
                .Characters.Font.Size = 10
                .HorizontalAlignment = xlHAlignCenter
            End With
            
            With .Shadow
                .Visible = msoTrue
                .Type = msoShadow30
               .Visible = msoFalse
            End With
            
        
        End If
    End With
    Me.Protect
End Sub


