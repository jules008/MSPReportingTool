Attribute VB_Name = "CreatePowerPoint"
Sub CreatePowerPoint()

    'First we declare the variables we will be using
    Dim newPowerPoint As PowerPoint.Application
    Dim activeSlide As PowerPoint.Slide
    Dim pptable As PowerPoint.Table
    Dim ppshape As PowerPoint.Shape
    Dim cht As Excel.ChartObject
    Dim xlcopy As Range
    Dim ws(1 To 50) As Worksheet '
    Dim wsname(1 To 50) As String  '
    Dim wsnum As Long
    Dim wsrows(1 To 50) As Long
    Dim wspages(1 To 50) As Long
    Dim wspagestart(1 To 50) As Long
    Dim wspagefinish(1 To 50) As Long
    Dim rowref, nextrow As Long
    Dim wsref, j, maxrows As Long
    Dim startrows, heighttable, maxheight, minheight As Long
    Dim morerows, lessrows As Boolean
    Dim i As Long
         
    ' identify no of workstreams to process
    
    On Error Resume Next
    
    dteNow = Now 'Format(Sheets("Control Panel").Range("F16").Value, "dd mmmm yyyy")
    startrows = 5 'Sheets("Control Panel").Range("G22").Value ' taken from control panel (value to use as the default to start the copy/paste process)
    maxheight = 500 ' when adding new lines this is the max limit to fit on a page
    minheight = 450 'Sheets("Control Panel").Range("G23").Value ' taken from control panel (value to use for slide sizing parameter)
    scalefactor = 1 'Sheets("Control Panel").Range("G24").Value ' taken from control panel (value to use for scaling ppt table
    pasteformat = "Picture" 'Sheets("Control Panel").Range("G25").Value ' taken from control panel (value to use for scaling ppt table
    
    wsnum = ShtMain.Range("no_projs")
    If wsnum = 0 Then
        MsgBox "Please run the Look Ahead Report first to extract the presentation data", vbOKOnly + vbInformation
        Exit Sub
    End If
    ' set the wsname array with workstream names

    For i = 1 To wsnum ' set w/s names and row values

        wsname(i) = ShtMain.Cells(i + 2, 21)
        wsrows(i) = Application.WorksheetFunction.CountA(Worksheets(wsname(i)).Range("E:E"))
        rowref = wsrows(i) + 1 ' this is the start row for pasting additional pages
        j = 0 ' used for page number
        morerows = False
        lessrows = False
    
    ' Look for existing instance and if so use that one
         
        
    If PPTisopen Then
        Set newPowerPoint = GetObject(, "PowerPoint.Application")
    Else
        Set newPowerPoint = New PowerPoint.Application
    End If
    If newPowerPoint.Presentations.Count = 0 Then
        newPowerPoint.Presentations.Add
    End If
     
' Show the PowerPoint
    
    newPowerPoint.Visible = True
    newPowerPoint.Activate
    
    'Loop through each workstream report in the Excel worksheet and paste them into the PowerPoint
    
    Sheets(wsname(i)).Activate
    
    wsref = i
    
    Columns("F:F").EntireColumn.Hidden = True
    Columns("H:I").EntireColumn.Hidden = True

' loop through for each page according to page count
    
Newpage:

j = j + 1 ' increment page number

        'Add a new slide where we will paste the 12WLA Report
        
            newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutTitleOnly
            newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
            Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)
            
            newPowerPoint.ActiveWindow.ViewType = ppViewSlide
                
        'Copy the report and paste it into the PowerPoint as a Metafile Picture
            
' set row start and finish values for each page for copying

Addrows:

If j = 1 Then ' for first page

    wspagestart(j) = 8
    If morerows = False Then
        wspagefinish(j) = wspagestart(j) + startrows + 2
        
        If wspagefinish(j) > wsrows(i) Then ' we can get it all on first page
        
            wspagefinish(j) = wsrows(i)
        
        End If
        
    End If
    
Else ' for pages after the first

    If morerows = False Then
    
        wspagestart(j) = wspagefinish(j - 1) + 1
    
        wspagefinish(j) = wspagestart(j) + startrows - 1
    
        If wspagefinish(j) > wsrows(wsref) Then ' end page

            wspagefinish(j) = wsrows(wsref)
        
        Else ' middle page
    
        End If ' wspagefinish

    End If ' morerows
    
End If ' j = 1
    
        
    If j = 1 Then ' for first page
        
        Range("C" & wspagestart(j) & ":Q" & wspagefinish(j)).Select

        Application.CutCopyMode = False
        Set xlcopy = Selection
        Selection.Copy
        
' new code here..
        
'        newPowerPoint.ActivePresentation.Windows(1).View.PasteSpecial DataType:=ppPasteDefault
If pasteformat = "Table" Then
        newPowerPoint.ActivePresentation.Windows(1).View.PasteSpecial DataType:=ppPasteDefault
        Set pptable = activeSlide.Shapes(2).Table
        Set ppshape = activeSlide.Shapes(2)
Else
        newPowerPoint.ActivePresentation.Windows(1).View.PasteSpecial DataType:=ppPasteEnhancedMetafile
End If
            

GoTo nextstep1

pptable.Columns(1).Width = 15 ' Level
pptable.Columns(2).Width = 50 ' ref
pptable.Columns(3).Width = 100 ' name
pptable.Columns(4).Width = 60 ' BL
pptable.Columns(5).Width = 60 ' FC
pptable.Columns(6).Width = 20 ' impact
pptable.Columns(7).Width = 60 ' LRAG
pptable.Columns(8).Width = 60 ' RAG
pptable.Columns(9).Width = 150 ' issue
pptable.Columns(10).Width = 150 ' impact
pptable.Columns(11).Width = 150 ' action

With activeSlide.Shapes(2)

For R = 1 To pptable.Rows.Count
    
    If R = 1 Then
        pptable.Rows(R).Height = 30
    Else
        pptable.Rows(R).Height = 15
    End If
    pptable.Rows(R).Cells.Borders(ppBorderBottom).Visible = msoTrue
    pptable.Rows(R).Cells.Borders(ppBorderTop).Visible = msoTrue
    pptable.Rows(R).Cells.Borders(ppBorderLeft).Visible = msoTrue
    pptable.Rows(R).Cells.Borders(ppBorderRight).Visible = msoTrue
    
    For c = 1 To pptable.Columns.Count
    
        With pptable.Cell(R, c).Shape.TextFrame
                
                .TextRange.Font.Name = arial
                .TextRange.Font.Size = 8
                
                If (c = 2 Or c = 3 Or c >= 9) And R <> 1 Then
                
                    .MarginLeft = 3
        
                End If

        End With
        
    Next c
            
Next R

End With

nextstep1:

If pasteformat = "Table" Then
 If ppshape.Height <> xlcopy.Height Then
    ppshape.Table.ScaleProportionally scalefactor * (xlcopy.Height / ppshape.Height)
 End If
End If
            With activeSlide.Shapes(2)
                
                .LockAspectRatio = msoTrue
                .Width = 710
                .Top = 40
                .Left = 5
'                .TextEffect.FontName = arial

                heighttable = .Height
            
            End With
        
'        activeSlide.Shapes.PasteSpecial(DataType:=xlBitmap).Select
    
    
    Else ' for second and pages afterwards need to copy header row and specific rows

        Call createheader(wsname(wsref), rowref, nextrow, wsref)
        Rows(wspagestart(j) & ":" & wspagefinish(j)).Select
        Application.CutCopyMode = False
        
        Selection.Copy
            
        Rows(rowref & ":" & rowref).Select
            
        Selection.Insert Shift:=xlDown
            
        Range("C" & rowref - 2 & ":Q" & rowref + wspagefinish(j) - wspagestart(j)).Select
            
        Set xlcopy = Selection
        Selection.Copy
            
' new code here..
        
'        newPowerPoint.ActivePresentation.Windows(1).View.PasteSpecial DataType:=ppPasteDefault
If pasteformat = "Table" Then
        newPowerPoint.ActivePresentation.Windows(1).View.PasteSpecial DataType:=ppPasteDefault
        Set pptable = activeSlide.Shapes(2).Table
        Set ppshape = activeSlide.Shapes(2)
Else
        newPowerPoint.ActivePresentation.Windows(1).View.PasteSpecial DataType:=ppPasteEnhancedMetafile
End If
        
GoTo nextstep2

pptable.Columns(1).Width = 15 ' Level
pptable.Columns(2).Width = 50 ' ref
pptable.Columns(3).Width = 100 ' name
pptable.Columns(4).Width = 60 ' BL
pptable.Columns(5).Width = 60 ' FC
pptable.Columns(6).Width = 20 ' impact
pptable.Columns(7).Width = 60 ' LRAG
pptable.Columns(8).Width = 60 ' RAG
pptable.Columns(9).Width = 150 ' issue
pptable.Columns(10).Width = 150 ' impact
pptable.Columns(11).Width = 150 ' action

With activeSlide.Shapes(2)

For R = 1 To pptable.Rows.Count
    
    If R = 1 Then
        pptable.Rows(R).Height = 30
    Else
        pptable.Rows(R).Height = 15
    End If
    pptable.Rows(R).Cells.Borders(ppBorderBottom).Visible = msoTrue
    pptable.Rows(R).Cells.Borders(ppBorderTop).Visible = msoTrue
    pptable.Rows(R).Cells.Borders(ppBorderLeft).Visible = msoTrue
    pptable.Rows(R).Cells.Borders(ppBorderRight).Visible = msoTrue
    
    For c = 1 To pptable.Columns.Count
    
        With pptable.Cell(R, c).Shape.TextFrame
                
                .TextRange.Font.Name = arial
                .TextRange.Font.Size = 8
                
                If (c = 2 Or c = 3 Or c >= 9) And R <> 1 Then
                
                    .MarginLeft = 3
        
                End If

        End With
        
    Next c
            
Next R

End With

nextstep2:

If pasteformat = "Table" Then
 If ppshape.Height <> xlcopy.Height Then
    ppshape.Table.ScaleProportionally scalefactor * (xlcopy.Height / ppshape.Height)
 End If
End If

        With activeSlide.Shapes(2)
            
            .LockAspectRatio = msoTrue
            .Width = 710
            .Top = 40
            .Left = 5
'           .TextEffect.FontName = arial

            heighttable = .Height

        End With
        
'        activeSlide.Shapes.PasteSpecial(DataType:=xlBitmap).Select
                        
    End If
    
        'Set the title of the slide the same as the title of the chart - add page no if required
        
    If morerows = False Then
    
    activeSlide.Shapes(1).TextFrame.TextRange.Text = Range("C3").Value & " (" & j & ")" & Chr(10) _
        & "Period Ending " & dteNow
            
             With activeSlide.Shapes(1)
            
                .Width = 550
                .Left = 10
                .Top = 5
                .Height = 30
                HorizontalAlignment = xlLeft

            End With
            
        'Now let's change the font size of the callouts box
            activeSlide.Shapes(1).TextFrame.TextRange.Font.Name = arial
            activeSlide.Shapes(1).TextFrame.TextRange.Font.Size = 14
            activeSlide.Shapes(1).TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
            
    End If ' title
            
' change font size and borders in table

    If heighttable < minheight And wspagefinish(j) < wsrows(wsref) Then ' there are more rows to copy and we have room on page

        wspagefinish(j) = wspagefinish(j) + 1 ' add another row to paste
        activeSlide.Shapes(2).Delete ' delete the pasted shape
        morerows = True
        If j > 1 Then ' remove content for subsequent page to repopulate
        
            Rows(wsrows(i) + 1).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Delete Shift:=xlUp
        
        End If
        
        GoTo Addrows
        
    End If
        
    If wspagefinish(j) >= wsrows(wsref) Then ' finished
    
        GoTo Nextws
    
    Else ' more to copy and need new page
    
        morerows = False
' delete copied rows for current page ready for next page

        Rows(wsrows(i) + 1).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Delete Shift:=xlUp

        GoTo Newpage
    
    End If

Nextws:

' delete copied rows for current page ready for next page

    Rows(wsrows(i) + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp

    Columns("C:C").EntireColumn.Hidden = False
    Columns("F:F").EntireColumn.Hidden = False
    Columns("H:I").EntireColumn.Hidden = False

Range("A1").Select

    
Next i ' next workstream
     
    newPowerPoint.ActiveWindow.ViewType = ppViewNormal
     
'    AppActivate ("Microsoft PowerPoint")
    Set activeSlide = Nothing
    Set newPowerPoint = Nothing
    
''Sheets("Control Panel").Activate
'Range("A1").Select

MsgBox ("Reports complete for " & wsnum & " projects")
     
End Sub

Sub createheader(wsname, rowref, nextrow, wsref)

    With ShtTaskView1
        .Visible = xlSheetVisible
        .Range("C2:M2").Copy
        .Visible = xlSheetHidden
    End With
    Application.CutCopyMode = False
    
    Sheets(wsname).Activate

    Rows(rowref & ":" & rowref).Select
    Selection.Insert Shift:=xlDown

    rowref = rowref + 2

End Sub


