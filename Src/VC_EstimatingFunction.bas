Attribute VB_Name = "VC_EstimatingFunction"
Sub Est_HideUnHideZeroValueRow()
Dim mainworkSheet As Worksheet
Dim mainworkBook As Workbook
Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.ActiveSheet
Dim ticker, maxticker As Integer
maxticker = WorksheetFunction.Min(mainworkSheet.Cells(Rows.Count, 3).End(xlUp).Row, mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row, mainworkSheet.Cells(Rows.Count, 15).End(xlUp).Row)

For ticker = EstimateStartLine To maxticker
'test watch row below.
'ticker = mainworkSheet.Cells(Rows.Count, 3).End(xlUp).Rows

If (ZeroValueRowHiddenStatus = False) And (mainworkSheet.Cells(ticker, 4) <> "") And (mainworkSheet.Cells(ticker, 15) = 0) Then
mainworkSheet.Rows(ticker).Hidden = True
Else
mainworkSheet.Rows(ticker).Hidden = False
End If
If ZeroValueRowHiddenStatus = True Then
mainworkSheet.Rows(ticker).Hidden = False
End If
Next
ZeroValueRowHiddenStatus = Not ZeroValueRowHiddenStatus
End Sub



Sub ProjectRecoverable_New(SheetList As Variant)
On Error Resume Next
Application.CutCopyMode = False
Dim SheetNumber, RowCount, Row, LabourNameCount_Staff, LabourNameCount_Craft, i As Integer
Dim LoopStart, LoopEnd As Integer
Dim IsInArray As Boolean
Dim RecoverableArray_Staff, RecoverableArray_Craft, RecoverableHeader '(labour type (staff/craft), labour name, labour total MH, labour MH rate)
Dim mainworkSheet As Worksheet
Dim mainworkBook As Workbook
Set mainworkBook = ActiveWorkbook
ReDim RecoverableArray_Staff(4, 0)
ReDim RecoverableArray_Craft(4, 0)
Application.ScreenUpdating = False

LabourNameCount_Staff = 0
LabourNameCount_Craft = 0
For Each item In SheetList
    Set mainworkSheet = mainworkBook.Sheets(item)
    RowCount = mainworkSheet.Cells(Rows.Count, 3).End(xlUp).Row
    For Row = EstimateStartLine To RowCount
    If (mainworkSheet.Cells(Row, 1).Value = "s" Or mainworkSheet.Cells(Row, 1).Value = "S") And (mainworkSheet.Cells(Row, 8).Value > 0) Then
        RecoverableArray_Staff(0, LabourNameCount_Staff) = "Staff"
        RecoverableArray_Staff(1, LabourNameCount_Staff) = mainworkSheet.Cells(Row, 3).Value
        RecoverableArray_Staff(2, LabourNameCount_Staff) = mainworkSheet.Cells(Row, 8).Value
        RecoverableArray_Staff(3, LabourNameCount_Staff) = mainworkSheet.Cells(Row, 9).Value
        LabourNameCount_Staff = LabourNameCount_Staff + 1
        ReDim Preserve RecoverableArray_Staff(4, LabourNameCount_Staff)
    ElseIf (mainworkSheet.Cells(Row, 8).Value > 0) And (mainworkSheet.Cells(Row, 15).Value > 0) Then
    
        For i = 0 To LabourNameCount_Craft
            IsInArray = False
            If RecoverableArray_Craft(3, i) = mainworkSheet.Cells(Row, 9).Value Then
                RecoverableArray_Craft(2, i) = mainworkSheet.Cells(Row, 8) + RecoverableArray_Craft(2, i)
                IsInArray = True
                Exit For
            End If
        Next
        
        If IsInArray = False Then
            RecoverableArray_Craft(0, LabourNameCount_Craft) = "Craft"
            RecoverableArray_Craft(1, LabourNameCount_Craft) = mainworkSheet.Cells(Row, 3).Value
            RecoverableArray_Craft(2, LabourNameCount_Craft) = mainworkSheet.Cells(Row, 8).Value
            RecoverableArray_Craft(3, LabourNameCount_Craft) = mainworkSheet.Cells(Row, 9).Value
            LabourNameCount_Craft = LabourNameCount_Craft + 1
            ReDim Preserve RecoverableArray_Craft(4, LabourNameCount_Craft)
        End If
    End If
    Next
Next
Call AddConsolidationSheet("Recoverable Temp")
Set mainworkSheet = mainworkBook.Sheets("Recoverable Temp")
RecoverableHeader = Array("Type", "Description", "Current Est MHs", "Charge Out Rate", "Cost + Burden", "Delta", "Total Recoverable", "%")
With mainworkSheet

'Start writing data
For i = 0 To 7
.Cells(RecoverableStartLine - 1, i + 1).Value = RecoverableHeader(i)
Next
End With
With mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine - 1, RecoverableStartColumn), mainworkSheet.Cells(RecoverableStartLine - 1, RecoverableStartColumn + 7))
    .Borders(xlEdgeBottom).Weight = xlThick
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .WrapText = True
    .Font.FontStyle = "Bold"
    .Interior.Color = RGB(128, 100, 162)
    .ColumnWidth = 17
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Name = "Arial"
    .Font.Color = vbWhite
End With
'Leave 1 line for the Total
LoopStart = RecoverableStartLine + 1
LoopEnd = LabourNameCount_Staff + RecoverableStartLine
For Row = LoopStart To LoopEnd
mainworkSheet.Cells(Row, RecoverableStartColumn).Value = RecoverableArray_Staff(0, Row - LoopStart)
mainworkSheet.Cells(Row, RecoverableStartColumn + 1).Value = RecoverableArray_Staff(1, Row - LoopStart)
mainworkSheet.Cells(Row, RecoverableStartColumn + 2).Value = RecoverableArray_Staff(2, Row - LoopStart)
mainworkSheet.Cells(Row, RecoverableStartColumn + 3).Value = RecoverableArray_Staff(3, Row - LoopStart)
Next

LoopStart = LabourNameCount_Staff + RecoverableStartLine + 1
LoopEnd = LabourNameCount_Craft + LabourNameCount_Staff + RecoverableStartLine

For Row = LoopStart To LoopEnd
mainworkSheet.Cells(Row, RecoverableStartColumn).Value = RecoverableArray_Craft(0, Row - LoopStart)
mainworkSheet.Cells(Row, RecoverableStartColumn + 1).Value = "Labour " & (Row - LoopStart + 1)
mainworkSheet.Cells(Row, RecoverableStartColumn + 2).Value = RecoverableArray_Craft(2, Row - LoopStart)
mainworkSheet.Cells(Row, RecoverableStartColumn + 3).Value = RecoverableArray_Craft(3, Row - LoopStart)
Next

LoopStart = LabourNameCount_Craft + LabourNameCount_Staff + RecoverableStartLine + 1
LoopEnd = LabourNameCount_Craft + LabourNameCount_Staff + RecoverableStartLine + Material_Rec_Quantity

'Material line are not 0, not being used.
For Row = LoopStart To LoopEnd
mainworkSheet.Cells(Row, RecoverableStartColumn + 1).Value = "Material Recoverable " & (Row - LoopStart + 1)
mainworkSheet.Cells(Row, RecoverableStartColumn + 6).Interior.Color = RGB(211, 253, 173)
Next
'Nonrecoverable line
mainworkSheet.Cells(Row, RecoverableStartColumn).Value = ""
mainworkSheet.Cells(Row, RecoverableStartColumn + 1).Value = "Non-Recoverable"
mainworkSheet.Cells(Row, RecoverableStartColumn + 6).Interior.Color = RGB(211, 253, 173)
'total recoverable calculation
mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn).Value = ""
mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn + 1).Value = "Total"
mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn), mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn + 7)).Font.FontStyle = "Bold"
mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn + 6).Value = "=AGGREGATE(9,2,G:G)"

mainworkSheet.Cells(Row + 2, RecoverableStartColumn).Value = "Notes:"
mainworkSheet.Cells(Row + 2, RecoverableStartColumn + 6).Value = "Other Cost:"

'Formatting
With mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn + 3), mainworkSheet.Cells(Row, RecoverableStartColumn + 6))
.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
End With
With mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn), mainworkSheet.Cells(Row, RecoverableStartColumn + 7))
.Interior.Color = RGB(234, 234, 234)
End With
mainworkSheet.Cells(Row, RecoverableStartColumn).Value = ""
mainworkSheet.Cells(Row, RecoverableStartColumn + 1).Value = "Non-Recoverable"
mainworkSheet.Cells(Row, RecoverableStartColumn + 6).Interior.Color = RGB(211, 253, 173)

If (LabourNameCount_Craft + LabourNameCount_Staff) > 0 Then
'only labour cost need to calculate Burden
With mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine + 1, RecoverableStartColumn + 5), mainworkSheet.Cells(LabourNameCount_Craft + LabourNameCount_Staff + RecoverableStartLine, RecoverableStartColumn + 5))
.Formula = "=D4-E4"
End With
'Recoverable of eachline
With mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine + 1, RecoverableStartColumn + 6), mainworkSheet.Cells(LabourNameCount_Craft + LabourNameCount_Staff + RecoverableStartLine, RecoverableStartColumn + 6))
.Formula = "=C4*F4"
End With
With mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine + 1, RecoverableStartColumn + 4), mainworkSheet.Cells(LabourNameCount_Craft + LabourNameCount_Staff + RecoverableStartLine, RecoverableStartColumn + 4))
.Interior.Color = RGB(211, 253, 173)
End With
End If
With mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn + 7), mainworkSheet.Cells(Row, RecoverableStartColumn + 7))
.Formula = "=G3/Summary!$J$2"
.NumberFormat = "0.00%"
End With


With mainworkSheet.Range(mainworkSheet.Cells(RecoverableStartLine, RecoverableStartColumn), mainworkSheet.Cells(Row, RecoverableStartColumn + 7))
    .Borders(xlEdgeBottom).Weight = xlThin
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Color = RGB(128, 100, 162)
    .WrapText = False
    '.Interior.Color = RGB(128, 100, 162)
    '.HorizontalAlignment = xlCenter
    '.VerticalAlignment = xlCenter
    .Font.Name = "Arial"
End With

With mainworkSheet.Range(mainworkSheet.Cells(Row + 2, RecoverableStartColumn), mainworkSheet.Cells(Row + 2, RecoverableStartColumn + 7))
    .Borders(xlEdgeBottom).Weight = xlThick
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    .WrapText = False
    .Font.Name = "Arial"
End With
With mainworkSheet.Range(mainworkSheet.Cells(Row + 3, RecoverableStartColumn + 6), mainworkSheet.Cells(Row + 13, RecoverableStartColumn + 6))
    .WrapText = False
    .Font.Name = "Arial"
    .Interior.Color = RGB(211, 253, 173)
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
End With
'Copy
mainworkBook.Worksheets("Recoverable").Range(Worksheets("Recoverable").Rows(LastTimeRecoverable_Row + 3), Worksheets("Recoverable").Rows(LastTimeRecoverable_Row + 13)).Copy
'paste
mainworkSheet.Range(mainworkSheet.Rows(Row + 3), mainworkSheet.Rows(Row + 100)).Select
mainworkSheet.Paste
Call DeleteConsolidationSheet("Recoverable")
'Rename
LastTimeRecoverable_Row = Row
With mainworkSheet.Cells
.Font.Name = "Arial"
End With
'Adjust Page Setup
With mainworkSheet.PageSetup
.Zoom = False
.FitToPagesWide = 1
.PaperSize = xlPaperLetter
.Orientation = xlLandscape
.CenterHeader = "Project Recoverable"
End With
mainworkSheet.Name = "Recoverable"
mainworkSheet.Cells(Row + 3, 2).Select
Application.CutCopyMode = False
End Sub

