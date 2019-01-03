Attribute VB_Name = "VC_Duplicate"
Sub LockCost(Col1, Col2)
ActiveSheet.Unprotect ' EDITED
Cells.Locked = False
Columns(Col1 & ":" & Col2).EntireColumn.Locked = True
ActiveSheet.Protect Password = "3.14159265359", DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Sub CheckIfExistInArray(CheckThis As Variant, InHere As Variant)
'seems now I am using checking if item is in an array very often
'worth time to do a function so I can use it repeatly.
'2018-08-16 ehh maybe not at this time, as the condition seems to change all the time...
'fuck me,.
End Sub
Sub testunlock()
ActiveSheet.Unprotect Password = "3.14159265359"
'With Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(HeaderLine, EstimateStartLine), Worksheets("Consolidation").Cells(HeaderLine, UBound(EstDataTableHeader) + 1))
'   .Borders(xlEdgeBottom).Weight = xlThick
'    .Borders(xlEdgeBottom).LineStyle = xlContinuous
'    .WrapText = True
'    .Font.FontStyle = "Bold"
'    .Interior.Color = RGB(236, 159, 240)
'End With
End Sub


Sub DuplicateValuesFromColumns()
Call testunlock
'Declare All Variables
Dim myCell As Range
Dim myRow As Integer
Dim myRange As Range
Dim myCol As Integer
Dim i, j, k As Integer
'Dim DuplicatedCellArray As Variant
'Count number of rows & column
myRow = Range(Cells(1, 1), Cells(1, 1).End(xlDown)).Count
myCol = Range(Cells(1, 1), Cells(1, 1).End(xlToRight)).Count

With Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 11), Worksheets("Consolidation").Cells(myRow, 15))
.Interior.Color = RGB(202, 204, 206)
End With

'Loop each column to check duplicate values & highlight them.
'For i = 2 To myRow
i = 5
'j = 1
'k = 2
ReDim DuplicatedCellArray(0)
Set myRange = Range(Cells(2, i), Cells(myRow, i))
For Each myCell In myRange
If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
myCell.Interior.ColorIndex = 36
'ReDim Preserve DuplicatedCellArray(j - 1)
'DuplicatedCellArray(j - 1) = myCell.Value
'meant to store the duplicated value for some other operator
'j = j + 1
'not even used
ElseIf myCell.Value > 0 Then
myCell.Interior.ColorIndex = 35
End If
'Next
'k = k + 1
'row indicator, poorly named variable... not even used... only for debugging
Next
i = 1
j = 2
Set myRange = Range(Cells(2, i), Cells(myRow, i))
For Each myCell In myRange
If myCell.Value = "Header" Then
With Range(Cells(j, i), Cells(j, myCol))
    .Font.FontStyle = "Bold Italic"
    '.Font.Name = ""
    .Interior.Color = RGB(208, 240, 240)
End With
End If

If myCell.Value = "Division Line" Then
With Range(Cells(j, i), Cells(j, myCol))
    .Font.FontStyle = "Bold"
    .Interior.Color = RGB(144, 175, 244)
End With
End If
j = j + 1
Next
Call LockCost("k", "O")
End Sub

Sub NextError()
Call DuplicateValuesFromColumns
Call testunlock
Dim myCell As Range
Dim test
Dim myRow As Integer
Dim myRange As Range
Dim myCol As Integer
Dim i, k, TestTicker As Integer
Dim DuplicatedCellArray As Variant
Dim IsInArray As Boolean
If NextError_Click_Num = Empty Then
NextError_Click_Num = 0
End If
'Count number of rows & column
myRow = Range(Cells(1, 1), Cells(1, 1).End(xlDown)).Count
myCol = Range(Cells(1, 1), Cells(1, 1).End(xlToRight)).Count
'Loop each column to check duplicate values & highlight them.

i = 5

k = 0
ReDim DuplicatedCellArray(0)
Set myRange = Range(Cells(2, i), Cells(myRow, i))

For Each myCell In myRange
    If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
    IsInArray = False
    'there is a duplicate!
        test = myCell.Value 'Test if this value is existing in the array for error checking
        For TestTicker = 1 To UBound(DuplicatedCellArray)
            If DuplicatedCellArray(TestTicker - 1) = test Then
                IsInArray = True
                Exit For
            End If
        Next
        If IsInArray = False Then 'If not in the DuplicatedCellArray then lets add it in
            DuplicatedCellArray(k) = test
            k = k + 1
            ReDim Preserve DuplicatedCellArray(k)
        End If
    End If
Next
ReDim Preserve DuplicatedCellArray(k - 1)
For Each myCell In myRange
    If myCell.Value = DuplicatedCellArray(NextError_Click_Num) Then
        myCell.Interior.Color = RGB(255, 153, 0)
    End If
Next

NextError_Click_Num = NextError_Click_Num + 1
If NextError_Click_Num > UBound(DuplicatedCellArray) Then
MsgBox "!!This is the last one!!"
NextError_Click_Num = 0
End If
Call LockCost("k", "O")
End Sub
