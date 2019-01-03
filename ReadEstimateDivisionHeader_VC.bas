Attribute VB_Name = "ReadEstimateDivisionHeader_VC"
Sub ReadEstDivHeader(FirstEstimateSheet)
Dim mainworkBook As Workbook
Dim mainworkSheet As Worksheet
'Dim Header
'2018-06-22: Header Dim'd as Public
Dim ColumnNum, RowNum, ColumnCount As Integer

Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.Sheets(FirstEstimateSheet)
'RowCount = mainworkSheet.Cells(Rows.Count, 14).End(xlUp).Row
ColumnCount = mainworkSheet.Cells(3, Columns.Count).End(xlToLeft).Column
'MsgBox ColumnCount
'MsgBox RowCount

ReDim Header(ColumnCount - 1, 0)
For ColumnNum = 1 To ColumnCount
Header(ColumnNum - 1, 0) = mainworkSheet.Cells(HeaderLine, ColumnNum).Value
Next

End Sub
