Attribute VB_Name = "DataToDataTableOnActiveSheet_VC"
Sub ToDataTable(FromSheetName As Variant)
Dim mainworkBook As Workbook
Dim mainworkSheet As Worksheet
Dim SheetNumber, RowCount As Integer
Dim SkippedRow As Integer
On Error Resume Next
Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.Sheets(FromSheetName)
RowCount = mainworkSheet.Cells(Rows.Count, UBound(Header, 1)).End(xlUp).Row
'Row count in this sheet
TotalRowCount = TotalRowCount + RowCount
SkippedRow = 0
ReDim Preserve EstDataTableTrans(UBound(Header, 1), TotalRowCount - 1)
'ReDim EstDataTable(RowCount-1, UBound(Header, 1) - 1)

For rownumber = EstimateStartLine To RowCount

    For Colnumber = 1 To UBound(Header, 1) + 1
    
        If mainworkSheet.Cells(rownumber, SecTitleCol).Value <> "" And mainworkSheet.Cells(rownumber, SecTitleCol - 1).Value = "" Then
    'This tells if it is section header
            EstDataTableTrans(Colnumber - 1, TotalRowCountPrev + rownumber - EstimateStartLine - SkippedRow) = mainworkSheet.Cells(rownumber, Colnumber).Value
            EstDataTableTrans(0, TotalRowCountPrev + rownumber - EstimateStartLine - SkippedRow) = "Header"
    
        ElseIf mainworkSheet.Cells(rownumber, SecTitleCol).Value <> "" And mainworkSheet.Cells(rownumber, UBound(Header, 1) + 1).Value <> 0 And mainworkSheet.Cells(rownumber, 4).Value = 0 Then
    'This tells if it is Division Header line
            EstDataTableTrans(Colnumber - 1, TotalRowCountPrev + rownumber - EstimateStartLine - SkippedRow) = mainworkSheet.Cells(rownumber, Colnumber).Value
            EstDataTableTrans(0, TotalRowCountPrev + rownumber - EstimateStartLine - SkippedRow) = "Division Line"
        ElseIf mainworkSheet.Cells(rownumber, SecTitleCol).Value <> "" And mainworkSheet.Cells(rownumber, UBound(Header, 1) + 1).Value <> 0 Then
    'This tells if it is empty value line
            EstDataTableTrans(Colnumber - 1, TotalRowCountPrev + rownumber - EstimateStartLine - SkippedRow) = mainworkSheet.Cells(rownumber, Colnumber).Value
            EstDataTableTrans(0, TotalRowCountPrev + rownumber - EstimateStartLine - SkippedRow) = "CostLine"
        Else
            SkippedRow = SkippedRow + 1
        Exit For
        End If
    Next
Next
TotalRowCount = TotalRowCount - EstimateStartLine - SkippedRow + 1
TotalRowCountPrev = TotalRowCount
End Sub

