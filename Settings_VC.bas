Attribute VB_Name = "Settings_VC"

Sub InitializeSettings()
HeaderLine = 3
EstimateStartLine = 7
TotalRowCount = 0
TotalRowCountPrev = 0
ZeroValueRowHiddenStatus = False
RecoverableStartLine = 3
RecoverableStartColumn = 1
If sheetExists("Recoverable") = False Then
LastTimeRecoverable_Row = 0
Else
Dim mainworkSheet As Worksheet
Dim mainworkBook As Workbook
Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.Worksheets("Recoverable")
LastTimeRecoverable_Row = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
End If
End Sub

Sub PPBookSettings()
PPBookDataStartingLine = 10
End Sub

