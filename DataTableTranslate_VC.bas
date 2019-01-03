Attribute VB_Name = "DataTableTranslate_VC"
Sub EstDataTableToPMReviewDataTable()
Dim Col As Integer
'instead of using loop, i want to use Application Function such as Copy and Past to Complete this Task
HeaderLine = 1
EstimateStartLine = 1


Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(HeaderLine, EstimateStartLine), Worksheets("Consolidation Temp").Cells(TotalRowCount, UBound(Header) + 1)) = EstDataTable

'Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(HeaderLine, EstimateStartLine), Worksheets("Consolidation").Cells(HeaderLine, UBound(EstDataTableHeader))) = EstDataTableHeader
'With Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(HeaderLine, EstimateStartLine), Worksheets("Consolidation").Cells(HeaderLine, UBound(EstDataTableHeader)))
'    .Borders(xlEdgeBottom).Weight = xlThick
'    .Borders(xlEdgeBottom).LineStyle = xlContinuous
'End With

'Header/Cost Line Copy Paste Target Column 1
Col = 1
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 1), Worksheets("Consolidation").Cells(TotalRowCount + 1, 1)).Select
Worksheets("Consolidation").Paste

'Cost Code Copy Paste, Target Column 5
Col = 2
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 5), Worksheets("Consolidation").Cells(TotalRowCount + 1, 5)).Select
Worksheets("Consolidation").Paste

'Description, Target Column 6
Col = 3
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 6), Worksheets("Consolidation").Cells(TotalRowCount + 1, 6)).Select
Worksheets("Consolidation").Paste
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 6), Worksheets("Consolidation").Cells(TotalRowCount + 1, 6)).ColumnWidth = 30

'Cost Type, Target Column 7
'Col = 3
'Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
'Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 7), Worksheets("Consolidation").Cells(TotalRowCount + 1, 7)).Select
'Worksheets("Consolidation").Paste

'# of Units, Target Column 8
Col = 4
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 8), Worksheets("Consolidation").Cells(TotalRowCount + 1, 8)).Select
Worksheets("Consolidation").Paste

'Unit of Measure, Target Column 9
Col = 5
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 9), Worksheets("Consolidation").Cells(TotalRowCount + 1, 9)).Select
Worksheets("Consolidation").Paste

'Total Hours, Target Column 10
Col = 8
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 10), Worksheets("Consolidation").Cells(TotalRowCount + 1, 10)).Select
Worksheets("Consolidation").Paste

'Total Labour Cost, Target Column 11
Col = 10
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 11), Worksheets("Consolidation").Cells(TotalRowCount + 1, 11)).Select
Worksheets("Consolidation").Paste

'Total Material Cost, Target Column 12
Col = 12
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 12), Worksheets("Consolidation").Cells(TotalRowCount + 1, 12)).Select
Worksheets("Consolidation").Paste

'Total Equipment Cost, Target Column 12
'Total Subcontract Cost, Target Column 14
Col = 14
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 14), Worksheets("Consolidation").Cells(TotalRowCount + 1, 14)).Select
Worksheets("Consolidation").Paste

'Total Cost, Target Column 15
Col = 15
Worksheets("Consolidation Temp").Range(Worksheets("Consolidation Temp").Cells(1, Col), Worksheets("Consolidation Temp").Cells(TotalRowCount, Col)).Copy
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 15), Worksheets("Consolidation").Cells(TotalRowCount + 1, 15)).Select
Worksheets("Consolidation").Paste


'Copy n' Paste New Header to New Sheet
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(HeaderLine, EstimateStartLine), Worksheets("Consolidation").Cells(HeaderLine, UBound(EstDataTableHeader) + 1)) = EstDataTableHeader

With Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(HeaderLine, EstimateStartLine), Worksheets("Consolidation").Cells(HeaderLine, UBound(EstDataTableHeader) + 1))
    .Borders(xlEdgeBottom).Weight = xlThick
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .WrapText = True
    .Font.FontStyle = "Bold"
    .Interior.Color = RGB(236, 159, 240)
        
End With

'Setting up the formula for Cost Type Column 7
'Wait a sec, I dont recall Cost Type matters here...
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 7), Worksheets("Consolidation").Cells(TotalRowCount + 1, 7)).Formula = "=CostType(k2:N2,A2)"
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 2), Worksheets("Consolidation").Cells(TotalRowCount + 1, 2)).Formula = "=ContractItem(E2,A2)"
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 3), Worksheets("Consolidation").Cells(TotalRowCount + 1, 3)).Formula = "=ContractItemDescription(E2,A2)"
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 3), Worksheets("Consolidation").Cells(TotalRowCount + 1, 3)).ColumnWidth = 20
DeleteConsolidationSheet ("Consolidation Temp")


End Sub
Sub BillingScheduleChangetoBaseContract()
On Error GoTo Error_Handler
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 2), Worksheets("Consolidation").Cells(TotalRowCount + 1, 2)).Formula = "=ContractItem_BaseContract(E2,A2)"
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 3), Worksheets("Consolidation").Cells(TotalRowCount + 1, 3)).Formula = "=ContractItemDescription_BaseContract(E2,A2)"
Exit Sub
Error_Handler:
MsgBox "Error 404, Sheet Consolidation Not Found!"
End Sub
Sub BillingScheduleChangetoStandardBillingSchedule()
On Error GoTo Error_Handler
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 2), Worksheets("Consolidation").Cells(TotalRowCount + 1, 2)).Formula = "=ContractItem(E2,A2)"
Worksheets("Consolidation").Range(Worksheets("Consolidation").Cells(2, 3), Worksheets("Consolidation").Cells(TotalRowCount + 1, 3)).Formula = "=ContractItemDescription(E2,A2)"
Exit Sub
Error_Handler:
MsgBox "Error 404, Sheet Consolidation Not Found!"
End Sub

