Attribute VB_Name = "VC_ReformatAssembly"
' !!!ALL PUBLIC DELARED VARIABLES ARE HERE!!! HERE AND ONLY HERE!!!
Public AllSheetNames, RecoverableSheetNames
Public EstDataTable, EstDataTableTrans
Public Header
Public HeaderLine, EstimateStartLine As Integer
Public TotalRowCount As Integer
Public SecTitleCol
'Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Public EstDataTableHeader
Public TotalRowCountPrev
Public ContractItemArray
Public ContractItemCount, PhaseCodeCount, CostItemCount As Integer
Public PhaseCodeArray
Public CostItemArray
'use to store amount of rows in previous sheet
Public NextError_Click_Num As Integer
Public ZeroValueRowHiddenStatus As Boolean
Public RecoverableStartLine, LastTimeRecoverable_Row, RecoverableStartColumn As Integer
Public SheetHasChanged As Boolean

Option Explicit
Sub LoadForm(control As IRibbonControl)
'This goes to Step 1
Load SheetRangeForm
SheetRangeForm.Show
End Sub

Sub Reformat(SheetList As Variant)
Call InitializeSettings
Application.CutCopyMode = False
Dim item As Variant
'This function should reformat the target/active spread sheet into a data table with proper headings.
'All formatting should be deleted.
'The result data table should be ready for consolidation of cost code and Contract Items (for billing purpose).
'the new data table will be a sheet named "Consolidation Table". If no such table exist in the target workbook, a new one will be created.
'2018-06-22: I want the first step to be getting all sheet names, this way can give tolerance for future development.
Application.ScreenUpdating = False
'Dim t
't = Timer
'Dim i As Integer
DeleteConsolidationSheet ("Consolidation Temp")
DeleteConsolidationSheet ("Consolidation")
'Below is defined in Workbook_Open() event
'HeaderLine = 3
'EstimateStartLine = 7
'TotalRowCount = 0
'TotalRowCountPrev = 0
'Step 1
'Load SheetRangeForm

'Call FnGetSheetsName
'All of Sheet Names are Stored in Variable "AllSheetName()"

'Reading Estimate Header line and identify columns for code, total cost, description etc...
'Use multi dimensional array to store the data table
Call ReadEstDivHeader(3)
'Current Format is using the column Description to get rid of exceessive
SecTitleCol = Application.Match("Description", Header, False)


'2018-06-22: Ad ReDim can only change the last dimension, data to be poured into EstDataTableReverse then Flip the DataTable.
'ReDim EstDataTable(0, UBound(Header, 1) - 1)
'Now header data is read, data can be poured into EstDataTable based on header name
'what format should the datatable use??? i dont fucking know
'data table should allow PM to review each estmate line, also promt for duplicate costcode, allow modification on Units etc...
'PM Review Data table will be the following format (14 in total):Header/cost line, Contract Item, CostCode Naming Indicate, Cost Code, Description, Cost Type, # of Units, Unit of Measure, Total Hours (for labours only), Total Labour Cost, Total Material Cost, Total Equipment Cost, Total Subcontract Cost, Empty)
EstDataTableHeader = Array("Header/Cost", "Contract Item", "CI Description", "INDCTR", "Cost Code", "Description", "Cost Type", "QTY.", "UoM", "Total Hours", "LAB Cost", "MAT Cost", "EQT Cost", "SUB Cost", "Total Cost")
ReDim EstDataTableTrans((UBound(Header, 1)), 0)

'This is where we select from which sheet to which sheet is the detail estimate
'Can you made auto selection by finding certain strings, not there yet.
For Each item In SheetList
Call ToDataTable(item)
Next
'All Estimate data are now poured into EstDataTableTrans, but this table need to be trimed down as it is oversized with lots empty space at the end.


ReDim Preserve EstDataTableTrans(UBound(Header, 1), TotalRowCount)
ReDim EstDataTable(TotalRowCount, UBound(Header, 1))

EstDataTable = Application.Transpose(EstDataTableTrans)
Call AddConsolidationSheet("Consolidation Temp")
Call AddConsolidationSheet("Consolidation")

'Reuse variables headerline and estmatestartline to set starting point for PM Review Data Table

'Here assumed Header is only 1 row, can be changed to variables if needed.
'reuse variable EstDataTableTrans to reconsile the columns into the correct order.

ReDim EstDataTableTrans(TotalRowCount, UBound(Header, 1))
EstDataTableTrans = EstDataTable

Call EstDataTableToPMReviewDataTable
'Code Runs to here no bugs were found yet....
Call DuplicateValuesFromColumns
'MsgBox Timer - t
'MsgBox GetMemUsage
NextError_Click_Num = 0
Application.CutCopyMode = False
End Sub

Sub ToCSVFile(control As IRibbonControl)
Application.ScreenUpdating = False
'This goes to Step 2
On Error GoTo Error_Handler
Dim FileSaveName
Dim csvVal As String
Dim CSVRow, CSVCol, fnum As Integer
fnum = FreeFile
Call ContractItemSeleteUnique
Call PhaseCodeSeleteUnique
Call CostItemSelectUnique

FileSaveName = Application.GetSaveAsFilename( _
InitialFileName:="Budget Upload" + ".csv", fileFilter:="CSV (*.csv), *.csv")
If FileSaveName = False Then
Exit Sub
End If
Open FileSaveName For Output As #fnum

'Write Contract Items
For CSVRow = 1 To ContractItemCount
    For CSVCol = 1 To 8
        csvVal = csvVal & Chr(34) & ContractItemArray(CSVCol - 1, CSVRow - 1) & Chr(34) & ","
    Next
        Print #fnum, Left(csvVal, Len(csvVal))
        csvVal = ""
Next
'Write Cost Codes
For CSVRow = 1 To PhaseCodeCount
    For CSVCol = 1 To 6
        csvVal = csvVal & Chr(34) & PhaseCodeArray(CSVCol - 1, CSVRow - 1) & Chr(34) & ","
    Next
        Print #fnum, Left(csvVal, Len(csvVal))
        csvVal = ""
Next
'Write Cost Items
For CSVRow = 1 To CostItemCount
    For CSVCol = 1 To 9
        csvVal = csvVal & Chr(34) & CostItemArray(CSVCol - 1, CSVRow - 1) & Chr(34) & ","
    Next
        Print #fnum, Left(csvVal, Len(csvVal))
        csvVal = ""
Next
Close #fnum
MsgBox "CSV Exported"
Exit Sub
Error_Handler:
MsgBox "Error 404, Sheet Consolidation Not Found!"
End Sub

Sub Refresh(control As IRibbonControl)
Application.ScreenUpdating = False
Dim mainworkSheet As Worksheet
Dim mainworkBook As Workbook
On Error GoTo Error_Handler
Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.Sheets("Consolidation")
mainworkSheet.Activate
'Call DuplicateValuesFromColumns
Call NextError
Exit Sub
'Call LockCost("k", "O")
Error_Handler:
MsgBox "Sheet Consolidation or Duplicates Not Found!"
End Sub

Sub Hide_UnhideProductivity(control As IRibbonControl)
'Estimating Function
Application.ScreenUpdating = False
Dim mainworkSheet As Worksheet
Dim mainworkBook As Workbook
On Error GoTo Error_Handler
Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.ActiveSheet
Dim ColumnHiddenStatus As Boolean
ColumnHiddenStatus = mainworkSheet.Columns("F:F").Hidden
If ColumnHiddenStatus = False Then
mainworkSheet.Columns("F:F").Hidden = True
Else
mainworkSheet.Columns("F:F").Hidden = False
End If
Exit Sub
Error_Handler:
MsgBox "There seems to be some error, please see Siyuan."
End Sub

Sub Hide_Unhide_ZeroValueRow(control As IRibbonControl) '(control As IRibbonControl)
'Estimating Function
Application.ScreenUpdating = False
On Error GoTo Error_Handler
Call Est_HideUnHideZeroValueRow
Exit Sub
Error_Handler:
MsgBox "There seems to be some error, please see Siyuan."
End Sub
Sub ResetRefresh(control As IRibbonControl)
NextError_Click_Num = 0
End Sub
Sub ChangetoBaseContract(control As IRibbonControl)
Call BillingScheduleChangetoBaseContract
End Sub
Sub ChangetoStandardBillingSchedule(control As IRibbonControl)
Call BillingScheduleChangetoStandardBillingSchedule
End Sub
Sub ProjectRecoverableLoadForm(control As IRibbonControl)
'Estimating Function
Load RecoverableForm2
RecoverableForm2.Show

End Sub
Sub PPBookLoadForm(control As IRibbonControl)
'Estimating Function
Load PPBookSelection
PPBookSelection.Show

End Sub

