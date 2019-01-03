Attribute VB_Name = "PPBookFunctions_VC"
#If VBA7 Then
  Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
   Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub PPBook_PrintToPDF()
Dim FileSaveName

FileSaveName = Application.GetSaveAsFilename( _
InitialFileName:="Budget Upload" + ".csv", fileFilter:="CSV (*.csv), *.csv")
If FileSaveName = False Then
Exit Sub
End If
ThisWorkbook.Sheets(Array("Sheet1", "Sheet2")).Select

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    "C:\tempo.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
     IgnorePrintAreas:=False, OpenAfterPublish:=True

End Sub

Function PPBook_HeaderSummaryDetail(ByVal rownumber As Integer, SheetName As Variant)

'i guess this could be used for other modules for the same purpose but i am too lazy to change names....
If (Worksheets(SheetName).Cells(rownumber, 3).Value = "") Then
PPBook_HeaderSummaryDetail = "EmptyLine"
ElseIf Worksheets(SheetName).Cells(rownumber, 1).Value = "s" Then
PPBook_HeaderSummaryDetail = "Staff"
ElseIf Worksheets(SheetName).Cells(rownumber, 1).Value = "D" Then
PPBook_HeaderSummaryDetail = "Division"
ElseIf Worksheets(SheetName).Cells(rownumber, 1).Value = "SD" Then
PPBook_HeaderSummaryDetail = "Sum Division"
ElseIf Worksheets(SheetName).Cells(rownumber, 1).Value = "H" Then
PPBook_HeaderSummaryDetail = "Heading"
ElseIf Worksheets(SheetName).Cells(rownumber, 1).Value = "SH" Then
PPBook_HeaderSummaryDetail = "Sum Heading"
ElseIf (Worksheets(SheetName).Cells(rownumber, 15).Value > 0) And (Worksheets(SheetName).Cells(rownumber, 5).Value <> "") Then
PPBook_HeaderSummaryDetail = "Detail"
ElseIf (Worksheets(SheetName).Cells(rownumber, 15).Value = 0) And (Worksheets(SheetName).Cells(rownumber, 5).Value <> "") Then
PPBook_HeaderSummaryDetail = "EmptyLine"
Else
PPBook_HeaderSummaryDetail = "Detail"
End If

End Function
Sub PPBook_WriteTOPPBook(ByVal ToLine As Integer, ByVal FromLine As Integer, ByVal DataType As String, ByVal ToSheet As String, FromSheet As Variant)
With Worksheets(ToSheet)
        .Cells(ToLine, 1) = DataType
        .Cells(ToLine, 2) = Worksheets(FromSheet).Cells(FromLine, 3)
        .Cells(ToLine, 3) = Worksheets(FromSheet).Cells(FromLine, 4)
        .Cells(ToLine, 4) = Worksheets(FromSheet).Cells(FromLine, 5)
        .Cells(ToLine, 5) = Worksheets(FromSheet).Cells(FromLine, 15)
    End With
End Sub

Sub PPBook_FormatSummaryAndDetail(CurrentRow As Variant, CurrentSheet As Variant, NextSummaryRow As Integer)

Dim tempticker As Integer
Dim SumOfDetail As Variant
SumOfDetail = 0
Dim DetailDescription As String
For tempticker = CurrentRow + 1 To NextSummaryRow - 1
SumOfDetail = SumOfDetail + Worksheets(CurrentSheet).Cells(tempticker, 5).Value
DetailDescription = Worksheets(CurrentSheet).Cells(tempticker, 2).Value & " : ----- " & Worksheets(CurrentSheet).Cells(tempticker, 3).Value & "  " & Worksheets(CurrentSheet).Cells(tempticker, 4).Value & "  $ " & Round(Worksheets(CurrentSheet).Cells(tempticker, 5).Value, 0)
Worksheets(CurrentSheet).Cells(tempticker, 2).Value = DetailDescription
    With Worksheets(CurrentSheet).Range(Cells(tempticker, 1), Cells(tempticker, 5))
    
    .Font.Name = "Arial"
    .Font.Size = 8
    .IndentLevel = 2
    .Font.FontStyle = "Normal"
    End With
'Worksheets(CurrentSheet).Cells(tempticker, 1).Value = ""
Worksheets(CurrentSheet).Cells(tempticker, 3).Value = ""
Worksheets(CurrentSheet).Cells(tempticker, 4).Value = ""
Worksheets(CurrentSheet).Cells(tempticker, 5).Value = ""
Next

Worksheets(CurrentSheet).Cells(CurrentRow, 5).Value = Round(SumOfDetail, 0)
With Worksheets(CurrentSheet).Range(Cells(CurrentRow, 1), Cells(CurrentRow, 5))
    .IndentLevel = 1
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End With
'Worksheets(CurrentSheet).Cells(CurrentRow, 1).Value = ""

End Sub
Sub PPBook_ColumnFormatting(CurrentSheet)
Worksheets(CurrentSheet).Columns(2).ColumnWidth = 40
'Worksheets(CurrentSheet).Columns(1).Font.Color = RGB(255, 255, 255)
'the line below is such a powerful line lol
'Worksheets(CurrentSheet).Columns(1) = ""
'Might as well do page breaks here


End Sub
Sub PPBook_DivisionFormat(CurrentRow As Variant, CurrentSheet As Variant)
With Worksheets(CurrentSheet).Range(Cells(CurrentRow, 1), Cells(CurrentRow, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 12
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .Borders(xlEdgeBottom).Weight = xlThin
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    .WrapText = False
    .Interior.Color = RGB(217, 217, 217)
End With
End Sub

Sub PPBook_HeaderAndFooter(DetailPage As Variant)
Worksheets(DetailPage).Rows(1).Insert shift:=xlShiftDown
Worksheets(DetailPage).Rows(1).Insert shift:=xlShiftDown
Worksheets(DetailPage).Rows(1).Insert shift:=xlShiftDown
Worksheets(DetailPage).Rows(1).Insert shift:=xlShiftDown
Worksheets(DetailPage).Rows(1).Insert shift:=xlShiftDown

With Worksheets(DetailPage)
    .Cells(1, 4).Formula = "=""Project Start:  ""&TEXT(Summary!I4,""m/d/yyyy"")"
    .Cells(1, 1).Formula = "= ""Project: ""  & xlProjectName"
    .Cells(2, 1).Formula = "=""Location: ""&xlProjectLocation"
    .Cells(2, 4).Formula = "=""Estimator: ""&xlEstimatorName"
    
    .Cells(4, 1).Value = "Based on information presently available" _
    & " and furnished to SPL by the owner, architect, and/or others and various" _
    & " assumptions which have been made as to facts not yet know, this construction" _
    & " cost estimate has been prepared and furnished for the sole purpose of" _
    & " providing approximation of anticipated construction cost, this construction" _
    & " estimate should not, at this time be relied upon as a commitment that the" _
    & " contemplated project can or will be constructed for the estimated cost."
    With .Range(Cells(4, 1), Cells(4, 5))
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
        .NumberFormat = "General"
        .HorizontalAlignment = xlLeft
        .MergeCells = True
        .WrapText = True
        .RowHeight = 45
    End With
    With .Range(Cells(1, 1), Cells(4, 5))
        .IndentLevel = 0
        .Font.Name = "Arial"
        .Font.Size = 8
        '.Font.FontStyle = "Bold"
        .NumberFormat = "General"
    End With
End With

Worksheets(DetailPage).PageSetup.LeftHeaderPicture.Filename = _
        "http://synergybuilds.com/wp-content/themes/synergy_min_20170110/img/synergy-construction-logo.png"
    With Worksheets(DetailPage).PageSetup.LeftHeaderPicture
        .Height = 28.5
        .Width = 65.25
    End With
    With Worksheets(DetailPage).PageSetup
        .LeftHeader = "&G"
        .LeftHeaderPicture.Height = 28.5
        .LeftHeaderPicture.Width = 65.25
        .CenterHeader = _
        "&""Arial,Bold""&12Synergy Projects Inc.&10" & Chr(10) & "Conceptual Estimate Summary"
        .RightHeader = "&""Arial,Regular""&8Page: &P of &N" & Chr(10) & "&D" & Chr(10) & "&T"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
        .PrintTitleRows = "$1:$5"
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True


End Sub
Sub test()
'PPBook_HeaderAndFooter ("PP")
Worksheets("PP CoverPage").Visible = True
End Sub


Sub PPBook_SummaryPage(DetailPage As Variant, DetailStartLine As Variant)

'damn this is one long piece of code.
'the intent is to add the Summary page section to the detail page that is generated.
Dim tempticker, DivisionCount, i As Integer
Dim Array_Division As Variant
Dim TotalCost, StaffCost, InsuranceCost, FeePerctage, FeeCost, DirectCost, GeneralCost
Dim TargetLine, CurrentDivisionLine
Dim temptoggle As Boolean
ReDim Array_Division(1, 0) ' {(Name,cost);...}
TargetLine = DetailStartLine
TotalCost = 0
'StaffCost = 0
'InsuranceCost = 0
DivisionCount = 0
DirectCost = 0
GeneralCost = 0
temptoggle = False
For tempticker = DetailStartLine To Worksheets(DetailPage).Cells(Rows.Count, 2).End(xlUp).Row
If Worksheets(DetailPage).Cells(tempticker, 1).Value = "Division" And Worksheets(DetailPage).Cells(tempticker, 5).Value > 0 Then
TotalCost = TotalCost + Worksheets(DetailPage).Cells(tempticker, 5).Value
Array_Division(0, DivisionCount) = Worksheets(DetailPage).Cells(tempticker, 2).Value
Array_Division(1, DivisionCount) = Worksheets(DetailPage).Cells(tempticker, 5).Value
DivisionCount = DivisionCount + 1
ReDim Preserve Array_Division(1, DivisionCount)
End If

'not using staff cost, so not running these codes.
If Worksheets(DetailPage).Cells(tempticker, 1).Value = "Staff" Then
StaffCost = StaffCost + Worksheets(DetailPage).Cells(tempticker, 5).Value
End If

Next


'All Division and Cost now poured in.
'Now the General Expenses are special here, some of the cost are not in division 1... Geeeez....
'Project staff, Project Overhead, Insurance, Others????
'I decided to pull data in from summary page.
'fast coding then iterations...
'start writing, this is the boring part

Worksheets(DetailPage).Rows(TargetLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(TargetLine, 2).Value = "Direct Costs"
Worksheets(DetailPage).Cells(TargetLine, 1).Value = "SummaryPage"
With Worksheets(DetailPage).Range(Cells(TargetLine, 1), Cells(TargetLine, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 12
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .WrapText = False

End With
TargetLine = TargetLine + 1
Worksheets(DetailPage).Rows(TargetLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(TargetLine, 1).Value = "SummaryPage"

'THis is where to sum up the direct cost
For tempticker = 1 To DivisionCount - 1

Worksheets(DetailPage).Rows(TargetLine + tempticker).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(TargetLine + tempticker, 2).Value = Array_Division(0, tempticker)
Worksheets(DetailPage).Cells(TargetLine + tempticker, 5).Value = Array_Division(1, tempticker)
DirectCost = DirectCost + Array_Division(1, tempticker)
Worksheets(DetailPage).Cells(TargetLine + tempticker, 1).Value = "SummaryPage"

Next
With Worksheets(DetailPage).Range(Cells(TargetLine + 1, 1), Cells(TargetLine + tempticker - 1, 5))
    .IndentLevel = 1
    .Font.Name = "Arial"
    .Font.Size = 10
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .WrapText = False
End With
Worksheets(DetailPage).Rows(TargetLine + tempticker).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(TargetLine + tempticker, 1).Value = "SummaryPage"
tempticker = tempticker + 1
Worksheets(DetailPage).Rows(TargetLine + tempticker).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(TargetLine + tempticker, 2).Value = "Direct Costs"
Worksheets(DetailPage).Cells(TargetLine + tempticker, 1).Value = "SummaryPage"
Worksheets(DetailPage).Cells(TargetLine + tempticker, 5).Value = DirectCost
With Worksheets(DetailPage).Range(Cells(TargetLine + tempticker, 1), Cells(TargetLine + tempticker, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .Borders(xlEdgeBottom).Weight = xlThin
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    .Borders(xlEdgeTop).Weight = xlThin
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
    .WrapText = False
    .Interior.Color = RGB(217, 217, 217)
End With
CurrentDivisionLine = TargetLine + tempticker + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(CurrentDivisionLine, 1).Value = "SummaryPage"
CurrentDivisionLine = CurrentDivisionLine + 1

Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(CurrentDivisionLine, 2).Value = "General Expense Costs"
Worksheets(DetailPage).Cells(CurrentDivisionLine, 1).Value = "SummaryPage"

With Worksheets(DetailPage).Range(Cells(TargetLine + tempticker + 1, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 12
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .WrapText = False
    .Interior.ColorIndex = xlNone
End With
CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(CurrentDivisionLine, 2).Value = "Project Coordination"
Worksheets(DetailPage).Cells(CurrentDivisionLine, 1).Value = "SummaryPage"

For tempticker = 1 To 999
If Worksheets("Summary").Cells(tempticker, 3).Value = "Total General Conditions" Then
Exit For
End If
Next

Worksheets(DetailPage).Cells(CurrentDivisionLine, 5).Value = Worksheets("Summary").Cells(tempticker, 10).Value
GeneralCost = GeneralCost + Worksheets("Summary").Cells(tempticker, 10).Value

With Worksheets(DetailPage).Range(Cells(CurrentDivisionLine, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 1
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .WrapText = False
    .Interior.ColorIndex = xlNone
End With
For tempticker = 1 To 999
If Worksheets("Summary").Cells(tempticker, 3).Value = "G.C's & Misc. Costs" Then
Exit For
End If
Next
i = tempticker

For tempticker = i To 999

If Worksheets("Summary").Cells(tempticker, 3).Value = "Total Cost before Fee" Then

Exit For

End If

If Worksheets("Summary").Cells(tempticker, 10).Value > 0 Then
CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(CurrentDivisionLine, 2).Value = Trim(Worksheets("Summary").Cells(tempticker, 3).Value)
Worksheets(DetailPage).Cells(CurrentDivisionLine, 5).Value = Worksheets("Summary").Cells(tempticker, 10).Value
GeneralCost = GeneralCost + Worksheets("Summary").Cells(tempticker, 10).Value
Worksheets(DetailPage).Cells(CurrentDivisionLine, 1).Value = "SummaryPage"
With Worksheets(DetailPage).Range(Cells(CurrentDivisionLine, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 1
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .WrapText = False
    .Interior.ColorIndex = xlNone
End With
End If


Next


CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(CurrentDivisionLine, 1).Value = "SummaryPage"

CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(CurrentDivisionLine, 2).Value = "General Expense Costs"
Worksheets(DetailPage).Cells(CurrentDivisionLine, 5).Value = GeneralCost
Worksheets(DetailPage).Cells(CurrentDivisionLine, 1).Value = "SummaryPage"
With Worksheets(DetailPage).Range(Cells(CurrentDivisionLine, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .Borders(xlEdgeBottom).Weight = xlThin
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    .Borders(xlEdgeTop).Weight = xlThin
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
    .WrapText = False
    .Interior.Color = RGB(217, 217, 217)
End With
CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
With Worksheets(DetailPage).Range(Cells(CurrentDivisionLine, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .WrapText = False
    .Interior.ColorIndex = xlNone
End With

'CurrentDivisionLine = CurrentDivisionLine + 1
'Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(CurrentDivisionLine, 1).Value = "SummaryPage"
Worksheets(DetailPage).Cells(CurrentDivisionLine, 2).Value = "Total Cost before Fee"
Worksheets(DetailPage).Cells(CurrentDivisionLine, 5).Value = DirectCost + GeneralCost
With Worksheets(DetailPage).Range(Cells(CurrentDivisionLine, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .Borders(xlEdgeBottom).Weight = xlThin
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    .Borders(xlEdgeTop).Weight = xlThin
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
    .WrapText = False
    .Interior.Color = RGB(217, 217, 217)
End With

For tempticker = 1 To 999

If Trim(Worksheets("Summary").Cells(tempticker, 3).Value) = "Fee" Then
temptoggle = True
End If
If temptoggle = True Then
CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
With Worksheets(DetailPage).Range(Cells(CurrentDivisionLine, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .WrapText = False
    .Interior.ColorIndex = xlNone
End With
CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
Worksheets(DetailPage).Cells(CurrentDivisionLine, 1).Value = "SummaryPage"
Worksheets(DetailPage).Cells(CurrentDivisionLine, 2).Value = Trim(Worksheets("Summary").Cells(tempticker, 3).Value)
Worksheets(DetailPage).Cells(CurrentDivisionLine, 5).Value = Trim(Worksheets("Summary").Cells(tempticker, 10).Value)
With Worksheets(DetailPage).Range(Cells(CurrentDivisionLine, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .Borders(xlEdgeBottom).Weight = xlThin
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    .Borders(xlEdgeTop).Weight = xlThin
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
    .WrapText = False
    .Interior.Color = RGB(217, 217, 217)
End With

End If

If Trim(Worksheets("Summary").Cells(tempticker, 3).Value) = "Total Amount with GST" Then
temptoggle = False
Exit For
End If
Next
Worksheets(DetailPage).Rows(CurrentDivisionLine + 1).PageBreak = xlPageBreakManual


CurrentDivisionLine = CurrentDivisionLine + 1
Worksheets(DetailPage).Rows(CurrentDivisionLine).Insert shift:=xlShiftDown
With Worksheets(DetailPage).Range(Cells(CurrentDivisionLine, 1), Cells(CurrentDivisionLine, 5))
    .IndentLevel = 0
    .Font.Name = "Arial"
    .Font.Size = 10
    .Font.FontStyle = "Bold"
    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    .WrapText = False
    .Interior.ColorIndex = xlNone
End With

'now put in the percentage
TotalCost = DirectCost + GeneralCost
For tempticker = DetailStartLine To CurrentDivisionLine
    If Worksheets(DetailPage).Cells(tempticker, 5).Value > 0 Then
    Worksheets(DetailPage).Cells(tempticker, 4).Value = Worksheets(DetailPage).Cells(tempticker, 5).Value / TotalCost
    Worksheets(DetailPage).Cells(tempticker, 4).NumberFormat = "0.00%"
    End If
    
    If Worksheets(DetailPage).Cells(tempticker, 2).Value = "Total Cost before Fee" Then
    Exit For
    End If
Next

End Sub

Sub PPBook_DetailPage(ByVal SheetList As Variant)
'Application.PrintCommunication = False
Application.ScreenUpdating = False
DoEvents
Call DeleteConsolidationSheet("PP")
Call AddConsolidationSheet("PP")
Dim LineType As String
Dim mainworkSheet As Worksheet
Dim mainworkBook As Workbook
Set mainworkBook = ActiveWorkbook
Application.CutCopyMode = False
Dim TargetLine As Integer
Dim HeadingRow As Integer
Dim SumOfHeading
Dim LastDivisionName
TargetLine = 2
For Each item In SheetList
    Set mainworkSheet = mainworkBook.Worksheets(item)
    RowCount = WorksheetFunction.Min(mainworkSheet.Cells(Rows.Count, 3).End(xlUp).Row, mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row, mainworkSheet.Cells(Rows.Count, 15).End(xlUp).Row)
    For Row = EstimateStartLine - 2 To RowCount
    LineType = PPBook_HeaderSummaryDetail(Row, item)
    
    If LineType <> "EmptyLine" Then
    
    Call PPBook_WriteTOPPBook(TargetLine, Row, LineType, "PP", item)
    TargetLine = TargetLine + 1
    End If
    Next
Next
' Code runs to here fine this pulls all the data in. GOOD!
' "PP" sheet is used for debug, you need to replace it with a variable.
Set mainworkSheet = mainworkBook.Worksheets("PP")
TargetLine = 2
RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
For Row = TargetLine To RowCount
LineType = mainworkSheet.Cells(Row, 1).Value
    If LineType = "Heading" Then
        For HeadingRow = Row + 1 To RowCount
            If mainworkSheet.Cells(HeadingRow, 1).Value = "Heading" Or mainworkSheet.Cells(HeadingRow, 1).Value = "Sum Heading" Or mainworkSheet.Cells(HeadingRow, 1).Value = "Division" Then
            Exit For
            End If
        Next
        Call PPBook_FormatSummaryAndDetail(Row, "PP", HeadingRow)
        Row = HeadingRow - 1
    End If
Next
'code runs here fine.
'now deleting the o dollar heading and all sum heading as they are useless
'The code below to next reference also add space before and after heading to format it better.
RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
For Row = TargetLine To RowCount + 999
    LineType = mainworkSheet.Cells(Row, 1).Value
    If LineType = "Sum Heading" Then
        mainworkSheet.Rows(Row).Delete
        Row = Row - 1
        RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
    ElseIf LineType = "Heading" And mainworkSheet.Cells(Row, 5).Value = 0 Then
        mainworkSheet.Rows(Row).Delete
        Row = Row - 1
        RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
    ElseIf LineType = "Division" Then
        
        If SumOfHeading > 0 Then
        
        mainworkSheet.Rows(Row).Insert shift:=xlShiftDown
        mainworkSheet.Cells(Row, 1).Value = "Division"
        mainworkSheet.Cells(Row, 2).Value = LastDivisionName
        mainworkSheet.Cells(Row, 5).Value = SumOfHeading
        LastDivisionName = mainworkSheet.Cells(Row + 1, 2).Value
        mainworkSheet.Rows(Row).Insert shift:=xlShiftDown
        RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
        SumOfHeading = 0
        Row = Row + 2
        
        mainworkSheet.Rows(Row).Insert shift:=xlShiftDown
        mainworkSheet.Rows(Row + 1).PageBreak = xlPageBreakManual
        RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
        Else
        
        LastDivisionName = mainworkSheet.Cells(Row, 2).Value
        mainworkSheet.Cells(Row, 5).Value = ""
        mainworkSheet.Rows(Row).Insert shift:=xlShiftDown
        RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
        Row = Row + 1
        End If
    
    ElseIf LineType = "Heading" Then
        SumOfHeading = SumOfHeading + mainworkSheet.Cells(Row, 5).Value
        mainworkSheet.Rows(Row).Insert shift:=xlShiftDown

        Row = Row + 2
        mainworkSheet.Rows(Row).Insert shift:=xlShiftDown
        RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
        Row = Row - 1
    End If
Next
'This is the last line of sum of heading, one off case has to be addressed seperately.

If SumOfHeading > 0 Then
    RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
    mainworkSheet.Cells(RowCount + 1, 1).Value = "Division"
    mainworkSheet.Cells(RowCount + 1, 2).Value = LastDivisionName
    mainworkSheet.Cells(RowCount + 1, 5).Value = SumOfHeading
    mainworkSheet.Rows(RowCount + 1).Insert shift:=xlShiftDown
    
End If

RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
For Row = TargetLine To RowCount
    LineType = mainworkSheet.Cells(Row, 1).Value
    If LineType = "Division" Then
    Call PPBook_DivisionFormat(Row, "PP")
    End If
Next

'All detail estimate lines have been pulled in and filtered to this point
'next need to build and format the cover page

Call PPBook_SummaryPage("PP", TargetLine)

Call PPBook_ColumnFormatting("PP")
mainworkSheet.Columns(1) = ""
Call PPBook_HeaderAndFooter("PP")

'Prtint to PDF
Dim FileSaveName

FileSaveName = Application.GetSaveAsFilename( _
InitialFileName:="Conceptual Estimate Report" + ".pdf", fileFilter:="PDF (*.pdf), *.pdf")
If FileSaveName = False Then
Exit Sub
End If
Worksheets("PP CoverPage").Visible = True
mainworkBook.Sheets(Array("PP", "PP CoverPage")).Select

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    FileSaveName, Quality:=xlQualityStandard, IncludeDocProperties:=True, _
     IgnorePrintAreas:=False, OpenAfterPublish:=True
     
Worksheets("PP CoverPage").Visible = False
Application.CutCopyMode = False
End Sub
