Attribute VB_Name = "AddDeleteSheet_VC"

Sub DeleteConsolidationSheet(x)
'2018-08-16 i guess i really shouldnt name this function this way
'now it is fucking confusing.
Application.DisplayAlerts = False
On Error Resume Next
Worksheets(x).Delete
Application.DisplayAlerts = True
End Sub

Sub AddConsolidationSheet(x)
'and this one too
ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = x
End Sub

Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
