Attribute VB_Name = "GetSheetName_VC"

Sub FnGetSheetsName()

Dim mainworkBook As Workbook
Dim AddToRecoverableToggle As Boolean
Dim AllSheetNamesCount, RecoverableSheetNamesCount As Integer
On Error Resume Next
AllSheetNamesCount = 0
RecoverableSheetNamesCount = 0
'Dim AllSheetNames
'2018-06-22 AllSheetNames Dim'd as Public
ReDim AllSheetNames(0), RecoverableSheetNames(0)
Set mainworkBook = ActiveWorkbook
For i = 1 To mainworkBook.Sheets.Count
    
    If mainworkBook.Sheets(i).Name = "1" Then
    AddToRecoverableToggle = True
    End If
    If mainworkBook.Sheets(i - 1).Name = "15-16" Then
    AddToRecoverableToggle = False
    End If
    If AddToRecoverableToggle = True Then
    RecoverableSheetNames(RecoverableSheetNamesCount) = mainworkBook.Sheets(i).Name
    RecoverableSheetNamesCount = RecoverableSheetNamesCount + 1
    ReDim Preserve RecoverableSheetNames(RecoverableSheetNamesCount)
    Else
    AllSheetNames(AllSheetNamesCount) = mainworkBook.Sheets(i).Name
    AllSheetNamesCount = AllSheetNamesCount + 1
    ReDim Preserve AllSheetNames(AllSheetNamesCount)
    'Either we can put all names in an array , here we are printing all the names in Sheet 2
    'mainworkBook.Sheets("Sheet3").Range("A" & i) = AllSheetNames(i - 1)
    'mainworkBook.Sheets("Sheet2").Range("A" & i) = mainworkBook.Sheets(i).Name
    End If
Next i
ReDim Preserve RecoverableSheetNames(RecoverableSheetNamesCount - 1)
ReDim Preserve AllSheetNames(AllSheetNamesCount - 1)
End Sub
