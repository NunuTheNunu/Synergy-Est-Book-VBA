Attribute VB_Name = "VC_CustomMadeFunctions"
'Functions are in this Module

Function Revenue_Year(ByVal This_Year As Variant, ByVal Start_Date As Date, ByVal End_Date As Date, Revenue As Variant) As Variant
If This_Year < Year(Start_Date) Then
Revenue_Year = "Project Not Started"
ElseIf This_Year = Year(Start_Date) Then
Revenue_Year = (13 - Month(Start_Date)) / DateDiff("m", Start_Date, End_Date) * Revenue
ElseIf This_Year < Year(End_Date) Then
Revenue_Year = 12 / DateDiff("m", Start_Date, End_Date) * Revenue
ElseIf This_Year = Year(End_Date) Then
Revenue_Year = (Month(End_Date) - 1) / DateDiff("m", Start_Date, End_Date) * Revenue
ElseIf This_Year > Year(End_Date) Then
Revenue_Year = "Project Finished"
Else
Revenew_year = "How could this happen?"
End If
End Function
'CostType Function

Function CostType(CostRange As Range, HeaderCost As Range)

Dim CostArray, item As Variant
Dim counter As Integer
Dim CostTypeArray As Variant
CostTypeArray = Array("LAB", "MAT", "EQM", "SUB")
counter = 1
ReDim CostArray(UBound(Array(CostRange)))
CostArray = Application.Transpose(CostRange.Value)
For Each item In CostArray
'Debug.Print item
If item > 0 And counter > 1 Then
Exit For
Else
counter = counter + 1
End If
Next
If counter > 4 Then
counter = 1
End If
If HeaderCost.Value = "CostLine" Then
CostType = CostTypeArray(counter - 1)
Else
CostType = ""
End If
'Debug.Print CostType
End Function
Function ContractItem(CostCode As Range, HeaderCost As Range)
ContractItem = (Application.WorksheetFunction.RoundDown(CostCode / 1000, 0)) * 100
'Debug.Print ContractItem
If CostCode > 100 And CostCode < 1000 Then
ContractItem = 100
End If
If HeaderCost <> "CostLine" Then
ContractItem = ""
End If
End Function
Function ContractItemDescription(CostCode As Range, HeaderCost As Range)
On Error Resume Next
Dim ContractItemArray As Variant
ContractItemArray = Array("General Conditions", "Siteworks", "Concrete & Finishes", "Masonry", "Metals", "Wood", "Thermal & Moisture Protection", _
"Doors & Windows", "Finishes", "Specialties", "Equipment", "Furnishings", "General Building Items", "Conveying Systems", "Mechanical", "Electrical", "SUB", "SUB")
ContractItemDescription = ContractItemArray((CostCode \ 1000) - 1)
If CostCode < 1000 Then
ContractItemDescription = "General Conditions"
End If
If HeaderCost <> "CostLine" Then
ContractItemDescription = ""
End If
'Debug.Print ContractItemDescription
End Function

Function ContractItem_BaseContract(CostCode As Range, HeaderCost As Range)
On Error Resume Next
ContractItem_BaseContract = 1
If HeaderCost <> "CostLine" Then
ContractItem_BaseContract = ""
End If
End Function
Function ContractItemDescription_BaseContract(CostCode As Range, HeaderCost As Range)
On Error Resume Next
ContractItemDescription_BaseContract = "Base Contract"
If HeaderCost <> "CostLine" Then
ContractItemDescription_BaseContract = ""
End If
End Function

