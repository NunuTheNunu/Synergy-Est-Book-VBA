Attribute VB_Name = "VC_CSVFileCodes"
'---------------------------------------------------------------------------
'!!!DO NOT CHANGE CODE IN THIS PAGE UNLESS VIEWPOINT TEMPLATE HAS CHANGED!!!
'---------------------------------------------------------------------------

Sub ContractItemSeleteUnique()
Dim mainworkBook As Workbook
Dim mainworkSheet As Worksheet
Dim RowCount, RowCounter, i As Integer
'Dim ContractItemArray As Variant
Dim test
Dim IsInArray As Boolean
Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.Sheets("Consolidation")
RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
ReDim ContractItemArray(7, RowCount - 1)
ContractItemCount = 0
For RowCounter = 1 To RowCount
    If mainworkSheet.Cells(RowCounter, 1) = "CostLine" Then
        test = mainworkSheet.Cells(RowCounter, 2).Value
        IsInArray = False
    
        For i = 1 To RowCount
            If ContractItemArray(2, i - 1) = test Then
                IsInArray = True
                Exit For
            End If
        Next
    
        If IsInArray = False Then
        'Record Type
        
            ContractItemArray(0, ContractItemCount) = 1
            ContractItemArray(2, ContractItemCount) = mainworkSheet.Cells(RowCounter, 2).Value
            ContractItemArray(5, ContractItemCount) = mainworkSheet.Cells(RowCounter, 3).Value
            ContractItemArray(6, ContractItemCount) = mainworkSheet.Cells(RowCounter, 8).Value
            ContractItemArray(7, ContractItemCount) = "LS"
            ContractItemCount = ContractItemCount + 1
        End If
    End If
Next
ReDim Preserve ContractItemArray(7, ContractItemCount - 1)
End Sub

Sub PhaseCodeSeleteUnique()
Dim mainworkBook As Workbook
Dim mainworkSheet As Worksheet
Dim RowCount, RowCounter, i As Integer
Dim test
Dim IsInArray As Boolean
Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.Sheets("Consolidation")
RowCount = mainworkSheet.Cells(Rows.Count, 2).End(xlUp).Row
ReDim PhaseCodeArray(5, RowCount - 1)
PhaseCodeCount = 0
For RowCounter = 1 To RowCount
    If mainworkSheet.Cells(RowCounter, 1) = "CostLine" Then
        test = mainworkSheet.Cells(RowCounter, 5).Value
        IsInArray = False
        
        For i = 1 To RowCount
            If PhaseCodeArray(3, i - 1) = test Then
                IsInArray = True
                Exit For
            End If
        Next
        
        If IsInArray = False Then
        'Phase Code
        
            PhaseCodeArray(0, PhaseCodeCount) = 2
            PhaseCodeArray(2, PhaseCodeCount) = mainworkSheet.Cells(RowCounter, 2).Value
            PhaseCodeArray(3, PhaseCodeCount) = mainworkSheet.Cells(RowCounter, 5).Value
            PhaseCodeArray(4, PhaseCodeCount) = mainworkSheet.Cells(RowCounter, 6).Value
            PhaseCodeCount = PhaseCodeCount + 1
        ElseIf IsInArray = True And (mainworkSheet.Cells(RowCounter, 4) = "P" Or mainworkSheet.Cells(RowCounter, 4) = "p") Then
            PhaseCodeArray(4, i - 1) = mainworkSheet.Cells(RowCounter, 6).Value
        End If
    End If
Next
ReDim Preserve PhaseCodeArray(5, PhaseCodeCount - 1)
End Sub

Sub CostItemSelectUnique()
Dim mainworkBook As Workbook
Dim mainworkSheet As Worksheet
Dim RowCount, RowCounter, i, j As Integer
Dim TestCostCode, TestCostType
Dim CostCodeIsInArray, CostTypeIsInArray As Boolean
Set mainworkBook = ActiveWorkbook
Set mainworkSheet = mainworkBook.Sheets("Consolidation")
RowCount = mainworkSheet.Cells(Rows.Count, 5).End(xlUp).Row
ReDim CostItemArray(8, RowCount - 1)
TestCostType = 0
CostItemCount = 0
'I decided all codes with labour cost should be tracked seperately, however it will not shown in consolidation sheet
For RowCounter = 1 To RowCount
    If mainworkSheet.Cells(RowCounter, 1) = "CostLine" Then
        TestCostCode = mainworkSheet.Cells(RowCounter, 5).Value
        If mainworkSheet.Cells(RowCounter, 11).Value > 0 Then
            TestCostType = 1
            Else
            TestCostType = 0
        End If
        CostCodeIsInArray = False
        CostTypeIsInArray = False ' cost type can only be true when cost code is true
        
            For i = 1 To RowCount
                If CostItemArray(3, i - 1) = TestCostCode Then
                    CostCodeIsInArray = True
                    Exit For
                End If
            Next
            If CostCodeIsInArray = True Then
            For j = 1 To RowCount
                If CostItemArray(3, j - 1) = TestCostCode And CostItemArray(4, j - 1) = TestCostType Then
                    CostTypeIsInArray = True
                    Exit For
                End If
            Next
            End If
            If CostTypeIsInArray = False And TestCostType = 1 Then
            'Phase Code
            
                CostItemArray(0, CostItemCount) = 3
                CostItemArray(2, CostItemCount) = mainworkSheet.Cells(RowCounter, 2).Value
                CostItemArray(3, CostItemCount) = mainworkSheet.Cells(RowCounter, 5).Value
                
                CostItemArray(4, CostItemCount) = 1 'means this is labour code
                
                If mainworkSheet.Cells(RowCounter, 7).Value = "LAB" Then
                CostItemArray(5, CostItemCount) = mainworkSheet.Cells(RowCounter, 8).Value
                CostItemArray(6, CostItemCount) = mainworkSheet.Cells(RowCounter, 9).Value
                CostItemArray(7, CostItemCount) = mainworkSheet.Cells(RowCounter, 10).Value
                CostItemArray(8, CostItemCount) = mainworkSheet.Cells(RowCounter, 11).Value
                Else
                CostItemArray(5, CostItemCount) = mainworkSheet.Cells(RowCounter, 10).Value
                CostItemArray(6, CostItemCount) = "MHR"
                CostItemArray(7, CostItemCount) = mainworkSheet.Cells(RowCounter, 10).Value
                CostItemArray(8, CostItemCount) = mainworkSheet.Cells(RowCounter, 11).Value
                End If
                CostItemCount = CostItemCount + 1
            ElseIf CostTypeIsInArray = True And TestCostType = 1 Then
                CostItemArray(8, j - 1) = CostItemArray(8, j - 1) + mainworkSheet.Cells(RowCounter, 11).Value
                'CostItemArray(5, j - 1) = CostItemArray(5, j - 1) + mainworkSheet.Cells(RowCounter, 10).Value
                CostItemArray(7, j - 1) = CostItemArray(7, j - 1) + mainworkSheet.Cells(RowCounter, 10).Value
                'All hours and manhours will get rolled into the first line the duplicated lines
            End If
            
    End If
Next
'all labour lines are poured in
ReDim Preserve CostItemArray(8, CostItemCount - 1 + RowCount)

'now the material cost
For RowCounter = 1 To RowCount
    If mainworkSheet.Cells(RowCounter, 1) = "CostLine" Then
        TestCostCode = mainworkSheet.Cells(RowCounter, 5).Value
        If mainworkSheet.Cells(RowCounter, 12).Value > 0 Then
            'if material cost exist, then run material cost to csv
            TestCostType = 3
            Else
            GoTo NO_MATERIAL_COST
        End If
        CostCodeIsInArray = False
        CostTypeIsInArray = False ' cost type can only be true when cost code is true
        
            For i = 1 To RowCount
                If CostItemArray(3, i - 1) = TestCostCode Then
                    CostCodeIsInArray = True
                    Exit For
                End If
            Next
            If CostCodeIsInArray = True Then
            For j = 1 To RowCount
                If CostItemArray(3, j - 1) = TestCostCode And CostItemArray(4, j - 1) = TestCostType Then
                    CostTypeIsInArray = True
                    Exit For
                End If
            Next
            End If
            If CostTypeIsInArray = False And TestCostType <> 1 Then
            'Phase cost type not labour and not in array
            
                CostItemArray(0, CostItemCount) = 3
                CostItemArray(2, CostItemCount) = mainworkSheet.Cells(RowCounter, 2).Value
                CostItemArray(3, CostItemCount) = mainworkSheet.Cells(RowCounter, 5).Value
                CostItemArray(4, CostItemCount) = TestCostType
                CostItemArray(5, CostItemCount) = mainworkSheet.Cells(RowCounter, 8).Value
                CostItemArray(6, CostItemCount) = mainworkSheet.Cells(RowCounter, 9).Value
                
                'CostItemArray(7, CostItemCount) = mainworkSheet.Cells(RowCounter, 10).Value
                CostItemArray(7, CostItemCount) = 0
                
                CostItemArray(8, CostItemCount) = mainworkSheet.Cells(RowCounter, 12).Value
                CostItemCount = CostItemCount + 1
            ElseIf CostTypeIsInArray = True And TestCostType <> 1 Then
                CostItemArray(8, j - 1) = CostItemArray(8, j - 1) + mainworkSheet.Cells(RowCounter, 12).Value
                'CostItemArray(5, j - 1) = CostItemArray(5, j - 1) + mainworkSheet.Cells(RowCounter, 10).Value
            End If
            
    End If

'jump to here if no material cost
NO_MATERIAL_COST:
Next
ReDim Preserve CostItemArray(8, CostItemCount - 1 + RowCount)

'now the subcontract cost
For RowCounter = 1 To RowCount
    If mainworkSheet.Cells(RowCounter, 1) = "CostLine" Then
        TestCostCode = mainworkSheet.Cells(RowCounter, 5).Value
        If mainworkSheet.Cells(RowCounter, 14).Value > 0 Then
            'if subcontract cost exist, then run subcontract cost to csv
            TestCostType = 2
            Else
            GoTo NO_SUBCONTRACT_COST
        End If
        CostCodeIsInArray = False
        CostTypeIsInArray = False ' cost type can only be true when cost code is true
        
            For i = 1 To RowCount
                If CostItemArray(3, i - 1) = TestCostCode Then
                    CostCodeIsInArray = True
                    Exit For
                End If
            Next
            If CostCodeIsInArray = True Then
            For j = 1 To RowCount
                If CostItemArray(3, j - 1) = TestCostCode And CostItemArray(4, j - 1) = TestCostType Then
                    CostTypeIsInArray = True
                    Exit For
                End If
            Next
            End If
            If CostTypeIsInArray = False And TestCostType <> 1 Then
            'Phase cost type not labour and not in array
            
                CostItemArray(0, CostItemCount) = 3
                CostItemArray(2, CostItemCount) = mainworkSheet.Cells(RowCounter, 2).Value
                CostItemArray(3, CostItemCount) = mainworkSheet.Cells(RowCounter, 5).Value
                CostItemArray(4, CostItemCount) = TestCostType
                CostItemArray(5, CostItemCount) = mainworkSheet.Cells(RowCounter, 8).Value
                CostItemArray(6, CostItemCount) = mainworkSheet.Cells(RowCounter, 9).Value
                CostItemArray(7, CostItemCount) = 0
                
                CostItemArray(8, CostItemCount) = mainworkSheet.Cells(RowCounter, 14).Value
                CostItemCount = CostItemCount + 1
            ElseIf CostTypeIsInArray = True And TestCostType <> 1 Then
                CostItemArray(8, j - 1) = CostItemArray(8, j - 1) + mainworkSheet.Cells(RowCounter, 14).Value
            End If
            
    End If

'jump to here if no subcontract cost
NO_SUBCONTRACT_COST:
Next
ReDim Preserve CostItemArray(8, CostItemCount - 1)


End Sub


