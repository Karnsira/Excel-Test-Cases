
Function GetMergedCellValue(rng As Range) As Variant
    Dim cell As Range
    Set cell = rng.Cells(1, 1) ' Always reference the first cell in the range
    If cell.MergeCells Then
        ' If the cell is part of a merged range, get the value from the top-left cell of the merged range
        GetMergedCellValue = cell.MergeArea.Cells(1, 1).Value
    Else
        ' If the cell is not merged, get the value directly
        GetMergedCellValue = cell.Value
    End If
End Function
Function GetMergedCell(rng As Range) As Range
    Dim cell As Range
    Set cell = rng.Cells(1, 1) ' Always reference the first cell in the range
    If cell.MergeCells Then
        ' If the cell is part of a merged range, get the value from the top-left cell of the merged range
        GetMergedCell = cell.MergeArea.Cells
    Else
        ' If the cell is not merged, get the value directly
        GetMergedCell = cell
    End If
End Function
Function validateDbInputPosition(ByVal Target As Range) As Boolean
    Dim ws As Worksheet
    Dim currentCell As Range
    Dim aboveCell As Range
    Dim searchValues As Variant
    Dim matchFound As Boolean
    
    ' Set your worksheet, assuming it's the active one
    Set ws = ActiveSheet
    
    ' Set the current cell, assuming it's the active cell
    Set currentCell = Target
    
    ' Define the list of values to search for
    searchValues = Array("Service", "Type")
    
    ' Check if the current cell is not in the first row
    If currentCell.Row > 1 Then
        If currentCell.Cells.Count <> 1 Then
            Exit Function
        End If
        ' Get the cell above the current cell
        Set aboveCell = currentCell.Offset(-1, 0)
        
        ' Initialize matchFound to False
        matchFound = False
        
        ' Check if the value in the above cell matches any value in the searchValues array
        For Each Value In searchValues
            If aboveCell.Value = Value Then
                matchFound = True
                Exit For
            End If
        Next Value
        
        ' Output the result
        If matchFound Then
            validateDbInputPosition = True
        Else
            validateDbInputPosition = False
            MsgBox "The cell above does not contain 'Service' or 'Type'."
        End If
    Else
        validateDbInputPosition = False
        MsgBox "The current cell is in the first row, there is no cell above it."
    End If
End Function
Function GetLastRow(rng As Range) As Range
    Dim cell As Range
    Set cell = rng.Cells(1, 1) ' Always reference the first cell in the range
    For Each cell In rng
        If IsEmpty(GetMergedCellValue(cell)) Then
            Set GetLastRow = cell
            Exit For
        End If
    Next cell
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long

    Dim isNewService As Boolean
    
    Dim typeRange As Range
    Dim sourceRange As Range
    Dim sourceMergedRange As Range
    
    Dim tableInput As Object
    Dim serviceCell As Object
    Dim serviceInput As Object
    
    Dim dictChangesCell As New Dictionary

    Application.DisplayAlerts = False

    Set ws = Me

    'Define the range to monitor for changes (C and D columns)
    'If Not Intersect(Target, ws.Range("C2:D2")) Is Nothing Then
    If Not Intersect(Target, ws.Range("C:D")) Is Nothing Then
        ' Validate Database Input Position
        If Not validateDbInputPosition(Target) Then
            Exit Sub
        End If
        
        Dim NextFree As Range
        Dim currentRow As String
        Dim currentColumn As String
        Dim activeToLast As String
        Dim tempReplace As String
        
        currentRow = CStr(Target.Row)
        tempReplace = Replace(Target.Address(False, False), Target.Row, "")
        currentColumn = Replace(tempReplace, ":", "")
        activeToLast = Target.Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & currentColumn & 100
        
        'Find last row of current state
        Set NextFree = GetLastRow(Range(activeToLast))
        lastRow = NextFree.Offset(-1).Row
        
        'Set input cells
        Set serviceInput = ws.Cells(currentRow, "C")
        Set tableInput = ws.Cells(currentRow, "D")
        
        If Not IsEmpty(serviceInput.Value) And Not IsEmpty(tableInput.Value) Then
            'Loop through the table and build a dictionary
            For i = currentRow + 1 To lastRow
                Set serviceCell = ws.Cells(i, "C")
                serviceCellName = serviceCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                currentService = GetMergedCellValue(Range(serviceCellName))
                
                If Not dictChangesCell.Exists(currentService) Then
                    dictChangesCell(currentService) = serviceCellName
                End If
            Next i
            
            Application.EnableEvents = False

            isNewService = Not dictChangesCell.Exists(serviceInput.Value)
            If isNewService Then
                dictChangesCell(serviceInput.Value) = ws.Cells(lastRow, 3).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            End If

            Set sourceRange = Range(CStr(dictChangesCell(CStr(serviceInput))))
            If sourceRange.MergeCells Then
                Set sourceMergedRange = sourceRange.MergeArea
            Else
                Set sourceMergedRange = sourceRange
            End If
                
            'Insert row below changes cell
            sourceRange.Offset(1).Resize(1, 2).Insert
            
            'Fill value in new row with input
            sourceRange.Offset(1).Value = serviceInput
            sourceRange.Offset(1, 1).Value = tableInput
            
            'Merge main service & sub service
            If Not isNewService Then
                Range(sourceMergedRange, sourceMergedRange.Offset(1)).Merge
            End If
            
            'Decorate cells
            If isNewService Then
                sourceRange.Offset(1).Borders.Weight = xlThick
                sourceRange.Offset(1, 1).Borders.Weight = xlThick
            Else
                sourceRange.MergeArea.Borders.Weight = xlThick
                
                Dim lastIndex As Long
                Dim currentRight As Range
                Dim currentRange As Range
                lastIndex = sourceMergedRange.Rows.Count
                Set currentRight = sourceRange.Offset(0, 1)
                Set currentRange = Range(currentRight, currentRight.Offset(lastIndex))
                currentRange.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
                currentRange.BorderAround Weight:=xlThick
                
                'sourceRange.Offset(0, 1).Borders.Weight = xlThick
            End If
            
            
            Application.EnableEvents = True
            
        End If
    End If
End Sub
Sub CopyRangeAndPasteAfterSkippingRows()
    Dim ws As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim lastRow As Long
    Dim skipRows As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet3") ' Change "Sheet1" to your sheet name
    Set sourceRange = ws.Range("C1:D2") ' Change this to your fixed source range
    
    skipRows = 5
    
    ' Find the last row of the source range
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    Set MyMergedRange = ws.Cells(ws.Rows.Count, "C").End(xlUp).MergeArea
    FirstRow = MyMergedRange.Row
    lastRow = MyMergedRange.Row + MyMergedRange.Rows.Count - 1
    
    
    ' Set the target range 5 rows below the last row of the source range
    Set targetRange = ws.Cells(lastRow + skipRows + 1, sourceRange.Columns(1).Column)
    
    ' Copy the source range
    sourceRange.Copy
    
    ' Paste the source range to the target range
    targetRange.PasteSpecial Paste:=xlPasteAll
    
    ' Clear the clipboard to remove the copied range
    Application.CutCopyMode = False
End Sub



Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Private Sub Worksheet_TableUpdate(ByVal Target As TableObject)

End Sub
