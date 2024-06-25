Sub RemoveRowsWithNumbersInString()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim i As Long
    Dim col As String

    ' Set the worksheet and column to check
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change Sheet1 to your sheet name
    col = "B" ' Column to check for numbers

    ' Find the last row in the column
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    ' Loop through each cell in the column from the bottom to the top
    For i = lastRow To 1 Step -1
        Set cell = ws.Cells(i, col)
        If ContainsNumber(cell.Value) Then
            cell.EntireRow.Delete
        End If
    Next i
End Sub

Function ContainsNumber(text As String) As Boolean
    Dim i As Long
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            ContainsNumber = True
            Exit Function
        End If
    Next i
    ContainsNumber = False
End Function
