Sub writeIndexOfSheetInColumn()

For Each cell In ThisWorkbook.ActiveSheet.UsedRange
    If cell.Row <> 1 Then
        Cells(cell.Row, 6) = ThisWorkbook.ActiveSheet.Index
    End If
Next
End Sub
