Sub ExtractAndDisplayValidation()
    Dim rng As Range
    Dim cell As Range
    Dim dvList As Variant
    Dim ws As Worksheet

    ' Set reference to the appropriate worksheet
    ' Replace "Sheet1" with the name of your sheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Define the range you want to check for data validation in column B
    ' Adjust the row numbers as necessary
    Set rng = ws.Range("B1:B1000" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)

    ' Loop through each cell in the range
    For Each cell In rng
        dvList = "" ' Initialize dvList as empty for each cell
        ' Check if the cell has data validation applied
        If Not cell.Validation Is Nothing Then
            ' Use error handling in case the cell doesn't have a list-type validation
            On Error Resume Next
            ' Check if the validation type is a list
            If cell.Validation.Type = 3 Then ' 3 = xlValidateList
                ' Get the validation formula or list
                dvList = cell.Validation.Formula1
            End If
            ' Resume normal error handling
            On Error GoTo 0
        End If
        ' Write the validation formula or an empty string to the adjacent cell
        cell.Offset(0, 1).Value = dvList
    Next cell
End Sub
