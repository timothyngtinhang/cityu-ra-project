Sub SetValidationByText()
    Dim dataRange As Range
    Dim validationRange As Range
    Dim i As Long
    
    ' Turn off screen updating to speed up the macro.
    Application.ScreenUpdating = False
    
    ' Define the range where you want to apply the data validation
    Set dataRange = ThisWorkbook.Sheets("Validation VBA").Range("C2:C" & Rows.Count).End(xlUp)) ' Changed to dynamic range
    
    ' Define the range that contains the validation criteria for each cell
    Set validationRange = ThisWorkbook.Sheets("Validation VBA").Range("B2:B" & Rows.Count).End(xlUp)) ' Changed to dynamic range
    
    ' Clear all data validation from the target range before applying new validation
    dataRange.Validation.Delete
    
    ' Loop through each cell in the data range and set the validation
    For i = 1 To dataRange.Cells.Count
        ' Only attempt to add validation if the cell is not empty
        If Len(Trim(validationRange.Cells(i).Value)) > 0 Then
            With dataRange.Cells(i).Validation
                .Delete ' Clear any existing validation rules
                On Error Resume Next ' Skip any errors and proceed
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                      xlBetween, Formula1:=validationRange.Cells(i).Value
                If Err.Number <> 0 Then
                    ' If there was an error, clear the error
                    Err.Clear
                Else
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                    .ErrorTitle = "Invalid Input"
                    .ErrorMessage = "Please enter a value from the list."
                End If
                ' Reset error handling to default behavior
                On Error GoTo 0
            End With
        End If
    Next i
    
    ' Turn on screen updating again.
    Application.ScreenUpdating = True
End Sub

