Sub OptimizedFindAndReplaceTable()
    Dim dict As Object
    Dim mapSheet As Worksheet
    Dim keyCell As Range
    Dim targetTable As ListObject
    Dim tableColumn As ListColumn
    Dim cell As Range
    Dim key As Variant
    
    ' Turn off screen updating and set calculation to manual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize the dictionary object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Set the mapping worksheet
    Set mapSheet = ThisWorkbook.Worksheets("Mapping")
    
    ' Add find/replace pairs to the dictionary
    For Each keyCell In mapSheet.Range("A1:A" & mapSheet.Cells(mapSheet.Rows.Count, "A").End(xlUp).Row)
        If Not dict.exists(keyCell.value) Then
            dict.Add keyCell.value, keyCell.Offset(0, 1).value
        End If
    Next keyCell
    
    ' Set the target table on the active sheet (change the table name as necessary)
    Set targetTable = ActiveSheet.ListObjects(1) ' Changed to reference the first table
    
    ' Loop through each column in the table
    For Each tableColumn In targetTable.ListColumns
        ' Loop through each cell in the column
        For Each cell In tableColumn.DataBodyRange
            ' Check and replace structured references in the formula
            If cell.HasFormula Then
                Dim formula As String
                formula = cell.formula
                For Each key In dict.Keys
                    If InStr(formula, key) > 0 Then
                        formula = Replace(formula, key, dict(key))
                    End If
                Next key
                cell.formula = formula
            Else
                ' Or handle the cell value
                Dim value As String
                value = cell.value
                For Each key In dict.Keys
                    If InStr(value, key) > 0 Then
                        value = Replace(value, key, dict(key))
                    End If
                Next key
                cell.value = value
            End If
        Next cell
    Next tableColumn
    
    ' Turn on screen updating and set calculation back to automatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub
