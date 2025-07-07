Sub FindAndReplaceArray()
    Dim dict As Object
    Dim mapSheet As Worksheet
    Dim keyCell As Range
    Dim targetRange As Range
    Dim dataArray As Variant
    Dim i As Long, j As Long
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
        If Not dict.exists(keyCell.Value) Then
            dict.Add keyCell.Value, keyCell.Offset(0, 1).Value
        End If
    Next keyCell
    
    ' Set the target range on the active sheet (change the range as necessary)
    Set targetRange = ActiveSheet.UsedRange ' Changed to UsedRange
    
    ' Read the target range into an array for fast processing
    If Not targetRange Is Nothing Then
        dataArray = targetRange.Value
        
        ' Loop through each cell in the array
        For i = 1 To UBound(dataArray, 1)
            For j = 1 To UBound(dataArray, 2)
                ' Perform find and replace from the dictionary
                For Each key In dict.Keys
                    If InStr(dataArray(i, j), key) > 0 Then
                        dataArray(i, j) = Replace(dataArray(i, j), key, dict(key))
                    End If
                Next key
            Next j
        Next i
        
        ' Write the updated array back to the target range
        targetRange.Value = dataArray
    End If
    
    ' Turn on screen updating and set calculation back to automatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub
