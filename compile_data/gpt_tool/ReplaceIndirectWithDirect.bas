Sub ReplaceIndirectWithDirect()
    Dim cell As Range
    Dim formula As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim indirectRef As String
    Dim directRef As String

    ' Turn off screen updating and set calculation to manual
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Define the range to search for INDIRECT formulas
    Dim targetRange As Range
    Set targetRange = ActiveSheet.Range("B30:B542") ' Adjust the range to your needs

    ' Loop through each cell in the defined range
    For Each cell In targetRange
        If cell.HasFormula Then
            formula = cell.formula
            startPos = InStr(formula, "INDIRECT(""")
            While startPos > 0
                endPos = InStr(startPos, formula, """)")
                indirectRef = Mid(formula, startPos, endPos - startPos + 2)
                directRef = Replace(indirectRef, "INDIRECT(""", "")
                directRef = Replace(directRef, """)", "")
                formula = Replace(formula, indirectRef, directRef)
                startPos = InStr(formula, "INDIRECT(""")
            Wend
            cell.formula = formula
        End If
    Next cell

    ' Restore screen updating and calculation settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
