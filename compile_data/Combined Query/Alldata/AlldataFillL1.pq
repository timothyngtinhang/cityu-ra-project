let
    // Define the special numbers within the main let block
    specialNums = {6666, 7777, 8888, 9999},

    // Function to check the cell value
    CheckValue = (cellValue as any) =>
        let
            // Convert cellValue to number if it's a text representation of a number
            numberCellValue = try Number.From(cellValue) otherwise cellValue,
            // Check if the cellValue (or its numeric equivalent) is in the special numbers list
            specialCheck = List.Contains(specialNums, numberCellValue),
            result = if cellValue = null or cellValue = "#blank" or cellValue = "" then "Unfilled"
                    else if specialCheck then "Unfilled"
                    else "Filled"
        in
            result,

    // Placeholders for other steps in your query (e.g., data fetching, filtering, etc.)

    // Reference to the AllData query which should be already loaded

    ColumnNames = Table.ColumnNames(Alldata),

    // Apply the CheckValue function to each cell in each column (except the index column)
    TransformedTable = List.Accumulate(ColumnNames, Alldata, (state, currentColumn) => 
        if currentColumn = "Index" then
            state
        else
            let
                UpdatedColumn = Table.TransformColumns(state, {{currentColumn, each CheckValue(_), type text}})
            in
                UpdatedColumn
    )
in
    // Ensure the final step of your query is the appended data
    TransformedTable