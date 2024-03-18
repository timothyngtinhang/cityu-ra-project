let
    // Load Query CityU and headers
    ImportFromCityU = #"Import from CityU",
    CityUHeaders = Table.ColumnNames(ImportFromCityU),

    // Load Query HKU and headers
    ImportFromHKU = #"Import from HKU",
    HKUHeaders = Table.ColumnNames(ImportFromHKU),

    // Load Query EdU and headers
    ImportFromEdU = #"Import from EdU",
    EdUHeaders = Table.ColumnNames(ImportFromEdU),

    // Combine Query HKU and EdU based on header name
    TableHKUEdU = Table.Combine({ImportFromHKU, ImportFromEdU}),
    CombinedHeaders = Table.ColumnNames(TableHKUEdU),

    // Ensure it has the same number of columns as Query CityU
    CombinedColumnCount = List.Count(CombinedHeaders),
    CityUColumnCount = List.Count(CityUHeaders),

    // If the number of columns match, rename and append
    ST1CombinedQuery = if CombinedColumnCount = CityUColumnCount then
        let
            // Use the RenameColumnsByPosition function to rename headers in Query HKUEdU
            RenamedHKUEdU = RenameColumnsByPosition(TableHKUEdU, CityUHeaders),
            
            // Append CityU to RenamedHKUEdU
            AppendedTables = Table.Combine({ImportFromCityU, RenamedHKUEdU}),
            
            // Replace blank or null cells with "#blank"
            ReplaceBlanks = Table.TransformColumns(
                AppendedTables,
                List.Transform(Table.ColumnNames(AppendedTables), (columnName) =>
                    {columnName, each if _ is null or _ = "" then "#blank" else _})
            ),
            
        // Apply the date transformation to the specified columns
        DateColumns = {"Date accessed", "ST1_Date_entry", "ST1_Date", "ST1 Part 1A Q4", "ST1 Part 1B Q20.1", "ST1 Part 1B Q27.1"}, // List of date columns
        ReplaceBlanksAndTransformDates = (value as any) as any =>
            let
                textValue = Text.From(value),
                result = if List.Contains({"6666", "7777", "8888", "9999"}, textValue) then textValue
                        else if value is null or value = "" then "#blank"
                        else try DateTime.ToText(DateTime.From(value), "dd MMM yyyy") otherwise "#error"
            in
                result,
        // Create a list of transformations for the specified date columns
        Transformations = List.Transform(DateColumns, (columnName) => 
            {columnName, each ReplaceBlanksAndTransformDates(_), type text}
        ),
        // Apply the transformations to the table
        TransformedTable = Table.TransformColumns(ReplaceBlanks, Transformations)
    in
        TransformedTable
    else
        error "The number of columns in the combined HKU and EdU table does not match the number of columns in 'Import from CityU'.",

    // Add Index Column to the final table
    #"Added Index" = Table.AddIndexColumn(ST1CombinedQuery, "Index", 1, 1, Int64.Type)
in
    #"Added Index" // Removed the #"Changed Type" step since it is not valid after the previous transformations.