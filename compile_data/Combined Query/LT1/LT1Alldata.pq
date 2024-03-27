let
    // Load Query CityU and headers
    ImportFromCityU = LT1FromCityU,
    CityUHeaders = Table.ColumnNames(ImportFromCityU),

    // Load Query HKU and headers
    ImportFromHKU = LT1FromHKU,
    HKUHeaders = Table.ColumnNames(ImportFromHKU),

    // Load Query EdU and headers
    ImportFromEdU = LT1FromEdU,
    EdUHeaders = Table.ColumnNames(ImportFromEdU),

    // Combine Query HKU and EdU based on header name
    TableHKUEdU = Table.Combine({ImportFromHKU, ImportFromEdU}),
    CombinedHeaders = Table.ColumnNames(TableHKUEdU),

    // Ensure it has the same number of columns as Query CityU
    CombinedColumnCount = List.Count(CombinedHeaders),
    CityUColumnCount = List.Count(CityUHeaders),

    // If the number of columns match, rename and append
    LT1CombinedQuery = if CombinedColumnCount = CityUColumnCount then
        let
            // Append CityU to RenamedHKUEdU
            AppendedTables = Table.Combine({ImportFromCityU, TableHKUEdU})
        in
            AppendedTables
    else
        error "The number of columns in the combined HKU and EdU table does not match the number of columns in 'Import from CityU'.",

    // Add Index Column to the renamed table
    SortedRows = Table.Sort(LT1CombinedQuery,{{"LT1_Date_entry", Order.Ascending}}), // Make sure to use the correct new column name after renaming

    // Append suffix "_LT1" to every cell in the "ST1_NO" column
    SuffixedLT1_NO = Table.TransformColumns(
        SortedRows,
        {"LT1_NO", each _ & "_LT1", type text}
    )
in
    SuffixedLT1_NO