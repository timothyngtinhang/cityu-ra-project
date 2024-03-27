let
    // Reference the FolderPathsQuery
    FolderPathsQuery = #"FolderPathsQuery",
    // Retrieve the folder path from the table
    // Make sure to use the correct row index (0 for the first row)
    FolderPath = FolderPathsQuery{1}[FolderPath],
    // Get the files from the specified folder
    Source = Folder.Files(FolderPath),

    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Name], "CheckCurrent")),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Custom", each let
        // Access the content of the file
        ExcelContent = Excel.Workbook([Content]),

        // Select the specific sheet by name
        SpecificSheet = Table.SelectRows(ExcelContent, each [Name] = "SPSS Data Entry Template" and [Kind] = "Sheet"),
        FirstSheetData = SpecificSheet{0}[Data],
        
        // Skip the first row
        SkippedFirstRow = Table.Skip(FirstSheetData, 1),

        // Promote the first row to headers
        PromotedHeaders = Table.PromoteHeaders(SkippedFirstRow, [PromoteAllScalars=true]),
        
        // Create a list of non-empty header names
        NonEmptyHeaderNames = List.Select(Table.ColumnNames(PromotedHeaders), each _ <> "" and _ <> null and not Text.StartsWith(_, "Column")),
        
        // Make sure the table includes the 'LT1_Date_entry' column
        // If 'LT1_Date_entry' is not in the headers, this will throw an error
        LT1DateEntryColumn = List.First(List.Select(NonEmptyHeaderNames, each Text.Contains(_, "LT1_Date_entry"))),

        // Select only columns with non-empty headers
        CleanedTable = Table.SelectColumns(PromotedHeaders, NonEmptyHeaderNames),

        // Filter out rows containing 'LT1_Date_entry'
        FilteredForLT1DateEntry = Table.SelectRows(CleanedTable, each Record.Field(_, LT1DateEntryColumn) <> null 
            and Record.Field(_, LT1DateEntryColumn) <> "LT1_Date_entry"),

        // Remove rows where fewer than 3 cells are non-empty
        NonEmptyRows = Table.SelectRows(FilteredForLT1DateEntry, each List.NonNullCount(Record.FieldValues(_)) >= 3),

        // Get the schema of the table to determine column types
        Schema = Table.Schema(NonEmptyRows),

        // Filter the schema to only include record columns
        RecordColumns = Table.SelectRows(Schema, each [Kind] = "record"),

        // Get a list of record column names
        RecordColumnNames = RecordColumns[Name],

        // Expand the record columns if any are present
        ExpandedTable = List.Accumulate(RecordColumnNames, NonEmptyRows, (currentTable, columnName) =>
            Table.ExpandRecordColumn(currentTable, columnName, Record.FieldNames(currentTable{0}[columnName]))
        )
    in
        ExpandedTable),
    // Add a step to expand the nested tables
    #"Expanded Custom" = if Table.IsEmpty(#"Added Custom") then #"Added Custom" else Table.ExpandTableColumn(#"Added Custom", "Custom", Table.ColumnNames(#"Added Custom"[Custom]{0})),
    #"Sorted Rows" = Table.Sort(#"Expanded Custom",{{"LT1_Date_entry", Order.Ascending}}),
    #"Removed Columns" = Table.RemoveColumns(#"Sorted Rows",{"Content", "Extension", "Date modified", "Date created", "Attributes", "Duplicates?", "Check_Date"})
in
    #"Removed Columns"