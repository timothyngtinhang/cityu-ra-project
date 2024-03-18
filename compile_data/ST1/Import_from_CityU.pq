let
    Source = FolderPath,
    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Name], "CheckCurrent")),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Custom", each let
        // Access the content of the file
        ExcelContent = Excel.Workbook([Content]),

        // Select the specific sheet by name
        SpecificSheet = Table.SelectRows(ExcelContent, each [Name] = "SPSS Data Entry Template" and [Kind] = "Sheet"),
        FirstSheetData = SpecificSheet{0}[Data],

        // Promote the first row to headers
        PromotedHeaders = Table.PromoteHeaders(FirstSheetData, [PromoteAllScalars=true]),
        
        // Create a list of non-empty header names
        NonEmptyHeaderNames = List.Select(Table.ColumnNames(PromotedHeaders), each _ <> "" and _ <> null and not Text.StartsWith(_, "Column")),
        
        // Make sure the table includes the 'ST1_Date_entry' column
        // If 'ST1_Date_entry' is not in the headers, this will throw an error
        ST1DateEntryColumn = List.First(List.Select(NonEmptyHeaderNames, each Text.Contains(_, "ST1_Date_entry"))),

        // Select only columns with non-empty headers
        CleanedTable = Table.SelectColumns(PromotedHeaders, NonEmptyHeaderNames),

        // Filter out rows containing 'ST1_Date_entry'
        FilteredForST1DateEntry = Table.SelectRows(CleanedTable, each Record.Field(_, ST1DateEntryColumn) <> null 
            and Record.Field(_, ST1DateEntryColumn) <> "ST1_Date_entry"),

        // Remove rows where fewer than 3 cells are non-empty
        NonEmptyRows = Table.SelectRows(FilteredForST1DateEntry, each List.NonNullCount(Record.FieldValues(_)) >= 3),

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
    #"Sorted Rows" = Table.Sort(#"Expanded Custom",{{"ST1_Date_entry", Order.Ascending}}),
    #"Removed Columns" = Table.RemoveColumns(#"Sorted Rows",{"Content", "Extension", "Date modified", "Date created", "Attributes", "Check_Date", "Check_No", "Error_Items"})
in
    #"Removed Columns"