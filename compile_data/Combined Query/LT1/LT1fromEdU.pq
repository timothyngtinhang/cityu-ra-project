let
    // Reference the FolderPathsQuery
    FolderPathsQuery = #"FolderPathsQuery",
    // Retrieve the folder path from the table
    // Make sure to use the correct row index (0 for the first row)
    FolderPath = FolderPathsQuery{1}[FolderPath],
    // Get the files from the specified folder
    Source = Folder.Files(FolderPath),
    
    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Name], "LT1-EdUHK")),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Custom", each let
        // Access the content of the file
        ExcelContent = Excel.Workbook([Content]),
        // Select the first sheet
        FirstSheetData = ExcelContent{0}[Data],
        // Promote the first row to headers
        PromotedHeaders = Table.PromoteHeaders(FirstSheetData, [PromoteAllScalars=true]),
        // Create a list of non-empty header names
        NonEmptyHeaderNames = List.Select(Table.ColumnNames(PromotedHeaders), each _ <> "" and _ <> null and not Text.StartsWith(_, "Column")),
        // Make sure the table includes the 'LT1_P_agr' column
        // If 'LT1_P_agr' is not in the headers, this will throw an error
        LT1_P_agrColumn = List.First(List.Select(NonEmptyHeaderNames, each Text.Contains(_, "LT1_P_agr"))),
        // Select only columns with non-empty headers
        CleanedTable = Table.SelectColumns(PromotedHeaders, NonEmptyHeaderNames),
        // Filter out rows where 'LT1_P_agr' is null or blank
        FilteredForLT1_P_agr = Table.SelectRows(CleanedTable, each Record.Field(_, LT1_P_agrColumn) <> null and Record.Field(_, LT1_P_agrColumn) <> ""),
        // Remove rows where fewer than 3 cells are non-empty
        NonEmptyRows = Table.SelectRows(FilteredForLT1_P_agr, each List.NonNullCount(Record.FieldValues(_)) >= 3)
    in
        NonEmptyRows
    ),
    #"Expanded Custom" = Table.ExpandTableColumn(#"Added Custom", "Custom", Table.ColumnNames(#"Added Custom"[Custom]{0})),
    #"ChangedTypes" = Table.TransformColumnTypes(#"Expanded Custom", {{"LT1_SID", type text}, {"LT1_NO", type text}}, "en-HK"),

    #"Transformed LT1_SID" = Table.TransformColumns(#"ChangedTypes", {
        {"LT1_SID", each if Text.Length(Text.From(_)) = 1 
            then "S00" & Text.From(_) 
            else if Text.Length(Text.From(_)) = 2 
            then "S0" & Text.From(_) 
            else "S" & Text.From(_), type text}
    }),

    #"Transformed LT1_CID" = Table.TransformColumns(#"Transformed LT1_SID", {
        {"LT1_CID", each if Text.Length(Text.From(_)) = 1 
            then "1000" & Text.From(_) 
            else if Text.Length(Text.From(_)) = 2 
            then "100" & Text.From(_) 
            else if Text.Length(Text.From(_)) = 3
            then "10" & Text.From(_) 
            else if Text.Length(Text.From(_)) = 4 
            then "1" & Text.From(_) 
            else Text.From(_), type text}
    }),

    #"Added Prefix" = Table.TransformColumns(#"Transformed LT1_CID", {
        {"LT1_NO", each "EdU_" & Text.From(_), type text}
    }),

    // Add your new column with a constant value of 2
    #"Added University Column" = Table.AddColumn(#"Added Prefix", "University", each 2),
    
    // Get a list of all column names up to and including LT1_NO
    ColumnNamesUpToLT1_NO = List.FirstN(Table.ColumnNames(#"Added Prefix"), List.PositionOf(Table.ColumnNames(#"Added Prefix"), "LT1_NO") + 1),
    
    // Append the new column name to the list
    ColumnNamesWithNewColumn = ColumnNamesUpToLT1_NO & {"University"} & List.Skip(Table.ColumnNames(#"Added Prefix"), List.PositionOf(Table.ColumnNames(#"Added Prefix"), "LT1_NO") + 1),

    // Reorder the columns to place the new column after LT1_NO
    #"Reordered Columns" = Table.ReorderColumns(#"Added University Column", ColumnNamesWithNewColumn),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Content", "Extension", "Date created", "Attributes"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Date modified", "LT1_Date_entry"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"LT1_Date_entry", type date}})


in
    #"Changed Type"