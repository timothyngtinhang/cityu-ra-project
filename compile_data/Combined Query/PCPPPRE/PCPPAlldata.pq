let
    FolderPathsQuery = #"FolderPathsQuery",
    // Retrieve the folder path from the table
    FolderPath = FolderPathsQuery{2}[FolderPath],
    // Get the files from the specified folder
    Source = Folder.Files(FolderPath),
    // Filter the rows to only include files that contain "CheckCurrent" in their name
    FilteredRows = Table.SelectRows(Source, each Text.Contains([Name], "Check")), // Changed to a more generic filter

    // Define a function to process individual files
    ProcessFile = (File) => 
    let
        // Access the content of the file
        ExcelContent = Excel.Workbook(File[Content]),
        // Select the specific sheet by name
        SpecificSheet = Table.SelectRows(ExcelContent, each [Kind] = "Sheet"), // Removed specific sheet name
        FirstSheetData = SpecificSheet{0}[Data],
        // Promote the first row to headers
        PromotedHeaders = Table.PromoteHeaders(FirstSheetData, [PromoteAllScalars=true]),
        // Create a list of non-empty header names
        NonEmptyHeaderNames = List.Select(Table.ColumnNames(PromotedHeaders), each _ <> "" and _ <> null and not Text.StartsWith(_, "Column")),
        // Make sure the table includes the 'PCPP_PRE1_Date_entry' column
        PCPPDateEntryColumn = List.First(List.Select(NonEmptyHeaderNames, each Text.Contains(_, "PCPP_PRE1_Date_entry"))),
        // Select only columns with non-empty headers
        CleanedTable = Table.SelectColumns(PromotedHeaders, NonEmptyHeaderNames),
        // Filter out rows containing 'PCPP_PRE1_Date_entry'
        FilteredForPCPPDateEntry = Table.SelectRows(CleanedTable, each Record.Field(_, PCPPDateEntryColumn) <> null and Record.Field(_, PCPPDateEntryColumn) <> "PCPP_PRE1_Date_entry"),
        // Remove rows where fewer than 3 cells are non-empty
        NonEmptyRows = Table.SelectRows(FilteredForPCPPDateEntry, each List.NonNullCount(Record.FieldValues(_)) >= 3)
    in
        NonEmptyRows,

    // Apply the function to each file in the filtered rows
    AddedCustom = Table.AddColumn(FilteredRows, "Custom", each ProcessFile(_)),

    // Expand the 'Custom' column if 'AddedCustom' is not empty
    ExpandedCustom = if Table.IsEmpty(AddedCustom) then AddedCustom 
    else Table.ExpandTableColumn(AddedCustom, "Custom", Table.ColumnNames(AddedCustom[Custom]{0})),
    #"Removed Columns" = Table.RemoveColumns(ExpandedCustom,{"Content", "Extension", "Date modified", "Date created", "Attributes", "Check_Date", "Error_Items"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"PCPP_PRE1_CID", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each [PCPP_PRE1_P_agr] = 1),
        // Append suffix "_LT1" to every cell in the "ST1_NO" column
    SuffixedPCPP_NO = Table.TransformColumns(
        #"Filtered Rows",
        {"PCPP_PRE1_NO", each _ & "_PCPP-PRE1", type text}
    )
in
    SuffixedPCPP_NO