let
    // Reference the FolderPathsQuery
    FolderPathsQuery = #"FolderPathsQuery",
    // Retrieve the folder path from the table
    // Make sure to use the correct row index (0 for the first row)
    FolderPath = FolderPathsQuery{0}[FolderPath],
    // Get the files from the specified folder
    Source = Folder.Files(FolderPath),
    
    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Name], "Annual Qx")), // Changed to a more generic filter
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Custom", each let
        // Access the content of the file
        ExcelContent = Excel.Workbook([Content]),
        // Select the first sheet
        FirstSheetData = ExcelContent{0}[Data],
        // Promote the first row to headers
        PromotedHeaders = Table.PromoteHeaders(FirstSheetData, [PromoteAllScalars=true]),
        // Create a list of non-empty header names
        NonEmptyHeaderNames = List.Select(Table.ColumnNames(PromotedHeaders), each _ <> "" and _ <> null and not Text.StartsWith(_, "Column")),
        // Make sure the table includes the 'ST1_NO_Column' column
        // If 'ST1_NO_Column' is not in the headers, this will throw an error
        ST1_NO_Column = List.First(List.Select(NonEmptyHeaderNames, each Text.Contains(_, "ST1_NO"))),
        // Select only columns with non-empty headers
        CleanedTable = Table.SelectColumns(PromotedHeaders, NonEmptyHeaderNames),
        // Replace "#NULL!" values with null in the CleanedTable table
        ReplacedNulls = Table.ReplaceValue(CleanedTable,"#NULL!",null,Replacer.ReplaceValue,NonEmptyHeaderNames),
        // Handle errors in the entire table
        ReplacedErrors = Table.ReplaceErrorValues(ReplacedNulls, List.Transform(Table.ColumnNames(ReplacedNulls), each {_, null})),
        // Filter out rows containing 'ST1_NO_Column'
        FilteredForST1_NO_Column = Table.SelectRows(ReplacedErrors, each Record.Field(_, ST1_NO_Column) <> null and Record.Field(_, ST1_NO_Column) <> "ST1_NO"),

    ///inserting Check_No
        // Previous steps that produce the #"Reordered Columns" step
        #"Added Check_No" = Table.AddColumn(FilteredForST1_NO_Column, "Check_No", each "not applicable"),
        // Get a list of all column names after adding the new column
        ColumnNames = Table.ColumnNames(#"Added Check_No"),
        // Find the index of the "ST1_NO" column
        TargetColumnIndex = List.PositionOf(ColumnNames, "ST1_NO"),
        // Create a list of column names with "Check_No" inserted before "ST1_NO"
        ColumnOrderWith_CheckNo = List.InsertRange(List.RemoveItems(ColumnNames, {"Check_No"}), TargetColumnIndex, {"Check_No"}),
        // Now reorder the table based on the new list of column names
        CheckNoTable = Table.ReorderColumns(#"Added Check_No", ColumnOrderWith_CheckNo)
    in
        CheckNoTable
    ),
    #"Expanded Custom" = Table.ExpandTableColumn(#"Added Custom", "Custom", Table.ColumnNames(#"Added Custom"[Custom]{0})),
    // Rename 'Date created' to 'ST1_Date_entry' and reorder columns
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Custom",{{"Date created", "ST1_Date_entry"}}),
///inserting ST1_Date_entry
    // Get a list of all column names after adding the new column
    ColumnNames2 = Table.ColumnNames(#"Renamed Columns"),
    // Find the index of the "ST1_NO" column
    TargetColumnIndex2 = List.PositionOf(ColumnNames2, "Check_No"),
    // Create a list of column names with "Check_No" inserted before "ST1_NO"
    ColumnOrderWith_Date = List.InsertRange(List.RemoveItems(ColumnNames2, {"ST1_Date_entry"}), TargetColumnIndex2, {"ST1_Date_entry"}),
    // Now reorder the table based on the new list of column names
    Date_entryTable = Table.ReorderColumns(#"Renamed Columns", ColumnOrderWith_Date),
///recoding ST1_NO 
    // Ensure that 'ST1_NO' and 'ST1_SID' are text types
    TypedSource = Table.TransformColumnTypes(Date_entryTable, {{"ST1_NO", type text}, {"ST1_SID", type text}}),
    // Add a new column with the concatenated values
    #"Added_NEW_ST1_NO" = Table.AddColumn(TypedSource, "New_ST1_NO", each "HKU" & [ST1_SID] & "_" & [ST1_NO]),
    Removed_ST1_NO = Table.RemoveColumns(Added_NEW_ST1_NO, {"ST1_NO"}),
    // Optionally reorder columns to place 'New_ST1_NO' where desired
    // Assuming you want to place 'New_ST1_NO' before 'Check_No':
    ColumnNames = Table.ColumnNames(Removed_ST1_NO),
    University_Index = List.PositionOf(ColumnNames, "University"),
    ColumnOrder = List.InsertRange(List.RemoveItems(ColumnNames, {"New_ST1_NO"}), University_Index, {"New_ST1_NO"}),
    New_ST1_NOTable = Table.ReorderColumns(Removed_ST1_NO, ColumnOrder),
    // Rename 'New_ST1_NO' back to 'ST1_NO'
    #"RenamedTable" = Table.RenameColumns(New_ST1_NOTable,{{"New_ST1_NO", "ST1_NO"}}),
    #"Transformed ST1_SID" = Table.TransformColumns(#"RenamedTable", {
        {"ST1_SID", each if Text.Length(Text.From(_)) = 1 
            then "S00" & Text.From(_) 
            else if Text.Length(Text.From(_)) = 2 
            then "S0" & Text.From(_) 
            else "S" & Text.From(_), type text}
    }),
    #"FilteredRows" = Table.SelectRows(#"Transformed ST1_SID", each [ST1_NO] <> "HKU_"),
    #"Removed Columns" = Table.RemoveColumns(FilteredRows,{"Content", "Extension", "Date modified", "Attributes"})
in
    #"Removed Columns"