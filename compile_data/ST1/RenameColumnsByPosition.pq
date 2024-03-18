// Function to rename columns based on position
(tbl as table, newHeaders as list) as table =>
let
    OriginalHeaders = Table.ColumnNames(tbl),
    NumOfColumns = List.Count(OriginalHeaders),
    NumOfNewHeaders = List.Count(newHeaders),
    // Ensure the number of headers matches the number of columns in the table
    RenamedTable = if NumOfColumns = NumOfNewHeaders then
        Table.RenameColumns(tbl, List.Zip({OriginalHeaders, newHeaders}), MissingField.Ignore)
    else
        error "The number of columns in the table does not match the number of new headers."
in
    RenamedTable