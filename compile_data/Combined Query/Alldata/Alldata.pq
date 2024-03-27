let
    // Load the HeaderMapping table from the 'HeaderMapping' sheet
    HeaderMapping = Excel.CurrentWorkbook(){[Name="HeaderMapping"]}[Content],

    // Load the Annualdata, PCPP, and ParentNameMatch tables
    ST1_Original = Excel.CurrentWorkbook(){[Name="ST1Alldata"]}[Content],
    PCPP_Original = Excel.CurrentWorkbook(){[Name="PCPPAlldata"]}[Content],
    LT1_Original = Excel.CurrentWorkbook(){[Name="LT1Alldata"]}[Content],
    ParentNameMatch_Original = Excel.CurrentWorkbook(){[Name="ParentNameMatch"]}[Content],
    AnnualIDMatch = Excel.CurrentWorkbook(){[Name="AnnualIDMatch"]}[Content],
    // AdditionalStatus = Excel.CurrentWorkbook(){[Name="AdditionalStatus"]}[Content],


    // Function to apply header mapping
    ApplyHeaderMapping = (table as table, headerMapping as table, oldHeaderColumn as text, newHeaderColumn as text) as table =>
        let
            ExistingHeaders = Table.ColumnNames(table),
            GetNewHeaderName = (oldHeaderName as text) as text =>
                let
                    MappingRow = Table.SelectRows(headerMapping, each Record.Field(_, oldHeaderColumn) = oldHeaderName),
                    NewHeaderName = if Table.IsEmpty(MappingRow) then oldHeaderName else Record.Field(MappingRow{0}, newHeaderColumn)
                in
                    NewHeaderName,
            RenamedTable = Table.TransformColumnNames(table, GetNewHeaderName)
        in
            RenamedTable,

    // Apply header mapping to tables
    RenamedST1 = Table.Buffer(ApplyHeaderMapping(ST1_Original, HeaderMapping, "ST1 Header", "XT1 Header")),
    RenamedPCPP = ApplyHeaderMapping(PCPP_Original, HeaderMapping, "PCPP Header", "XT1 Header"),
    RenamedLT1 = ApplyHeaderMapping(LT1_Original, HeaderMapping, "LT1 Header", "XT1 Header"),

    // Transform "Child ID" and "XT1_CID" columns to upper case in tables
    UpperPCPP = Table.TransformColumns(RenamedPCPP, {{"XT1_CID", Text.Upper, type text}}),
    ParentNameMatch = Table.TransformColumns(ParentNameMatch_Original, {{"Child ID", Text.Upper, type text}}),
    UpperAnnualID = Table.TransformColumns(AnnualIDMatch, {{"Child ID", Text.Upper, type text}}),   

    // Remove unnecessary columns from the PCPP table
    #"DistilledPCPP" = Table.RemoveColumns(UpperPCPP,{"Name", "Date accessed", "Folder Path", "Check_No", "XT1_Date_entry", "XT1_NO", "XT1_University", 
    "XT1_Version", "XT1_Pversion", "XT1_Tier", "XT1_Group", "XT1_SID", "XT1_PID", "XT1_Note", "XT1_Note2", "XT1_P_agr", "XT1_PName", "XT1_CName", "XT1_Sign", 
    "XT1_Date", "XT1_Dummy1", "XT1_Dummy2", "XT1_Dummy34", "XT1_Dummy35"}),   

    // Perform a left outer join on the PCPP table with ParentNameMatch table
    ParentmatchedPCPP = Table.NestedJoin(#"DistilledPCPP", {"XT1_CID"}, ParentNameMatch, {"Child ID"}, "MergedColumn", JoinKind.LeftOuter),
    ExpandedParentmatchedPCPP = Table.ExpandTableColumn(ParentmatchedPCPP, "MergedColumn", {"ST1matchPCPP"}, {"ST1matchPCPP"}),

    // Perform a left outer join on the ST1 table with the AnnualIDMatch table
    MergedTable = Table.NestedJoin(RenamedST1, {"XT1_NO"}, UpperAnnualID, {"XT1_NO"}, "MergedData", JoinKind.LeftOuter),

    // Remove the original "XT1_CID" column to avoid duplication
    RemovedOriginalXT1_CID = Table.RemoveColumns(MergedTable, "XT1_CID"),

    // Expand the "MergedData" column to include "Child ID" from the AnnualIDMatch table
    ST1withChildID = Table.ExpandTableColumn(RemovedOriginalXT1_CID, "MergedData", {"Child ID"}, {"XT1_CID"}),

    // Prepare the PCPP table for joining by filtering to include only matched rows
    PCPPonly = Table.SelectRows(ExpandedParentmatchedPCPP, each [ST1matchPCPP] = "Match"),

    // Perform an inner join on the ST1 with ChildID table and the filtered PCPP table
    ST1WithPCPP_beforeExpand = Table.NestedJoin(ST1withChildID, {"XT1_CID"}, PCPPonly, {"XT1_CID"}, "MergedData", JoinKind.Inner),

    // Extract the first row from the joined table for column name generation
    FirstRow = Table.First(ST1WithPCPP_beforeExpand),

    // Retrieve the 'MergedData' field, which contains a nested table
    FirstMergedDataTable = FirstRow[MergedData],

    // Get column names from the nested table within 'MergedData'
    ColumnNamesFromMergedData = Table.ColumnNames(FirstMergedDataTable),

    // Get existing column names from the ST1WithPCPP_beforeExpand table
    ExistingColumnNames = Table.ColumnNames(ST1WithPCPP_beforeExpand),

    // Create new column names to prevent duplicates during expansion
    NewColumnNames = List.Transform(ColumnNamesFromMergedData, each if List.Contains(ExistingColumnNames, _) then _ & "_Expanded" else _),

    // Expand all columns from the nested 'MergedData' tables using new column names
    ST1WithPCPP_0 = Table.ExpandTableColumn(ST1WithPCPP_beforeExpand, "MergedData", ColumnNamesFromMergedData, NewColumnNames),

    // Remove columns that are no longer needed
    ST1WithPCPP_1 = Table.RemoveColumns(ST1WithPCPP_0, {"XT1_CID_Expanded", "ST1matchPCPP"}),

    // Transform the 'XT1_PID' column to have all cells with the value "Matched_ST1PCPP"
    ST1WithPCPP = Table.TransformColumns(ST1WithPCPP_1, {{"XT1_PID", each "Matched_ST1PCPP", type text}}),

    // Create a list of "XT1_NO" values from the ST1WithPCPP table
    // ListXT1_NO_WithPCPP = Table.Column(#"Removed Columns", "XT1_NO"),

    // Join ST1withChildID with ST1WithPCPP to identify non-matching rows (replacing ST1WithoutPCPP step)
    ST1WithoutPCPP0 = Table.NestedJoin(ST1withChildID, {"XT1_CID"}, ST1WithPCPP, {"XT1_CID"}, "PCPPData", JoinKind.LeftAnti),

    // Remove the temporary join column "PCPPData" as it's not needed in the final output
    ST1WithoutPCPP = Table.RemoveColumns(ST1WithoutPCPP0, {"PCPPData"}),

    // Combine the LT1 table with the ST1WithPCPP and ST1WithoutPCPP tables
    CombinedData = Table.Combine({RenamedLT1, ST1WithPCPP, ST1WithoutPCPP}),
    
    // // Perform a left outer join on CombinedTable with the AdditionalStatus table
    // CombinedTableWithStatus = Table.NestedJoin(CombinedData, {"Child ID"}, AdditionalStatus, {"Child ID"}, "AdditionalStatus", JoinKind.LeftOuter),
    // // Expand the WithdrawalStatus column from the joined table
    // ExpandedCombinedTableWithStatus = Table.ExpandTableColumn(CombinedTableWithStatus, "AdditionalStatus", {"WithdrawalStatus"}),
    // // Add a custom column to conditionally update XT1_Note based on WithdrawalStatus
    // FinalCombinedTable = Table.AddColumn(ExpandedCombinedTableWithStatus, "XT1_Note", each if [WithdrawalStatus] <> "ID not found" then [WithdrawalStatus] else [XT1_Note]),
    // // Remove the original XT1_Note column if it's no longer needed
    // RemovedOriginalXT1_Note = Table.RemoveColumns(FinalCombinedTable, {"XT1_Note"}),
    // // Rename the custom XT1_Note to the original column name if needed
    // RenamedXT1_Note = Table.RenameColumns(RemovedOriginalXT1_Note, {{"XT1_Note.1", "XT1_Note"}}),

    // Replace blank or null cells with "#blank"
    ReplaceBlanks = Table.TransformColumns(
        CombinedData,
        List.Transform(Table.ColumnNames(CombinedData), (columnName) =>
            {columnName, each if _ is null or _ = "" then "#blank" else _})
    ),
    DateOnly = Table.TransformColumnTypes(ReplaceBlanks,{{"XT1_Date_entry", type date}}),

    // // Apply the date transformation to the specified columns to TEXT
    // DateColumns = {"Date accessed", "XT1_Date_entry", "XT1_Date", "XT1_B_Day", "XT1_M_DOB", "XT1_F_DOB"}, // List of date columns
    // ReplaceBlanksAndTransformDates = (value as any) as any =>
    //     let
    //         textValue = Text.From(value),
    //         result = if List.Contains({"6666", "7777", "8888", "9999"}, textValue) then textValue
    //                 else if value is null or value = "" then "#blank"
    //                 else try DateTime.ToText(DateTime.From(value), "dd/MM/yyyy") otherwise textValue
    //     in
    //         result,
    // Create a list of transformations for the specified date columns
    // Transformations = List.Transform(DateColumns, (columnName) => 
    //     {columnName, each ReplaceBlanksAndTransformDates(_), type text}),
    // Apply the transformations to the table
    // TransformedTable = Table.TransformColumns(DateColumns, Transformations),

    #"Sorted Rows" = Table.Sort(DateOnly,{{"XT1_Date_entry", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "Index", 1, 1, Int64.Type)
in 
    #"Added Index"