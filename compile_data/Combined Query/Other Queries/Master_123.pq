let
    Source = Excel.Workbook(Web.Contents("https://example.com/your_master_data.xlsx"), null, true),
    #"Master (123)_Sheet" = Source{[Item="Master (123)",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Master (123)_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"School ID", type text}, {"Group", type any}, {"District (Eng)", type text}, {"District", type text}, {"School Name (Chi)", type text}, {"Address (Chi)", type text}, 
    {"Contact Person", type text}, {"Contact Number", type any}, {"Number of students (received + sent) #(lf)#(lf)(!!Do not edit!!)#(lf)", Int64.Type}, {"Number of students (sent) #(lf)#(lf)(!!Do not edit!!)", Int64.Type}, 
    {"Chinese Question-naire Distributed #(lf)#(lf)(!!Do not edit!!)", Int64.Type}, {"English Question-naire Distributed #(lf)#(lf)(!!Do not edit!!)", Int64.Type}, {"Distribution #(lf)Date", type any}, {"Distribution Mean ", type text}, 
    {"SF Express Tracking No. ", type text}, {"Number of questionnaire (ST1) received (Agree + Disagree)#(lf)#(lf)(Edit here)", type any}, {"Collection Date#(lf)#(lf)(Edit here)", type date}, {"Collected by#(lf)#(lf)(Edit here)", type text}, 
    {"coupon distribution reminder day", type date}, {"coupon distribution due day", type date}, {"Suggested to be collected by which university #(lf)#(lf)(!!Do not edit!!)", type any}, {"Remarks", type text}}),
    #"Filtered Column Names" = List.Select(Table.ColumnNames(#"Changed Type"), each Text.StartsWith(_, "Column")),
    #"Removed Blank Header Columns" = Table.RemoveColumns(#"Changed Type", #"Filtered Column Names"),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Blank Header Columns",{"District", "Address (Chi)", "Contact Person", "Contact Number", "Chinese Question-naire Distributed #(lf)#(lf)(!!Do not edit!!)", 
    "English Question-naire Distributed #(lf)#(lf)(!!Do not edit!!)", "Distribution #(lf)Date", "Distribution Mean ", "SF Express Tracking No. ", "Suggested to be collected by which university #(lf)#(lf)(!!Do not edit!!)", "Remarks"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Number of students (received + sent) #(lf)#(lf)(!!Do not edit!!)#(lf)", "Number of students (received + sent)"}, {"Number of students (sent) #(lf)#(lf)(!!Do not edit!!)", 
    "Number of students (sent)"}, {"Number of questionnaire (ST1) received (Agree + Disagree)#(lf)#(lf)(Edit here)", "Number of students (received)"}, {"Collection Date#(lf)#(lf)(Edit here)", "Collection Date"}, 
    {"Collected by#(lf)#(lf)(Edit here)", "Collected by"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Group", Int64.Type}})
in
    #"Changed Type1"