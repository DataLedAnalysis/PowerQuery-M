/* 
- Add feb filter (already done) 
- based off reporting month. If Month_No = x then return FinancialMonth 

*/
let

  //declaration section                     
  Source = Exchange.Contents("commercialperformance@thameswater.co.uk"),
  FinancialMonths = {10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9},  //declaring reporting months                                                                                                                                                                                                    
    //this is made dynamic, used for reporting month                                                                                                                                                                     
  //actual query starts here                          
  Mail_item = Source{[Name = "Mail"]}[Data],
  DateAttchFilter = Table.SelectRows(
    Mail_item,
    each ([Folder Path] = "\Inbox\")
      and ([HasAttachments] = true)
      and (
      /* Make input for year and month dynamic */
        Date.IsInPreviousMonth([DateTimeReceived])//set mail datetime filter  
        )
      and (
        Date.IsInCurrentYear([DateTimeReceived])
      ) 
  ),  
  DateCell = DateAttchFilter{0}[DateTimeReceived],

  AttchExpand = Table.ExpandTableColumn(
    DateAttchFilter,
    "Attachments",
    {"Name", "Extension", "ContentType", "AttachmentContent"},
    {
      "Attachments.Name",
      "Attachments.Extension",
      "Attachments.ContentType",
      "Attachments.AttachmentContent"
    }
  ),
  #"Expanded Sender" = Table.ExpandRecordColumn(
    AttchExpand,
    "Sender",
    {"Name", "Address"},
    {"Sender.Name", "Sender.Address"}
  ),
  //actual transformations begin here                                   
  replace_ = Table.ReplaceValue(
    #"Expanded Sender",
    "_",
    " ",
    Replacer.ReplaceText,
    {"Attachments.Name"}
  ),
  #"replace-" = Table.ReplaceValue(replace_, "-", " ", Replacer.ReplaceText, {"Attachments.Name"}),
  #"Filtered Rows" = Table.SelectRows(
    #"replace-",
    each Text.Contains([Attachments.Name], "spend report", Comparer.OrdinalIgnoreCase)
      and Text.Contains([Attachments.Extension], ".xlsx")
  ),  //attachment name and extension filter                                                                                                                                                                                                                                                                                  
  #"Filtered Hidden Files1" = Table.SelectRows(
    #"Filtered Rows",
    each [Attributes]?[Hidden]? <> true
  ),
  #"Filtered Hidden Files2" = Table.SelectRows(
    #"Filtered Hidden Files1",
    each [Attributes]?[Hidden]? <> true
  ),
  #"Invoke Custom Function1" = Table.AddColumn(
    #"Filtered Hidden Files2",
    "Transform File (3)",
    each #"Transform File (3)"([Attachments.AttachmentContent])
  ),
  RemovedErrorsColumns = Table.SelectColumns(#"Invoke Custom Function1", {"Transform File (3)"}),
  #"Removed Errors1" = Table.RemoveRowsWithErrors(RemovedErrorsColumns, {"Transform File (3)"}),
  #"Expanded Table Column1" = Table.ExpandTableColumn(
    #"Removed Errors1",
    "Transform File (3)",
    Table.ColumnNames(#"Transform File (3)"(#"Sample File (3)"))
  ),
  //transformations                 
  #"Removed Top Rows" = Table.Skip(#"Expanded Table Column1", 4),
  #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows"),
  #"Repeating rows filtered" = Table.SelectRows(
    #"Promoted Headers",
    each (
      [Supplier Name]
        <> null and [Supplier Name]
        <> " " and [Supplier Name]
        <> "e.g. ABC Limited" and [Supplier Name]
        <> "Enter Supplier Name" and [Supplier Name]
        <> "Monthly Spend Report" and [Supplier Name]
        <> "Supplier Name"
    )
  ),
  ColumnIthRemoved = Table.RemoveColumns(#"Repeating rows filtered", {"Column1", "Column20"}),
  #"Date casting" = Table.TransformColumnTypes(
    ColumnIthRemoved,
    {
      {"PO Number", type text},
      {"Invoice date", Date.Type},
      {"Delivery/Service Completion Date", Date.Type}
    }
  ),
  //this section is all about reporting period        

  //temporary variable

  d = Date.AddMonths(Date.From(DateTime.LocalNow()),-2), //reporting month is a month behind the email month
    //for combinbing data with errors 

  //reporting times                                       
  ReportingMonth = Table.AddColumn(#"Date casting", "REPORTING MONTH", each Date.ToText(d, "MMM" & "-" & "yy")),  //based off todays date                                                                                                                                                                                 
  ReportingMonthNo = Table.AddColumn(
    ReportingMonth,
    "REPORTING MONTH NUMBER",
    each Date.Month(d)
  ),
  ReportingMonthName = Table.AddColumn(
    ReportingMonthNo,
    "REPORTING MONTH NAME",
    each Date.ToText(d, "MMM")
  ),
  ReportingYear = Table.AddColumn(ReportingMonthName, "REPORTING-YEAR", each Date.Year(d)),
  //not based on invoice date                           
  AddFinancialMonth = Table.AddColumn(
    ReportingYear,
    "FINANCIAL_PERIOD",
    each FinancialMonths{Number.From([REPORTING MONTH NUMBER]) - 1}
  ),  //based off reporting month. If Month_No = x then return FinancialMonth                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
  //added financial year                      
  FY = Table.AddColumn(
    AddFinancialMonth,
    "Financial year",
    each
      if Date.Month([Invoice date]) > 3 then
        (Date.ToText([Invoice date], "yyyy"))
          & "/"
          & (Date.ToText(Date.AddYears([Invoice date], 1), "yy"))
      else
        (Date.ToText(Date.AddYears([Invoice date], - 1), "yyyy"))
          & "/"
          & (Date.ToText([Invoice date], "yy"))
  ),
  //ProperTrim            
  ProperUnderscore = Table.TransformColumnNames(FY, each Text.Upper(Text.Replace(_, " ", "_"))),
  FANumber = Table.NestedJoin(
    ProperUnderscore,
    {"FRAMEWORK_AGREEMENT_NUMBER"},
    d_Mapping,
    {"Framework Number"},
    "Mapping",
    JoinKind.LeftOuter
  ),
  //Mapping left outer join                         
  #"Expanded Mapping" = Table.ExpandTableColumn(
    FANumber,
    "Mapping",
    {"Business Unit name", "Contracts Manager", "Status"},
    {"Business Unit name", "Contracts Manager", "Status"}
  ),
  InMillions = Table.AddColumn(
    #"Expanded Mapping",
    "Spend (£m)",
    each Value.Divide(Number.From([#"LINE_ITEM_VALUE_£."]), Number.Power(10, 6))
  ),
  #"Renamed Columns" = Table.RenameColumns(
    InMillions,
    {
      {"LINE_ITEM_VALUE_£.", "Line Item Value £"},
      {"Business Unit name", "BUSINESS_UNIT"},
      {"Contracts Manager", "CONTRACT_MANAGER_BASED_ON_FA"},
      {"Status", "LIVE_FA"},
      {"FRAMEWORK_AGREEMENT_NUMBER", "Framework Agreement Number"},
      {"REPORTING-YEAR", "REPORTING_YEAR"},
      {"DELIVERY/SERVICE_COMPLETION_DATE", "DELIVERY_SERVICE_COMP_DATE"}
    }
  ),
  Datatype = Table.TransformColumnTypes(
    #"Renamed Columns",
    {
      {"Line Item Value £", type number},
      {"SUPPLIER_NAME", type text},
      {"INVOICE_NUMBER", type text},
      {"INVOICE_DATE", type date},
      {"PO_NUMBER", type text},
      {"ALLIANCES_PO_NUMBER", type text},
      {"DELIVERY_LOCATION", type text},
      {"DELIVERY_SERVICE_COMP_DATE", type date},
      {"DELIVERY_NOTE_NUMBER", type text},
      {"THAMES_WATER/ALLIANCE_NAME", type text},
      {"SUB_CONTRACTOR_NAME", type text},
      {"Framework Agreement Number", type text},
      {"LOT_NUMBER", type text},
      {"PO_LINE_ITEM_NUMBER", type text},
      {"LINE_ITEM_DESCRIPTION", type text},
      {"QUANTITY_ORDERED", Int64.Type},
      {"COMMENTS", type text},
      {"REPORTING_MONTH", type text},
      {"REPORTING_MONTH_NUMBER", Int64.Type},
      {"REPORTING_MONTH_NAME", type text},
      {"REPORTING_YEAR", Int64.Type},
      {"FINANCIAL_PERIOD", Int64.Type},
      {"FINANCIAL_YEAR", type text},
      {"BUSINESS_UNIT", type text},
      {"CONTRACT_MANAGER_BASED_ON_FA", type text},
      {"LIVE_FA", type text},
      {"Spend (£m)", type number}
    }
  ),
  RemovedStringRows = Table.RemoveRowsWithErrors(
    Datatype,
    {"Line Item Value £", "QUANTITY_ORDERED"}
  ),
  SACCharLimit = Table.TransformColumns(
    RemovedStringRows,
    {{"LINE_ITEM_DESCRIPTION", each Text.Start(_, 250), type text}}
  ),  //Removing rows of no value in line item                                          
  LineItemNotBlank = Table.SelectRows(SACCharLimit, each ([#"Line Item Value £"] <> null)),
  oldSupplierNaming = Table.RenameColumns(
    LineItemNotBlank,
    {{"SUPPLIER_NAME", "old_SUPPLIER_NAME"}}
  ),
    //added Jon's queries                       
  NameVendor = Table.NestedJoin(
    oldSupplierNaming,
    {"old_SUPPLIER_NAME"},
    #"d_SupplierName&VendorID",
    {"SupplierID"},
    "d_SupplierName&VendorID",
    JoinKind.LeftOuter
  ),
  Alliance = Table.NestedJoin(
    NameVendor,
    {"THAMES_WATER/ALLIANCE_NAME"},
    d_ParentAlliance,
    {"AllianceID"},
    "d_ParentAlliance",
    JoinKind.LeftOuter
  ),
  ExpandedAlliance = Table.ExpandTableColumn(Alliance, "d_ParentAlliance", {"PARENT_ALLIANCE_NAME"}),
  ExpandedSupplyVendor = Table.ExpandTableColumn(
    ExpandedAlliance,
    "d_SupplierName&VendorID",
    {"SAP_VENDOR_ID", "SUPPLIER_NAME"}
  ),
  ReorderedNewSupplierName = Table.ReorderColumns(
    ExpandedSupplyVendor,
    {
      "SUPPLIER_NAME",
      "old_SUPPLIER_NAME",
      "INVOICE_NUMBER",
      "INVOICE_DATE",
      "PO_NUMBER",
      "ALLIANCES_PO_NUMBER",
      "DELIVERY_LOCATION",
      "DELIVERY_SERVICE_COMP_DATE",
      "DELIVERY_NOTE_NUMBER",
      "THAMES_WATER/ALLIANCE_NAME",
      "SUB_CONTRACTOR_NAME",
      "Framework Agreement Number",
      "LOT_NUMBER",
      "PO_LINE_ITEM_NUMBER",
      "LINE_ITEM_DESCRIPTION",
      "QUANTITY_ORDERED",
      "Line Item Value £",
      "COMMENTS",
      "REPORTING_MONTH",
      "REPORTING_MONTH_NUMBER",
      "REPORTING_MONTH_NAME",
      "REPORTING_YEAR",
      "FINANCIAL_PERIOD",
      "SAP_VENDOR_ID",
      "FINANCIAL_YEAR",
      "BUSINESS_UNIT",
      "CONTRACT_MANAGER_BASED_ON_FA",
      "LIVE_FA",
      "Spend (£m)"
    }
  ),
  #"Trimmed Text" = Table.TransformColumns(
    ReorderedNewSupplierName,
    {{"Framework Agreement Number", Text.Trim, type text}}
  ),
  #"Cleaned Text" = Table.TransformColumns(
    #"Trimmed Text",
    {{"Framework Agreement Number", Text.Clean, type text}}
  ),
  #"Uppercased Text" = Table.TransformColumns(
    #"Cleaned Text",
    {{"Framework Agreement Number", Text.Upper, type text}}
  ),
    #"Appended Query" = Table.Combine({#"Uppercased Text", f_Append_to_SupplierSpend})
in
    #"Appended Query"
