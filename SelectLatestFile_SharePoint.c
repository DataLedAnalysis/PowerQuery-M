let
  Source = SharePoint.Files(
    "https://thameswater.sharepoint.com/sites/CommStratInsightReport",
    [ApiVersion = 15]
  ),
  #"Reordered Columns" = Table.ReorderColumns(
    Source,
    {
      "Folder Path",
      "Content",
      "Name",
      "Extension",
      "Date accessed",
      "Date modified",
      "Date created",
      "Attributes"
    }
  ),
  #"Filtered Rows" = Table.SelectRows(
    #"Reordered Columns",
    each (
      [Folder Path]
        = "https://thameswater.sharepoint.com/sites/CommStratInsightReport/Regular reports/PULSE pivot/"
    )
  ),
  FileNameToBeginning = Table.ReorderColumns(
    #"Filtered Rows",
    {
      "Name",
      "Folder Path",
      "Content",
      "Extension",
      "Date accessed",
      "Date modified",
      "Date created",
      "Attributes"
    }
  ),
  LatestDate = Table.SelectRows(
    FileNameToBeginning,
    let
      latest = List.Max(FileNameToBeginning[Date created])
    in
      each [Date created] = latest
  ),
  //this allows us to by dynamic with queries                                             
  OpenFile = LatestDate
    {
      [
        Name =
          let
            TableRowVal = LatestDate[Name],
            SelectedRow = TableRowVal{0}
          in
            SelectedRow,
        #"Folder Path"
          = "https://thameswater.sharepoint.com/sites/CommStratInsightReport/Regular reports/PULSE pivot/"
      ]
    }
    [Content],
  //above function return gets imported to this line                                                    
  #"Imported Excel Workbook" = Excel.Workbook(OpenFile),
  PQ_Sheet = #"Imported Excel Workbook"{[Item = "PQ", Kind = "Sheet"]}[Data],
  #"Promoted Headers" = Table.PromoteHeaders(PQ_Sheet, [PromoteAllScalars = true]),
  #"Changed Type" = Table.TransformColumnTypes(
    #"Promoted Headers",
    {
      {"Project Code", type text},
      {"Project Name", type text},
      {"Sum of Savings - High confidence", type number},
      {"Sum of Savings - Medium confidence", type number},
      {"Sum of Savings - Low confidence", type number}
    }
  )
in
  #"Changed Type"
