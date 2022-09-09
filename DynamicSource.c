
//MainReport
let
      Source = List.Generate(()=>
[Result = try pageRecords(1) otherwise null,Page = 1],each [Result] <> null, each
[Result = try pageRecords([Page]+1) otherwise null,Page = [Page]+1],
each [Result]
),
  TblFromList = Table.FromList(
    Source, 
    Splitter.SplitByNothing(), 
    null, 
    null, 
    ExtraValues.Error
  ),
    ExpandCol = Table.ExpandTableColumn(TblFromList, "Column1", {"Project Code", "Project Name", "Project Owner", "Business Unit name", "Project Type", "PMO status", "FA number (General tab)", "Date Standstill Period Ends", "PMO closed down date", "PSC Action -  Date MGT records set up  (TA Relevant)", "ITN/RfP Submission Deadline", "PQQ Submission Deadline"}, {"Project Code", "Project Name", "Project Owner", "Business Unit name", "Project Type", "PMO status", "FA number (General tab)", "Date Standstill Period Ends", "PMO closed down date", "PSC Action -  Date MGT records set up  (TA Relevant)", "ITN/RfP Submission Deadline", "PQQ Submission Deadline"}),
    Casted = Table.TransformColumnTypes(ExpandCol,{{"Project Code", type text}, {"Project Name", type text}, {"Project Owner", type text}, {"Business Unit name", type text}, {"Project Type", type text}, {"PMO status", type text}, {"FA number (General tab)", type text}, {"Date Standstill Period Ends", type date}, {"PMO closed down date", type date}, {"PSC Action -  Date MGT records set up  (TA Relevant)", type date}, {"ITN/RfP Submission Deadline", type date}, {"PQQ Submission Deadline", type date}}), 
  SS_StandStillNotNull = Table.SelectRows(
    Casted, 
    each ([Date Standstill Period Ends] <> null)
  ),
    #"Sorted Rows" = Table.Sort(SS_StandStillNotNull,{{"Date Standstill Period Ends", Order.Ascending}}),
    NotCapProc = Table.SelectRows(
    #"Sorted Rows", 
    each not Text.Contains([Project Type], "Capital Procurement Projects")
  ),
    SS_AfterApr2021 = Table.SelectRows(NotCapProc, each [Date Standstill Period Ends] > #date(2021, 4, 1))
in
    SS_AfterApr2021
  //Replaced blank pmo dates with today  
  
  //Calling function
  
  (Page as number) =>

let
  Source = Json.Document(
    Web.Contents("https://api.keyedinprojects.co.uk/V3/api/report", [RelativePath="?key=771",Query=[pageNumber=Number.ToText(Page)]]
  )

),
    Data = Source[Data],
    ToTbl = Table.FromList(Data, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    ExpandRecords = Table.ExpandRecordColumn(ToTbl, "Column1", {"Project Code", "Project Name", "Project Owner", "Business Unit name", "Project Type", "PMO status", "FA number (General tab)", "Date Standstill Period Ends", "PMO closed down date", "PSC Action -  Date MGT records set up  (TA Relevant)", "ITN/RfP Submission Deadline", "PQQ Submission Deadline"}, {"Project Code", "Project Name", "Project Owner", "Business Unit name", "Project Type", "PMO status", "FA number (General tab)", "Date Standstill Period Ends", "PMO closed down date", "PSC Action -  Date MGT records set up  (TA Relevant)", "ITN/RfP Submission Deadline", "PQQ Submission Deadline"})
in
    ExpandRecords
