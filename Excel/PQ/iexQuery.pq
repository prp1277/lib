/** Get Data */
let
Parameter = Excel.CurrentWorkbook(){[Name="Parameters"]}[Content],
URL = Parameter{0}[Value],
Source = Json.Document(Web.Contents(URL)),
    #"Converted to Table" = Record.ToTable(Source)
in
    #"Converted to Table"

/** News */
let
    Source = URLQuery,
    #"Expanded Value" = Table.ExpandRecordColumn(Source, "Value", {"news"}, {"news"}),
    #"Expanded news" = Table.ExpandListColumn(#"Expanded Value", "news"),
    #"Expanded news1" = Table.ExpandRecordColumn(#"Expanded news", "news", {"datetime", "headline", "source", "url", "summary", "related", "image"}, {"datetime", "headline", "source", "url", "summary", "related", "image"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded news1",{{"datetime", type datetimezone}, {"headline", type text}, {"source", type text}, {"url", type text}, {"summary", type text}, {"related", type text}, {"image", type text}}),
    #"Trimmed Text" = Table.TransformColumns(#"Changed Type",{{"image", Text.Trim, type text}, {"related", Text.Trim, type text}, {"summary", Text.Trim, type text}, {"url", Text.Trim, type text}, {"source", Text.Trim, type text}, {"headline", Text.Trim, type text}}),
    #"Cleaned Text" = Table.TransformColumns(#"Trimmed Text",{{"image", Text.Clean, type text}, {"related", Text.Clean, type text}, {"summary", Text.Clean, type text}, {"url", Text.Clean, type text}, {"source", Text.Clean, type text}, {"headline", Text.Clean, type text}})
in
    #"Cleaned Text"

/** Chart */
let
    Source = URLQuery,
    #"Expanded Value" = Table.ExpandRecordColumn(Source, "Value", {"chart"}, {"chart"}),
    #"Expanded chart" = Table.ExpandListColumn(#"Expanded Value", "chart"),
    #"Expanded chart1" = Table.ExpandRecordColumn(#"Expanded chart", "chart", {"date", "open", "high", "low", "close", "volume", "unadjustedVolume", "change", "changePercent", "label", "changeOverTime"}, {"date", "open", "high", "low", "close", "volume", "unadjustedVolume", "change", "changePercent", "label", "changeOverTime"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded chart1",{{"date", type date}, {"open", Currency.Type}, {"high", Currency.Type}, {"low", Currency.Type}, {"close", Currency.Type}, {"volume", Int64.Type}, {"unadjustedVolume", Int64.Type}, {"label", type date}, {"change", Currency.Type}, {"changePercent", Percentage.Type}, {"changeOverTime", Percentage.Type}}),
    #"Divided Column" = Table.TransformColumns(#"Changed Type", {{"changePercent", each _ / 100, Percentage.Type}}),
    #"Sorted Rows" = Table.Sort(#"Divided Column",{{"Name", Order.Ascending}, {"date", Order.Ascending}})
in
    #"Sorted Rows"