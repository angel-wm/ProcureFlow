let
    RawDataFolder =
        Text.From(
            Excel.CurrentWorkbook(){[Name="cfg_RawDataFolder"]}[Content]{0}[Column1]
        ),
    SourceFile = RawDataFolder & "quality_incidents.csv",
    ImportedCSV =
        Csv.Document(
            File.Contents(SourceFile),
            [
                Delimiter = ",",
                Encoding = 65001,
                QuoteStyle = QuoteStyle.Csv
            ]
        ),
    PromotedHeaders =
        Table.PromoteHeaders(
            ImportedCSV,
            [PromoteAllScalars = true]
        )
in
    PromotedHeaders