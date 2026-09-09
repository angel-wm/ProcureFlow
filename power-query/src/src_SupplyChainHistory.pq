let
    RawDataFolder =
        Text.From(
            Excel.CurrentWorkbook(){[Name="cfg_RawDataFolder"]}[Content]{0}[Column1]
        ),
    SourceFile = RawDataFolder & "supply_chain_history.csv",
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