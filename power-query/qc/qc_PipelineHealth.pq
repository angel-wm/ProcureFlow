let
    RawDataFolder =
        Text.From(
            Excel.CurrentWorkbook(){[Name="cfg_RawDataFolder"]}[Content]{0}[Column1]
        ),

    SourceDefinitions = {
        [
            FileKey = "Products",
            FileName = "parts_master.csv",
            RequiredColumns = {
                "part_id",
                "part_family",
                "criticality_class",
                "unit_cost",
                "lead_time_days",
                "supplier_id_primary",
                "supplier_risk_class",
                "is_repairable",
                "shelf_life_days"
            }
        ],
        [
            FileKey = "Inventory",
            FileName = "supply_chain_history.csv",
            RequiredColumns = {
                "date",
                "site_id",
                "part_id",
                "planned_maintenance",
                "consumption_qty",
                "on_hand_qty",
                "backorder_qty",
                "blocked_qty",
                "forecast_qty",
                "forecast_type",
                "forecast_uplift_pct"
            }
        ],
        [
            FileKey = "PurchaseOrders",
            FileName = "purchase_orders.csv",
            RequiredColumns = {
                "po_id",
                "supplier_id",
                "site_id",
                "part_id",
                "order_date",
                "promised_date",
                "receipt_date",
                "ordered_qty",
                "received_qty"
            }
        ],
        [
            FileKey = "Quality",
            FileName = "quality_incidents.csv",
            RequiredColumns = {
                "incident_id",
                "incident_date",
                "part_id",
                "supplier_id",
                "site_id",
                "defect_severity",
                "defect_type",
                "scrap_qty"
            }
        ]
    },

    ReadSource = (Definition as record) as record =>
        let
            FullPath = RawDataFolder & Definition[FileName],

            CsvAttempt =
                try
                    Csv.Document(
                        File.Contents(FullPath),
                        [
                            Delimiter = ",",
                            Encoding = 65001,
                            QuoteStyle = QuoteStyle.Csv
                        ]
                    ),

            PromoteAttempt =
                if CsvAttempt[HasError] then
                    null
                else
                    try
                        Table.PromoteHeaders(
                            CsvAttempt[Value],
                            [PromoteAllScalars = true]
                        ),

            ReadError =
                if CsvAttempt[HasError] then
                    true
                else
                    PromoteAttempt[HasError],

            Data =
                if ReadError then
                    null
                else
                    PromoteAttempt[Value],

            MissingColumns =
                if Data = null then
                    Definition[RequiredColumns]
                else
                    List.Difference(
                        Definition[RequiredColumns],
                        Table.ColumnNames(Data)
                    ),

            SourceRowCount =
                if Data = null then
                    null
                else
                    Table.RowCount(Data)
        in
            [
                FileKey = Definition[FileKey],
                FileName = Definition[FileName],
                Available = not ReadError,
                MissingColumns = MissingColumns,
                SourceRowCount = SourceRowCount
            ],

    SourceRecords =
        List.Transform(
            SourceDefinitions,
            each ReadSource(_)
        ),

    GetSourceRecord = (FileKey as text) as record =>
        List.First(
            List.Select(
                SourceRecords,
                each [FileKey] = FileKey
            )
        ),

    SourceFileCount = List.Count(SourceRecords),

    SourceAvailabilityExceptions =
        List.Count(
            List.Select(
                SourceRecords,
                each [Available] = false
            )
        ),

    SourceAvailableCount =
        SourceFileCount - SourceAvailabilityExceptions,

    MissingColumnCount =
        List.Sum(
            List.Transform(
                SourceRecords,
                each List.Count([MissingColumns])
            )
        ),

    StagingAttempts = {
        try stg_Products,
        try stg_InventoryHistory,
        try stg_PurchaseOrders,
        try stg_QualityIncidents
    },

    ErrorRowCount = (Attempt as record) as number =>
        if Attempt[HasError] then
            1
        else
            try
                Table.RowCount(
                    Table.SelectRowsWithErrors(Attempt[Value])
                )
            otherwise
                1,

    StagingErrorRows =
        List.Sum(
            List.Transform(
                StagingAttempts,
                each ErrorRowCount(_)
            )
        ),

    ProductsSourceRows =
        GetSourceRecord("Products")[SourceRowCount],

    InventorySourceRows =
        GetSourceRecord("Inventory")[SourceRowCount],

    PurchaseOrdersSourceRows =
        GetSourceRecord("PurchaseOrders")[SourceRowCount],

    QualitySourceRows =
        GetSourceRecord("Quality")[SourceRowCount],

    ProductsLoadedAttempt = try Table.RowCount(dim_Product),
    InventoryLoadedAttempt = try Table.RowCount(fact_InventoryWeekly),
    PurchaseOrdersLoadedAttempt = try Table.RowCount(fact_PurchaseOrders),
    QualityLoadedAttempt = try Table.RowCount(fact_QualityIncidents),

    SafeCount = (Attempt as record) as nullable number =>
        if Attempt[HasError] then
            null
        else
            Attempt[Value],

    ProductsLoadedRows =
        SafeCount(ProductsLoadedAttempt),

    InventoryLoadedRows =
        SafeCount(InventoryLoadedAttempt),

    PurchaseOrdersLoadedRows =
        SafeCount(PurchaseOrdersLoadedAttempt),

    QualityLoadedRows =
        SafeCount(QualityLoadedAttempt),

    RowDifference = (
        SourceRows as nullable number,
        LoadedRows as nullable number
    ) as number =>
        if SourceRows = null or LoadedRows = null then
            1
        else
            Number.Abs(SourceRows - LoadedRows),

    RowCountExceptions =
        RowDifference(ProductsSourceRows, ProductsLoadedRows)
        + RowDifference(InventorySourceRows, InventoryLoadedRows)
        + RowDifference(PurchaseOrdersSourceRows, PurchaseOrdersLoadedRows)
        + RowDifference(QualitySourceRows, QualityLoadedRows),

    CountText = (Value as nullable number) as text =>
        if Value = null then
            "ERROR"
        else
            Number.ToText(Value, "0", "en-US"),

    RowCountResult =
        "Products "
        & CountText(ProductsSourceRows)
        & "/"
        & CountText(ProductsLoadedRows)
        & " | Inventory "
        & CountText(InventorySourceRows)
        & "/"
        & CountText(InventoryLoadedRows)
        & " | PO "
        & CountText(PurchaseOrdersSourceRows)
        & "/"
        & CountText(PurchaseOrdersLoadedRows)
        & " | Quality "
        & CountText(QualitySourceRows)
        & "/"
        & CountText(QualityLoadedRows),

    QueryExecutionAttempts = {
        try Table.RowCount(src_PartsMaster),
        try Table.RowCount(src_SupplyChainHistory),
        try Table.RowCount(src_PurchaseOrders),
        try Table.RowCount(src_QualityIncidents),

        try Table.RowCount(stg_Products),
        try Table.RowCount(stg_InventoryHistory),
        try Table.RowCount(stg_PurchaseOrders),
        try Table.RowCount(stg_QualityIncidents),

        try Table.RowCount(dim_Product),
        try Table.RowCount(dim_Site),
        try Table.RowCount(dim_Supplier),
        try Table.RowCount(dim_Date),

        try Table.RowCount(fact_InventoryWeekly),
        try Table.RowCount(fact_PurchaseOrders),
        try Table.RowCount(fact_QualityIncidents)
    },

    PowerQueryErrorCount =
        List.Count(
            List.Select(
                QueryExecutionAttempts,
                each [HasError]
            )
        ),

    TechnicalExceptionTotal =
        SourceAvailabilityExceptions
        + MissingColumnCount
        + StagingErrorRows
        + RowCountExceptions
        + PowerQueryErrorCount,

    EvaluationTimestamp =
        DateTime.LocalNow(),

    Output =
        Table.FromRecords({
            [
                ControlID = "QC-001",
                Result =
                    Text.From(SourceAvailableCount)
                    & " / "
                    & Text.From(SourceFileCount)
                    & " available",
                Expected = "4 / 4 available",
                ExceptionCount = SourceAvailabilityExceptions
            ],
            [
                ControlID = "QC-002",
                Result =
                    if MissingColumnCount = 0 then
                        "All required columns present"
                    else
                        Text.From(MissingColumnCount)
                        & " required column(s) missing",
                Expected = "All required columns present",
                ExceptionCount = MissingColumnCount
            ],
            [
                ControlID = "QC-003",
                Result =
                    if StagingErrorRows = 0 then
                        "No staging/type errors"
                    else
                        Text.From(StagingErrorRows)
                        & " staging error row(s)",
                Expected = "No staging/type errors",
                ExceptionCount = StagingErrorRows
            ],
            [
                ControlID = "QC-004",
                Result = RowCountResult,
                Expected = "Source rows = loaded rows",
                ExceptionCount = RowCountExceptions
            ],
            [
                ControlID = "QC-030",
                Result =
                    if TechnicalExceptionTotal = 0 then
                        DateTime.ToText(
                            EvaluationTimestamp,
                            "yyyy-MM-dd HH:mm:ss"
                        )
                    else
                        "Current refresh not valid",
                Expected = "Successful current refresh",
                ExceptionCount =
                    if TechnicalExceptionTotal = 0 then 0 else 1
            ],
            [
                ControlID = "QC-031",
                Result =
                    if PowerQueryErrorCount = 0 then
                        "No query execution errors"
                    else
                        Text.From(PowerQueryErrorCount)
                        & " query execution error(s)",
                Expected = "No query execution errors",
                ExceptionCount = PowerQueryErrorCount
            ]
        })
in
    Output
