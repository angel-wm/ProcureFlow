let
    Source = src_SupplyChainHistory,

    RenamedColumns =
        Table.RenameColumns(
            Source,
            {
                {"date", "WeekStartDate"},
                {"site_id", "SiteID"},
                {"part_id", "ProductID"},
                {"planned_maintenance", "PlannedMaintenance"},
                {"consumption_qty", "ConsumptionQty"},
                {"on_hand_qty", "OnHandQty"},
                {"backorder_qty", "BackorderQty"},
                {"blocked_qty", "BlockedQty"},
                {"forecast_qty", "ForecastQty"},
                {"forecast_type", "ForecastType"},
                {"forecast_uplift_pct", "ForecastUpliftPct"}
            }
        ),

    ConvertedMaintenance =
        Table.TransformColumns(
            RenamedColumns,
            {
                {
                    "PlannedMaintenance",
                    each
                        if Text.Lower(Text.From(_)) = "true" then true
                        else if Text.Lower(Text.From(_)) = "false" then false
                        else error "Unexpected PlannedMaintenance value",
                    type logical
                }
            }
        ),

    ChangedTypes =
        Table.TransformColumnTypes(
            ConvertedMaintenance,
            {
                {"WeekStartDate", type date},
                {"SiteID", type text},
                {"ProductID", type text},
                {"ConsumptionQty", Int64.Type},
                {"OnHandQty", Int64.Type},
                {"BackorderQty", Int64.Type},
                {"BlockedQty", Int64.Type},
                {"ForecastQty", Int64.Type},
                {"ForecastType", type text},
                {"ForecastUpliftPct", type number}
            },
            "en-US"
        )
in
    ChangedTypes