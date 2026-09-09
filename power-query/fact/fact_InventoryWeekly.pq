let
    Source = stg_InventoryHistory,

    SelectedColumns =
        Table.SelectColumns(
            Source,
            {
                "WeekStartDate",
                "SiteID",
                "ProductID",
                "PlannedMaintenance",
                "ConsumptionQty",
                "OnHandQty",
                "BackorderQty",
                "BlockedQty",
                "ForecastQty",
                "ForecastType",
                "ForecastUpliftPct"
            }
        )
in
    SelectedColumns