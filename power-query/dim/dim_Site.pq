let
    Source = stg_InventoryHistory,

    SelectedColumns =
        Table.SelectColumns(
            Source,
            {"SiteID"}
        ),

    DistinctSites =
        Table.Distinct(
            SelectedColumns
        ),

    SortedSites =
        Table.Sort(
            DistinctSites,
            {{"SiteID", Order.Ascending}}
        )
in
    SortedSites