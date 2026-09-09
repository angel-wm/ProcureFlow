let
    Source = stg_Products,

    SelectedColumns =
        Table.SelectColumns(
            Source,
            {
                "PrimarySupplierID",
                "SupplierRiskClass"
            }
        ),

    RenamedColumns =
        Table.RenameColumns(
            SelectedColumns,
            {
                {"PrimarySupplierID", "SupplierID"}
            }
        ),

    DistinctSuppliers =
        Table.Distinct(
            RenamedColumns
        ),

    SortedSuppliers =
        Table.Sort(
            DistinctSuppliers,
            {{"SupplierID", Order.Ascending}}
        )
in
    SortedSuppliers