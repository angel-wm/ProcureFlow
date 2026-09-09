let
    Source = stg_Products,

    SelectedColumns =
        Table.SelectColumns(
            Source,
            {
                "ProductID",
                "PartFamily",
                "CriticalityClass",
                "UnitCost",
                "MasterLeadTimeDays",
                "PrimarySupplierID",
                "IsRepairable",
                "ShelfLifeDays"
            }
        )
in
    SelectedColumns