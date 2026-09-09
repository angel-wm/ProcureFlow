let
    Source = src_PartsMaster,

    RenamedColumns =
        Table.RenameColumns(
            Source,
            {
                {"part_id", "ProductID"},
                {"part_family", "PartFamily"},
                {"criticality_class", "CriticalityClass"},
                {"unit_cost", "UnitCost"},
                {"lead_time_days", "MasterLeadTimeDays"},
                {"supplier_id_primary", "PrimarySupplierID"},
                {"supplier_risk_class", "SupplierRiskClass"},
                {"is_repairable", "IsRepairable"},
                {"shelf_life_days", "ShelfLifeDays"}
            }
        ),

    ConvertedRepairable =
        Table.TransformColumns(
            RenamedColumns,
            {
                {
                    "IsRepairable",
                    each
                        if _ = "Yes" then true
                        else if _ = "No" then false
                        else error "Unexpected IsRepairable value",
                    type logical
                }
            }
        ),

    ChangedTypes =
        Table.TransformColumnTypes(
            ConvertedRepairable,
            {
                {"ProductID", type text},
                {"PartFamily", type text},
                {"CriticalityClass", type text},
                {"UnitCost", type number},
                {"MasterLeadTimeDays", Int64.Type},
                {"PrimarySupplierID", type text},
                {"SupplierRiskClass", type text},
                {"ShelfLifeDays", Int64.Type}
            },
            "en-US"
        )
in
    ChangedTypes