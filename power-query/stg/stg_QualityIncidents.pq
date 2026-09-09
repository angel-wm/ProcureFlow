let
    Source = src_QualityIncidents,

    RenamedColumns =
        Table.RenameColumns(
            Source,
            {
                {"incident_id", "QualityIncidentID"},
                {"incident_date", "IncidentDate"},
                {"part_id", "ProductID"},
                {"supplier_id", "SupplierID"},
                {"site_id", "SiteID"},
                {"defect_severity", "DefectSeverity"},
                {"defect_type", "DefectType"},
                {"scrap_qty", "ScrapQty"}
            }
        ),

    ChangedTypes =
        Table.TransformColumnTypes(
            RenamedColumns,
            {
                {"QualityIncidentID", type text},
                {"IncidentDate", type date},
                {"ProductID", type text},
                {"SupplierID", type text},
                {"SiteID", type text},
                {"DefectSeverity", type text},
                {"DefectType", type text},
                {"ScrapQty", Int64.Type}
            },
            "en-US"
        )
in
    ChangedTypes