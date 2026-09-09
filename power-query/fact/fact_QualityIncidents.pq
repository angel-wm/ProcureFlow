let
    Source = stg_QualityIncidents,

    SelectedColumns =
        Table.SelectColumns(
            Source,
            {
                "QualityIncidentID",
                "IncidentDate",
                "ProductID",
                "SupplierID",
                "SiteID",
                "DefectSeverity",
                "DefectType",
                "ScrapQty"
            }
        )
in
    SelectedColumns