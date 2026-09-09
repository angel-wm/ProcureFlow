let
    Source = src_PurchaseOrders,

    RenamedColumns =
        Table.RenameColumns(
            Source,
            {
                {"po_id", "PurchaseOrderID"},
                {"supplier_id", "SupplierID"},
                {"site_id", "SiteID"},
                {"part_id", "ProductID"},
                {"order_date", "OrderDate"},
                {"promised_date", "PromisedDate"},
                {"receipt_date", "ReceiptDate"},
                {"ordered_qty", "OrderedQty"},
                {"received_qty", "ReceivedQty"}
            }
        ),

    ChangedTypes =
        Table.TransformColumnTypes(
            RenamedColumns,
            {
                {"PurchaseOrderID", type text},
                {"SupplierID", type text},
                {"SiteID", type text},
                {"ProductID", type text},
                {"OrderDate", type date},
                {"PromisedDate", type date},
                {"ReceiptDate", type date},
                {"OrderedQty", Int64.Type},
                {"ReceivedQty", Int64.Type}
            },
            "en-US"
        )
in
    ChangedTypes