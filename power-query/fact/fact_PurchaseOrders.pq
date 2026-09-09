let
    Source = stg_PurchaseOrders,

    SelectedColumns =
        Table.SelectColumns(
            Source,
            {
                "PurchaseOrderID",
                "SupplierID",
                "SiteID",
                "ProductID",
                "OrderDate",
                "PromisedDate",
                "ReceiptDate",
                "OrderedQty",
                "ReceivedQty"
            }
        ),

    AddedPromisedLeadTime =
        Table.AddColumn(
            SelectedColumns,
            "PromisedLeadTimeDays",
            each Duration.Days([PromisedDate] - [OrderDate]),
            Int64.Type
        ),

    AddedActualLeadTime =
        Table.AddColumn(
            AddedPromisedLeadTime,
            "ActualLeadTimeDays",
            each Duration.Days([ReceiptDate] - [OrderDate]),
            Int64.Type
        ),

    AddedLateReceipt =
        Table.AddColumn(
            AddedActualLeadTime,
            "IsLateReceipt",
            each [ReceiptDate] > [PromisedDate],
            type logical
        ),

    AddedPartialReceipt =
        Table.AddColumn(
            AddedLateReceipt,
            "IsPartialReceipt",
            each [ReceivedQty] < [OrderedQty],
            type logical
        )
in
    AddedPartialReceipt