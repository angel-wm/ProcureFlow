let
    InventoryDates =
        Table.SelectColumns(
            fact_InventoryWeekly,
            {"WeekStartDate"}
        ),

    InventoryDatesRenamed =
        Table.RenameColumns(
            InventoryDates,
            {{"WeekStartDate", "Date"}}
        ),

    PurchaseOrderDates =
        Table.FromColumns(
            {
                List.Combine(
                    {
                        fact_PurchaseOrders[OrderDate],
                        fact_PurchaseOrders[PromisedDate],
                        fact_PurchaseOrders[ReceiptDate]
                    }
                )
            },
            {"Date"}
        ),

    QualityDates =
        Table.SelectColumns(
            fact_QualityIncidents,
            {"IncidentDate"}
        ),

    QualityDatesRenamed =
        Table.RenameColumns(
            QualityDates,
            {{"IncidentDate", "Date"}}
        ),

    AllFactDates =
        Table.Combine(
            {
                InventoryDatesRenamed,
                PurchaseOrderDates,
                QualityDatesRenamed
            }
        ),

    MinimumDate =
        List.Min(AllFactDates[Date]),

    MaximumDate =
        List.Max(AllFactDates[Date]),

    DateList =
        List.Dates(
            MinimumDate,
            Duration.Days(MaximumDate - MinimumDate) + 1,
            #duration(1, 0, 0, 0)
        ),

    DateTable =
        Table.FromList(
            DateList,
            Splitter.SplitByNothing(),
            {"Date"}
        ),

    ChangedDateType =
        Table.TransformColumnTypes(
            DateTable,
            {{"Date", type date}}
        ),

    AddedYear =
        Table.AddColumn(
            ChangedDateType,
            "Year",
            each Date.Year([Date]),
            Int64.Type
        ),

    AddedQuarter =
        Table.AddColumn(
            AddedYear,
            "Quarter",
            each "Q" & Text.From(Date.QuarterOfYear([Date])),
            type text
        ),

    AddedMonthNumber =
        Table.AddColumn(
            AddedQuarter,
            "MonthNumber",
            each Date.Month([Date]),
            Int64.Type
        ),

    AddedMonthName =
        Table.AddColumn(
            AddedMonthNumber,
            "MonthName",
            each Date.MonthName([Date], "en-US"),
            type text
        ),

    AddedYearMonth =
        Table.AddColumn(
            AddedMonthName,
            "YearMonth",
            each Date.ToText([Date], "yyyy-MM"),
            type text
        ),

    AddedISOYear =
        Table.AddColumn(
            AddedYearMonth,
            "ISOYear",
            each
                Date.Year(
                    Date.AddDays(
                        [Date],
                        3 - Date.DayOfWeek([Date], Day.Monday)
                    )
                ),
            Int64.Type
        ),

    AddedISOWeekNumber =
        Table.AddColumn(
            AddedISOYear,
            "ISOWeekNumber",
            each
                let
                    ISOYearValue = [ISOYear],
                    FirstISOWeekMonday =
                        Date.StartOfWeek(
                            #date(ISOYearValue, 1, 4),
                            Day.Monday
                        )
                in
                    Number.IntegerDivide(
                        Duration.Days([Date] - FirstISOWeekMonday),
                        7
                    ) + 1,
            Int64.Type
        ),

    AddedISOYearWeek =
        Table.AddColumn(
            AddedISOWeekNumber,
            "ISOYearWeek",
            each
                Text.From([ISOYear])
                & "-W"
                & Text.PadStart(
                    Text.From([ISOWeekNumber]),
                    2,
                    "0"
                ),
            type text
        ),

    AddedWeekStartDate =
        Table.AddColumn(
            AddedISOYearWeek,
            "WeekStartDate",
            each Date.StartOfWeek([Date], Day.Monday),
            type date
        ),

    AddedDayOfWeekNumber =
        Table.AddColumn(
            AddedWeekStartDate,
            "DayOfWeekNumber",
            each Date.DayOfWeek([Date], Day.Monday) + 1,
            Int64.Type
        ),

    AddedDayName =
        Table.AddColumn(
            AddedDayOfWeekNumber,
            "DayName",
            each Date.DayOfWeekName([Date], "en-US"),
            type text
        )
in
    AddedDayName