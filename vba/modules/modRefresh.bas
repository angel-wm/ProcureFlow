Attribute VB_Name = "modRefresh"
Option Explicit

Public Sub ValidatePowerQueryPipeline()

    Dim targets As Collection
    Dim target As Variant
    Dim targetQuery As QueryTable
    Dim targetConnection As WorkbookConnection

    Set targets = GetPowerQueryTargets()

    For Each target In targets

        Set targetQuery = GetTargetQueryTable( _
            CStr(target(0)), _
            CStr(target(1)) _
        )

        Set targetConnection = _
            ThisWorkbook.Connections(CStr(target(2)))

    Next target

End Sub

Public Sub RefreshPowerQueryPipeline()

    Dim targets As Collection
    Dim target As Variant

    ValidatePowerQueryPipeline

    Set targets = GetPowerQueryTargets()

    For Each target In targets

        RefreshPowerQueryTarget _
            CStr(target(0)), _
            CStr(target(1)), _
            CStr(target(2))

    Next target

End Sub

Private Function GetPowerQueryTargets() As Collection

    Dim targets As Collection

    Set targets = New Collection

    targets.Add Array( _
        "10_DATA_Products", _
        "tblProducts", _
        "Query - dim_Product" _
    )

    targets.Add Array( _
        "11_DATA_Sites", _
        "tblSites", _
        "Query - dim_Site" _
    )

    targets.Add Array( _
        "12_DATA_Suppliers", _
        "tblSuppliers", _
        "Query - dim_Supplier" _
    )

    targets.Add Array( _
        "16_DATA_Date", _
        "tblDate", _
        "Query - dim_Date" _
    )

    targets.Add Array( _
        "13_DATA_Inventory", _
        "tblInventoryHistory", _
        "Query - fact_InventoryWeekly" _
    )

    targets.Add Array( _
        "14_DATA_PurchaseOrders", _
        "tblPurchaseOrders", _
        "Query - fact_PurchaseOrders" _
    )

    targets.Add Array( _
        "15_DATA_Quality", _
        "tblQualityIncidents", _
        "Query - fact_QualityIncidents" _
    )

    targets.Add Array( _
        "02_CONTROL", _
        "tblQCPipelineHealth", _
        "Query - qc_PipelineHealth" _
    )

    Set GetPowerQueryTargets = targets

End Function

Private Function GetTargetQueryTable( _
    ByVal worksheetName As String, _
    ByVal tableName As String _
) As QueryTable

    On Error GoTo ErrorHandler

    Set GetTargetQueryTable = _
        ThisWorkbook _
            .Worksheets(worksheetName) _
            .ListObjects(tableName) _
            .QueryTable

    Exit Function

ErrorHandler:

    Err.Raise _
        vbObjectError + 962, _
        "modRefresh.GetTargetQueryTable", _
        "Required Power Query output table '" & _
        tableName & _
        "' was not found on worksheet '" & _
        worksheetName & "'."

End Function

Private Sub RefreshPowerQueryTarget( _
    ByVal worksheetName As String, _
    ByVal tableName As String, _
    ByVal connectionName As String _
)

    Dim targetQuery As QueryTable
    Dim refreshResult As Boolean

    On Error GoTo ErrorHandler

    Set targetQuery = GetTargetQueryTable( _
        worksheetName, _
        tableName _
    )

    If targetQuery.Refreshing Then

        Err.Raise _
            vbObjectError + 960, _
            "modRefresh.RefreshPowerQueryTarget", _
            "Query '" & connectionName & _
            "' was already refreshing before the controlled refresh began."

    End If

    Application.StatusBar = _
        "ProcureFlow: refreshing " & connectionName & "..."

    refreshResult = _
        targetQuery.Refresh(BackgroundQuery:=False)

    If Not refreshResult Then

        Err.Raise _
            vbObjectError + 961, _
            "modRefresh.RefreshPowerQueryTarget", _
            "Query '" & connectionName & _
            "' did not report a successful refresh."

    End If

    Exit Sub

ErrorHandler:

    Err.Raise _
        Err.Number, _
        "modRefresh.RefreshPowerQueryTarget [" & _
        connectionName & "]", _
        Err.Description

End Sub