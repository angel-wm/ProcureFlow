Attribute VB_Name = "modVBAFoundations"
Option Explicit

' =============================================================================
' Module: modVBAFoundations
'
' Purpose:
'   Educational VBA module created during ProcureFlow Phase 8.
'
'   This module demonstrates foundational VBA concepts using real ProcureFlow
'   workbook objects and data structures. The procedures are learning examples
'   only and are not production automation.
'
'   This module will be removed from the operational workbook before Phase 8
'   is formally closed. Its exported source will remain in the repository as
'   learning evidence.
' =============================================================================


' -----------------------------------------------------------------------------
' Navigates to the ProcureFlow Quality Control worksheet.
'
' Demonstrates:
'   - Public procedures
'   - Worksheet object references
'   - Variables and data types
'   - Calling private procedures and functions
'   - Basic structured error handling
' -----------------------------------------------------------------------------
Public Sub GoToControl()

    Dim controlSheet As Worksheet
    Dim sheetName As String
    Dim targetRow As Long
    Dim targetColumn As Long
    Dim navigationCompleted As Boolean

    On Error GoTo ErrorHandler

    ' Define the navigation target.
    sheetName = "02_CONTROL"
    targetRow = 1
    targetColumn = 1
    navigationCompleted = False

    ' Validate that the target worksheet exists before using it.
    If WorksheetExists(sheetName) Then

        Set controlSheet = ThisWorkbook.Worksheets(sheetName)

        controlSheet.Activate
        controlSheet.Cells(targetRow, targetColumn).Select

        navigationCompleted = True

    Else

        Debug.Print "Worksheet not found: " & sheetName

    End If

    ' Report whether the navigation completed successfully.
    If navigationCompleted Then
        PrintNavigationResult controlSheet.Name
    Else
        Debug.Print "Navigation did not complete."
    End If

CleanExit:
    Exit Sub

ErrorHandler:

    ' Report runtime error information without hiding the failure.
    Debug.Print "Procedure: GoToControl"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description

    Resume CleanExit

End Sub


' -----------------------------------------------------------------------------
' Prints the result of a successful worksheet navigation.
'
' This is a private helper because it is not intended to be executed directly
' by the workbook user.
' -----------------------------------------------------------------------------
Private Sub PrintNavigationResult(ByVal worksheetName As String)

    Debug.Print "Navigation completed successfully: " & worksheetName

End Sub


' -----------------------------------------------------------------------------
' Returns True when the requested worksheet exists in ThisWorkbook.
'
' The function performs explicit validation instead of relying on hidden
' runtime errors to determine whether a worksheet exists.
' -----------------------------------------------------------------------------
Private Function WorksheetExists(ByVal worksheetName As String) As Boolean

    Dim ws As Worksheet

    WorksheetExists = False

    For Each ws In ThisWorkbook.Worksheets

        If ws.Name = worksheetName Then
            WorksheetExists = True
            Exit For
        End If

    Next ws

End Function


' -----------------------------------------------------------------------------
' Demonstrates a basic numeric For loop.
' -----------------------------------------------------------------------------
Public Sub ExampleForLoop()

    Dim rowNumber As Long

    For rowNumber = 1 To 5
        Debug.Print "Row number: " & rowNumber
    Next rowNumber

End Sub


' -----------------------------------------------------------------------------
' Demonstrates For Each by iterating through the Worksheets collection.
' -----------------------------------------------------------------------------
Public Sub ExampleForEachLoop()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "Worksheet: " & ws.Name
    Next ws

End Sub


' -----------------------------------------------------------------------------
' Demonstrates a Do While loop with an explicit termination condition.
'
' The counter is incremented on every iteration to ensure that the loop
' eventually terminates.
' -----------------------------------------------------------------------------
Public Sub ExampleDoWhileLoop()

    Dim counter As Long

    counter = 1

    Do While counter <= 5

        Debug.Print "Counter: " & counter
        counter = counter + 1

    Loop

End Sub


' -----------------------------------------------------------------------------
' Inspects the structure of tblReplenishment without modifying workbook data.
'
' Demonstrates:
'   - Worksheet and ListObject references
'   - ListRows, ListColumns, ListRow, and ListColumn
'   - HeaderRowRange and DataBodyRange
'   - Structured access to fields by column name
' -----------------------------------------------------------------------------
Public Sub InspectReplenishmentTable()

    Dim ws As Worksheet
    Dim replenishmentTable As ListObject
    Dim tableColumn As ListColumn
    Dim firstRow As ListRow

    Set ws = ThisWorkbook.Worksheets("20_CALC_Replenishment")
    Set replenishmentTable = ws.ListObjects("tblReplenishment")
    Set firstRow = replenishmentTable.ListRows(1)

    ' Report the main table structure.
    Debug.Print "Worksheet: " & ws.Name
    Debug.Print "Table: " & replenishmentTable.Name
    Debug.Print "Rows: " & replenishmentTable.ListRows.Count
    Debug.Print "Columns: " & replenishmentTable.ListColumns.Count

    ' Inspect the physical ranges that belong to the structured table.
    Debug.Print "Header range: " & replenishmentTable.HeaderRowRange.Address
    Debug.Print "Data range: " & replenishmentTable.DataBodyRange.Address
    Debug.Print "First table row range: " & firstRow.Range.Address

    Debug.Print "----- FIRST RECORD -----"

    ' Read business fields from the first record by structured column name.
    Debug.Print _
        "ProductID: " & _
        firstRow.Range.Cells( _
            1, _
            replenishmentTable.ListColumns("ProductID").Index _
        ).Value

    Debug.Print _
        "SiteID: " & _
        firstRow.Range.Cells( _
            1, _
            replenishmentTable.ListColumns("SiteID").Index _
        ).Value

    Debug.Print _
        "InventoryStatus: " & _
        firstRow.Range.Cells( _
            1, _
            replenishmentTable.ListColumns("InventoryStatus").Index _
        ).Value

    Debug.Print _
        "RecommendedOrderQty: " & _
        firstRow.Range.Cells( _
            1, _
            replenishmentTable.ListColumns("RecommendedOrderQty").Index _
        ).Value

    Debug.Print "----- COLUMNS -----"

    ' List every structured column currently defined in the table.
    For Each tableColumn In replenishmentTable.ListColumns

        Debug.Print _
            tableColumn.Index & _
            " - " & _
            tableColumn.Name

    Next tableColumn

End Sub


' -----------------------------------------------------------------------------
' Reads the first replenishment record and demonstrates conditional logic.
'
' Important:
'   The procedure does not calculate replenishment logic in VBA.
'   It only reads business outputs already calculated by the Excel model.
'
' The priority messages are educational examples and are not official
' ProcureFlow business rules.
' -----------------------------------------------------------------------------
Public Sub InspectFirstReplenishmentDecision()

    Dim ws As Worksheet
    Dim replenishmentTable As ListObject
    Dim firstRow As ListRow

    Dim inventoryStatus As String
    Dim recommendedOrderQty As Long

    Set ws = ThisWorkbook.Worksheets("20_CALC_Replenishment")
    Set replenishmentTable = ws.ListObjects("tblReplenishment")
    Set firstRow = replenishmentTable.ListRows(1)

    ' Read existing calculated business outputs from the first record.
    inventoryStatus = _
        firstRow.Range.Cells( _
            1, _
            replenishmentTable.ListColumns("InventoryStatus").Index _
        ).Value

    recommendedOrderQty = _
        firstRow.Range.Cells( _
            1, _
            replenishmentTable.ListColumns("RecommendedOrderQty").Index _
        ).Value

    Debug.Print "Inventory Status: " & inventoryStatus
    Debug.Print "Recommended Order Qty: " & recommendedOrderQty

    ' Determine whether the existing model recommends a purchase quantity.
    If recommendedOrderQty > 0 Then
        Debug.Print "Decision: Purchasing action is recommended."
    Else
        Debug.Print "Decision: No purchase recommendation."
    End If

    ' Demonstrate ElseIf using the existing inventory status.
    If inventoryStatus = "STOCKOUT" Then

        Debug.Print "Priority: Immediate attention."

    ElseIf inventoryStatus = "CRITICAL" Then

        Debug.Print "Priority: High."

    ElseIf inventoryStatus = "REORDER" Then

        Debug.Print "Priority: Replenishment review."

    Else

        Debug.Print "Priority: Standard review."

    End If

End Sub


' -----------------------------------------------------------------------------
' Reads a limited number of tblReplenishment records using a numeric For loop.
'
' Only the first five records are inspected so that the Immediate Window
' remains readable during the learning exercise.
' -----------------------------------------------------------------------------
Public Sub InspectReplenishmentRows()

    Dim ws As Worksheet
    Dim replenishmentTable As ListObject

    Dim rowNumber As Long
    Dim rowsToInspect As Long

    Dim productIDColumn As Long
    Dim siteIDColumn As Long
    Dim inventoryStatusColumn As Long
    Dim recommendedOrderQtyColumn As Long

    Set ws = ThisWorkbook.Worksheets("20_CALC_Replenishment")
    Set replenishmentTable = ws.ListObjects("tblReplenishment")

    ' Stop safely if the structured table contains no data records.
    If replenishmentTable.ListRows.Count = 0 Then

        Debug.Print "No replenishment records found."
        Exit Sub

    End If

    ' Resolve structured column positions once before entering the loop.
    productIDColumn = _
        replenishmentTable.ListColumns("ProductID").Index

    siteIDColumn = _
        replenishmentTable.ListColumns("SiteID").Index

    inventoryStatusColumn = _
        replenishmentTable.ListColumns("InventoryStatus").Index

    recommendedOrderQtyColumn = _
        replenishmentTable.ListColumns("RecommendedOrderQty").Index

    ' Limit the educational inspection to a maximum of five records.
    rowsToInspect = replenishmentTable.ListRows.Count

    If rowsToInspect > 5 Then
        rowsToInspect = 5
    End If

    Debug.Print "----- REPLENISHMENT RECORDS -----"

    For rowNumber = 1 To rowsToInspect

        Debug.Print _
            "Row " & rowNumber & _
            " | ProductID: " & _
            replenishmentTable.ListRows(rowNumber).Range.Cells( _
                1, _
                productIDColumn _
            ).Value & _
            " | SiteID: " & _
            replenishmentTable.ListRows(rowNumber).Range.Cells( _
                1, _
                siteIDColumn _
            ).Value & _
            " | Status: " & _
            replenishmentTable.ListRows(rowNumber).Range.Cells( _
                1, _
                inventoryStatusColumn _
            ).Value & _
            " | Recommended Qty: " & _
            replenishmentTable.ListRows(rowNumber).Range.Cells( _
                1, _
                recommendedOrderQtyColumn _
            ).Value

        ' Read the existing Excel result; do not recreate the calculation in VBA.
        If replenishmentTable.ListRows(rowNumber).Range.Cells( _
            1, _
            recommendedOrderQtyColumn _
        ).Value > 0 Then

            Debug.Print "    -> Purchase recommendation exists."

        End If

    Next rowNumber

End Sub


' -----------------------------------------------------------------------------
' Iterates through the columns of tblReplenishment using For Each.
'
' Two calculated business-output columns are highlighted for educational
' purposes without modifying their data or formulas.
' -----------------------------------------------------------------------------
Public Sub InspectReplenishmentColumns()

    Dim ws As Worksheet
    Dim replenishmentTable As ListObject
    Dim tableColumn As ListColumn

    Set ws = ThisWorkbook.Worksheets("20_CALC_Replenishment")
    Set replenishmentTable = ws.ListObjects("tblReplenishment")

    Debug.Print "----- REPLENISHMENT COLUMNS -----"
    Debug.Print "Total columns: " & replenishmentTable.ListColumns.Count

    For Each tableColumn In replenishmentTable.ListColumns

        If tableColumn.Name = "InventoryStatus" Or _
           tableColumn.Name = "RecommendedOrderQty" Then

            Debug.Print _
                "BUSINESS OUTPUT -> " & _
                tableColumn.Index & _
                " - " & _
                tableColumn.Name

        Else

            Debug.Print _
                tableColumn.Index & _
                " - " & _
                tableColumn.Name

        End If

    Next tableColumn

End Sub


' -----------------------------------------------------------------------------
' Demonstrates ByVal.
'
' Changing numberValue inside ChangeValueByVal does not modify originalValue
' in the calling procedure.
' -----------------------------------------------------------------------------
Public Sub ExampleByVal()

    Dim originalValue As Long

    originalValue = 10

    Debug.Print "Before procedure: " & originalValue

    ChangeValueByVal originalValue

    Debug.Print "After procedure: " & originalValue

End Sub


' -----------------------------------------------------------------------------
' Helper procedure used by ExampleByVal.
' -----------------------------------------------------------------------------
Private Sub ChangeValueByVal(ByVal numberValue As Long)

    numberValue = 50

    Debug.Print "Inside procedure: " & numberValue

End Sub


' -----------------------------------------------------------------------------
' Demonstrates ByRef.
'
' Changing numberValue inside ChangeValueByRef also changes originalValue
' in the calling procedure.
' -----------------------------------------------------------------------------
Public Sub ExampleByRef()

    Dim originalValue As Long

    originalValue = 10

    Debug.Print "Before procedure: " & originalValue

    ChangeValueByRef originalValue

    Debug.Print "After procedure: " & originalValue

End Sub


' -----------------------------------------------------------------------------
' Helper procedure used by ExampleByRef.
'
' ByRef is intentionally used here because modifying the caller variable is
' the specific behavior being demonstrated.
' -----------------------------------------------------------------------------
Private Sub ChangeValueByRef(ByRef numberValue As Long)

    numberValue = 50

    Debug.Print "Inside procedure: " & numberValue

End Sub


' -----------------------------------------------------------------------------
' Returns True when a named Excel Table exists on a named worksheet.
'
' Validation is explicit:
'   1. Confirm that the worksheet exists.
'   2. Iterate through its ListObjects collection.
'   3. Return True when the requested table is found.
'
' Runtime errors are not used as the normal validation mechanism.
' -----------------------------------------------------------------------------
Private Function TableExists( _
    ByVal worksheetName As String, _
    ByVal tableName As String _
) As Boolean

    Dim ws As Worksheet
    Dim tableObject As ListObject

    TableExists = False

    ' A table cannot exist on a worksheet that does not exist.
    If Not WorksheetExists(worksheetName) Then
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(worksheetName)

    For Each tableObject In ws.ListObjects

        If tableObject.Name = tableName Then

            TableExists = True
            Exit For

        End If

    Next tableObject

End Function


' -----------------------------------------------------------------------------
' Demonstrates validation of required ProcureFlow workbook objects.
'
' The public procedure coordinates the validation while private helper
' functions contain the reusable validation logic.
' -----------------------------------------------------------------------------
Public Sub ValidateReplenishmentObjects()

    Dim worksheetName As String
    Dim tableName As String

    worksheetName = "20_CALC_Replenishment"
    tableName = "tblReplenishment"

    If TableExists(worksheetName, tableName) Then

        Debug.Print _
            "Validation passed: " & _
            worksheetName & _
            " contains " & _
            tableName

    Else

        Debug.Print _
            "Validation failed: " & _
            tableName & _
            " was not found in " & _
            worksheetName

    End If

End Sub

