Attribute VB_Name = "modVBAFoundations"
Option Explicit

Sub GoToControl()

    Dim controlSheet As Worksheet
    Dim sheetName As String
    Dim targetRow As Long
    Dim targetColumn As Long
    Dim navigationCompleted As Boolean

    On Error GoTo ErrorHandler

    sheetName = "02_CONTROL"
    targetRow = 1
    targetColumn = 1
    navigationCompleted = False

    If WorksheetExists(sheetName) Then

        Set controlSheet = ThisWorkbook.Worksheets(sheetName)

        controlSheet.Activate
        controlSheet.Cells(targetRow, targetColumn).Select

        navigationCompleted = True

    Else

        Debug.Print "Worksheet not found: " & sheetName

    End If

    If navigationCompleted Then
        PrintNavigationResult controlSheet.Name
    Else
        Debug.Print "Navigation did not complete."
    End If

    Exit Sub

ErrorHandler:

    Debug.Print "Procedure: GoToControl"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description

End Sub

Sub PrintNavigationResult(ByVal worksheetName As String)

    Debug.Print "Navigation completed successfully: " & worksheetName

End Sub

Function WorksheetExists(ByVal worksheetName As String) As Boolean

    Dim ws As Worksheet

    WorksheetExists = False

    For Each ws In ThisWorkbook.Worksheets

        If ws.Name = worksheetName Then
            WorksheetExists = True
            Exit For
        End If

    Next ws

End Function

Sub ExampleForLoop()

    Dim rowNumber As Long

    For rowNumber = 1 To 5
        Debug.Print "Row number: " & rowNumber
    Next rowNumber

End Sub

Sub ExampleForEachLoop()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "Worksheet: " & ws.Name
    Next ws

End Sub

Sub ExampleDoWhileLoop()

    Dim counter As Long

    counter = 1

    Do While counter <= 5
        Debug.Print "Counter: " & counter
        counter = counter + 1
    Loop

End Sub

Sub InspectReplenishmentTable()

    Dim ws As Worksheet
    Dim replenishmentTable As ListObject
    Dim tableColumn As ListColumn
    Dim firstRow As ListRow

    Set ws = ThisWorkbook.Worksheets("20_CALC_Replenishment")
    Set replenishmentTable = ws.ListObjects("tblReplenishment")
    Set firstRow = replenishmentTable.ListRows(1)

    Debug.Print "Worksheet: " & ws.Name
    Debug.Print "Table: " & replenishmentTable.Name
    Debug.Print "Rows: " & replenishmentTable.ListRows.Count
    Debug.Print "Columns: " & replenishmentTable.ListColumns.Count

    Debug.Print "Header range: " & replenishmentTable.HeaderRowRange.Address
    Debug.Print "Data range: " & replenishmentTable.DataBodyRange.Address
    Debug.Print "First table row range: " & firstRow.Range.Address

    Debug.Print "----- FIRST RECORD -----"

    Debug.Print _
        "ProductID: " & _
        firstRow.Range.Cells(1, replenishmentTable.ListColumns("ProductID").Index).Value

    Debug.Print _
        "SiteID: " & _
        firstRow.Range.Cells(1, replenishmentTable.ListColumns("SiteID").Index).Value

    Debug.Print _
        "InventoryStatus: " & _
        firstRow.Range.Cells(1, replenishmentTable.ListColumns("InventoryStatus").Index).Value

    Debug.Print _
        "RecommendedOrderQty: " & _
        firstRow.Range.Cells(1, replenishmentTable.ListColumns("RecommendedOrderQty").Index).Value

    Debug.Print "----- COLUMNS -----"

    For Each tableColumn In replenishmentTable.ListColumns
        Debug.Print tableColumn.Index & " - " & tableColumn.Name
    Next tableColumn

End Sub

