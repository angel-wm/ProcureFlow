Attribute VB_Name = "modPivotRefresh"
Option Explicit

Private Const EXPECTED_TOTAL_PIVOT_TABLES As Long = 10
Private Const EXPECTED_UNIQUE_PIVOT_CACHES As Long = 5

Public Sub ValidatePivotRefreshConfiguration()

    Dim sheetNames As Collection
    Dim sheetName As Variant

    Dim targetSheet As Worksheet
    Dim targetPivot As PivotTable

    Dim uniqueCacheIndexes As Collection

    Dim expectedSheetCount As Long
    Dim totalPivotTables As Long

    Set sheetNames = GetAnalysisPivotSheetNames()
    Set uniqueCacheIndexes = New Collection

    For Each sheetName In sheetNames

        Set targetSheet = GetRequiredWorksheet(CStr(sheetName))

        expectedSheetCount = _
            GetExpectedPivotTableCount(CStr(sheetName))

        If targetSheet.PivotTables.Count <> expectedSheetCount Then

            Err.Raise _
                vbObjectError + 980, _
                "modPivotRefresh.ValidatePivotRefreshConfiguration", _
                "Worksheet '" & CStr(sheetName) & _
                "' was expected to contain " & _
                CStr(expectedSheetCount) & _
                " PivotTables but contains " & _
                CStr(targetSheet.PivotTables.Count) & "."

        End If

        totalPivotTables = _
            totalPivotTables + targetSheet.PivotTables.Count

        For Each targetPivot In targetSheet.PivotTables

            If Not targetPivot.PivotCache.EnableRefresh Then

                Err.Raise _
                    vbObjectError + 981, _
                    "modPivotRefresh.ValidatePivotRefreshConfiguration", _
                    "PivotTable '" & targetPivot.Name & _
                    "' on worksheet '" & CStr(sheetName) & _
                    "' uses a PivotCache that cannot be refreshed."

            End If

            AddUniqueCacheIndex _
                uniqueCacheIndexes, _
                targetPivot.PivotCache.Index

        Next targetPivot

    Next sheetName

    If totalPivotTables <> EXPECTED_TOTAL_PIVOT_TABLES Then

        Err.Raise _
            vbObjectError + 982, _
            "modPivotRefresh.ValidatePivotRefreshConfiguration", _
            "Expected " & _
            CStr(EXPECTED_TOTAL_PIVOT_TABLES) & _
            " analysis PivotTables but found " & _
            CStr(totalPivotTables) & "."

    End If

    If uniqueCacheIndexes.Count <> EXPECTED_UNIQUE_PIVOT_CACHES Then

        Err.Raise _
            vbObjectError + 983, _
            "modPivotRefresh.ValidatePivotRefreshConfiguration", _
            "Expected " & _
            CStr(EXPECTED_UNIQUE_PIVOT_CACHES) & _
            " unique analysis PivotCaches but found " & _
            CStr(uniqueCacheIndexes.Count) & "."

    End If

End Sub

Public Sub RefreshAnalysisPivotTables()

    Dim sheetNames As Collection
    Dim uniqueCacheIndexes As Collection

    Dim sheetName As Variant
    Dim cacheIndex As Variant

    Dim targetSheet As Worksheet
    Dim targetPivot As PivotTable

    Dim cacheNumber As Long

    On Error GoTo ErrorHandler

    ValidatePivotRefreshConfiguration

    Set sheetNames = GetAnalysisPivotSheetNames()
    Set uniqueCacheIndexes = GetUniqueAnalysisCacheIndexes()

    cacheNumber = 0

    For Each cacheIndex In uniqueCacheIndexes

        cacheNumber = cacheNumber + 1

        Application.StatusBar = _
            "ProcureFlow: refreshing PivotCache " & _
            CStr(cacheNumber) & _
            " of " & _
            CStr(uniqueCacheIndexes.Count) & "..."

        ThisWorkbook _
            .PivotCaches(CLng(cacheIndex)) _
            .Refresh

    Next cacheIndex

    For Each sheetName In sheetNames

        Set targetSheet = _
            GetRequiredWorksheet(CStr(sheetName))

        For Each targetPivot In targetSheet.PivotTables

            Application.StatusBar = _
                "ProcureFlow: updating PivotTable '" & _
                targetPivot.Name & "'..."

            targetPivot.Update

        Next targetPivot

    Next sheetName

CleanExit:

    Application.StatusBar = False
    Exit Sub

ErrorHandler:

    Application.StatusBar = False

    Err.Raise _
        Err.Number, _
        "modPivotRefresh.RefreshAnalysisPivotTables", _
        Err.Description

End Sub

Private Function GetAnalysisPivotSheetNames() As Collection

    Dim sheetNames As Collection

    Set sheetNames = New Collection

    sheetNames.Add "30_PVT_Inventory"
    sheetNames.Add "31_PVT_Procurement"
    sheetNames.Add "32_PVT_Suppliers"

    Set GetAnalysisPivotSheetNames = sheetNames

End Function

Private Function GetExpectedPivotTableCount( _
    ByVal worksheetName As String _
) As Long

    Select Case worksheetName

        Case "30_PVT_Inventory"
            GetExpectedPivotTableCount = 3

        Case "31_PVT_Procurement"
            GetExpectedPivotTableCount = 4

        Case "32_PVT_Suppliers"
            GetExpectedPivotTableCount = 3

        Case Else

            Err.Raise _
                vbObjectError + 984, _
                "modPivotRefresh.GetExpectedPivotTableCount", _
                "No expected PivotTable count is defined for worksheet '" & _
                worksheetName & "'."

    End Select

End Function

Private Function GetRequiredWorksheet( _
    ByVal worksheetName As String _
) As Worksheet

    On Error GoTo ErrorHandler

    Set GetRequiredWorksheet = _
        ThisWorkbook.Worksheets(worksheetName)

    Exit Function

ErrorHandler:

    Err.Raise _
        vbObjectError + 985, _
        "modPivotRefresh.GetRequiredWorksheet", _
        "Required PivotTable worksheet '" & _
        worksheetName & _
        "' was not found."

End Function

Private Function GetUniqueAnalysisCacheIndexes() As Collection

    Dim sheetNames As Collection
    Dim uniqueCacheIndexes As Collection

    Dim sheetName As Variant

    Dim targetSheet As Worksheet
    Dim targetPivot As PivotTable

    Set sheetNames = GetAnalysisPivotSheetNames()
    Set uniqueCacheIndexes = New Collection

    For Each sheetName In sheetNames

        Set targetSheet = _
            GetRequiredWorksheet(CStr(sheetName))

        For Each targetPivot In targetSheet.PivotTables

            AddUniqueCacheIndex _
                uniqueCacheIndexes, _
                targetPivot.PivotCache.Index

        Next targetPivot

    Next sheetName

    Set GetUniqueAnalysisCacheIndexes = _
        uniqueCacheIndexes

End Function

Private Sub AddUniqueCacheIndex( _
    ByVal cacheIndexes As Collection, _
    ByVal cacheIndex As Long _
)

    Dim existingIndex As Variant

    For Each existingIndex In cacheIndexes

        If CLng(existingIndex) = cacheIndex Then
            Exit Sub
        End If

    Next existingIndex

    cacheIndexes.Add cacheIndex

End Sub