Attribute VB_Name = "modQualityControl"
Option Explicit

Private Const CONTROL_SHEET_NAME As String = "02_CONTROL"
Private Const QUALITY_CONTROL_TABLE_NAME As String = "tblQualityControl"

Private Const COLUMN_CONTROL_ID As String = "ControlID"
Private Const COLUMN_SEVERITY As String = "Severity"
Private Const COLUMN_STATUS As String = "Status"

Private Const QC_LAST_SUCCESSFUL_REFRESH As String = "QC-030"
Private Const QC_PIVOT_REFRESH_STATUS As String = "QC-032"

Public Const QUALITY_STATUS_PASS As String = "PASS"
Public Const QUALITY_STATUS_WARNING As String = "WARNING"
Public Const QUALITY_STATUS_FAIL As String = "FAIL"
Public Const QUALITY_STATUS_NA As String = "N/A"

Public Sub ValidateQualityControlConfiguration()

    Dim qualityTable As ListObject
    Dim columnIndex As Long

    Set qualityTable = GetQualityControlTable()

    columnIndex = GetRequiredColumnIndex( _
        qualityTable, _
        COLUMN_CONTROL_ID _
    )

    columnIndex = GetRequiredColumnIndex( _
        qualityTable, _
        COLUMN_SEVERITY _
    )

    columnIndex = GetRequiredColumnIndex( _
        qualityTable, _
        COLUMN_STATUS _
    )

    ValidateRequiredControl _
        qualityTable, _
        QC_LAST_SUCCESSFUL_REFRESH

    ValidateRequiredControl _
        qualityTable, _
        QC_PIVOT_REFRESH_STATUS

End Sub

Public Function GetPreliminaryQualityStatus() As String

    ValidateQualityControlConfiguration

    GetPreliminaryQualityStatus = _
        EvaluateQualityStatus(True, True)

End Function

Public Function GetPostPivotQualityStatus() As String

    ValidateQualityControlConfiguration

    GetPostPivotQualityStatus = _
        EvaluateQualityStatus(True, False)

End Function

Public Function GetFullQualityStatus() As String

    ValidateQualityControlConfiguration

    GetFullQualityStatus = _
        EvaluateQualityStatus(False, False)

End Function

Private Function EvaluateQualityStatus( _
    ByVal excludeQC030 As Boolean, _
    ByVal excludeQC032 As Boolean _
) As String

    Dim qualityTable As ListObject
    Dim controlRow As ListRow

    Dim controlIdColumnIndex As Long
    Dim statusColumnIndex As Long

    Dim controlId As String
    Dim controlStatus As String

    Dim warningDetected As Boolean

    Set qualityTable = GetQualityControlTable()

    controlIdColumnIndex = GetRequiredColumnIndex( _
        qualityTable, _
        COLUMN_CONTROL_ID _
    )

    statusColumnIndex = GetRequiredColumnIndex( _
        qualityTable, _
        COLUMN_STATUS _
    )

    For Each controlRow In qualityTable.ListRows

        controlId = UCase$(Trim$(CStr( _
            controlRow.Range.Cells( _
                1, _
                controlIdColumnIndex _
            ).Value _
        )))

        If Not ShouldSkipControl( _
            controlId, _
            excludeQC030, _
            excludeQC032 _
        ) Then

            controlStatus = UCase$(Trim$(CStr( _
                controlRow.Range.Cells( _
                    1, _
                    statusColumnIndex _
                ).Value _
            )))

            Select Case controlStatus

                Case QUALITY_STATUS_FAIL

                    EvaluateQualityStatus = QUALITY_STATUS_FAIL
                    Exit Function

                Case QUALITY_STATUS_WARNING

                    warningDetected = True

                Case QUALITY_STATUS_PASS

                    ' No action required.

                Case QUALITY_STATUS_NA

                    ' Not-applicable controls do not affect precedence.

                Case vbNullString

                    warningDetected = True

                Case Else

                    warningDetected = True

            End Select

        End If

    Next controlRow

    If warningDetected Then
        EvaluateQualityStatus = QUALITY_STATUS_WARNING
    Else
        EvaluateQualityStatus = QUALITY_STATUS_PASS
    End If

End Function

Private Function ShouldSkipControl( _
    ByVal controlId As String, _
    ByVal excludeQC030 As Boolean, _
    ByVal excludeQC032 As Boolean _
) As Boolean

    If excludeQC030 Then

        If StrComp( _
            controlId, _
            QC_LAST_SUCCESSFUL_REFRESH, _
            vbTextCompare _
        ) = 0 Then

            ShouldSkipControl = True
            Exit Function

        End If

    End If

    If excludeQC032 Then

        If StrComp( _
            controlId, _
            QC_PIVOT_REFRESH_STATUS, _
            vbTextCompare _
        ) = 0 Then

            ShouldSkipControl = True
            Exit Function

        End If

    End If

    ShouldSkipControl = False

End Function

Private Function GetQualityControlTable() As ListObject

    On Error GoTo ErrorHandler

    Set GetQualityControlTable = _
        ThisWorkbook _
            .Worksheets(CONTROL_SHEET_NAME) _
            .ListObjects(QUALITY_CONTROL_TABLE_NAME)

    Exit Function

ErrorHandler:

    Err.Raise _
        vbObjectError + 970, _
        "modQualityControl.GetQualityControlTable", _
        "Required table '" & QUALITY_CONTROL_TABLE_NAME & _
        "' was not found on worksheet '" & _
        CONTROL_SHEET_NAME & "'."

End Function

Private Function GetRequiredColumnIndex( _
    ByVal qualityTable As ListObject, _
    ByVal columnName As String _
) As Long

    On Error GoTo ErrorHandler

    GetRequiredColumnIndex = _
        qualityTable.ListColumns(columnName).Index

    Exit Function

ErrorHandler:

    Err.Raise _
        vbObjectError + 971, _
        "modQualityControl.GetRequiredColumnIndex", _
        "Required column '" & columnName & _
        "' was not found in table '" & _
        qualityTable.Name & "'."

End Function

Private Sub ValidateRequiredControl( _
    ByVal qualityTable As ListObject, _
    ByVal requiredControlId As String _
)

    Dim controlRow As ListRow
    Dim controlIdColumnIndex As Long

    Dim currentControlId As String
    Dim matchCount As Long

    controlIdColumnIndex = GetRequiredColumnIndex( _
        qualityTable, _
        COLUMN_CONTROL_ID _
    )

    For Each controlRow In qualityTable.ListRows

        currentControlId = Trim$(CStr( _
            controlRow.Range.Cells( _
                1, _
                controlIdColumnIndex _
            ).Value _
        ))

        If StrComp( _
            currentControlId, _
            requiredControlId, _
            vbTextCompare _
        ) = 0 Then

            matchCount = matchCount + 1

        End If

    Next controlRow

    If matchCount <> 1 Then

        Err.Raise _
            vbObjectError + 972, _
            "modQualityControl.ValidateRequiredControl", _
            "Quality Control requires exactly one '" & _
            requiredControlId & _
            "' row, but found " & _
            CStr(matchCount) & "."

    End If

End Sub