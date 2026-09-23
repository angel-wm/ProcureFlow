Attribute VB_Name = "modAutomationState"
Option Explicit

Private Const CONTROL_SHEET_NAME As String = "02_CONTROL"
Private Const AUTOMATION_STATE_TABLE_NAME As String = "tblAutomationState"

Private Const COLUMN_STATE_KEY As String = "StateKey"
Private Const COLUMN_STATE_VALUE As String = "StateValue"
Private Const COLUMN_UPDATED_AT As String = "UpdatedAt"
Private Const COLUMN_NOTES As String = "Notes"

Private Const STATE_WORKFLOW_STATUS As String = "WorkflowStatus"
Private Const STATE_LAST_ATTEMPT_STARTED As String = "LastAttemptStarted"
Private Const STATE_LAST_ATTEMPT_COMPLETED As String = "LastAttemptCompleted"
Private Const STATE_LAST_SUCCESSFUL_REFRESH As String = "LastSuccessfulRefresh"
Private Const STATE_PIVOT_REFRESH_STATUS As String = "PivotRefreshStatus"
Private Const STATE_FAILURE_STEP As String = "FailureStep"
Private Const STATE_LAST_ERROR_NUMBER As String = "LastErrorNumber"
Private Const STATE_LAST_ERROR_DESCRIPTION As String = "LastErrorDescription"

Public Const AUTOMATION_STATUS_NOT_RUN As String = "NOT_RUN"
Public Const AUTOMATION_STATUS_RUNNING As String = "RUNNING"
Public Const AUTOMATION_STATUS_SUCCESS As String = "SUCCESS"
Public Const AUTOMATION_STATUS_WARNING As String = "WARNING"
Public Const AUTOMATION_STATUS_FAILED As String = "FAILED"

Public Const PIVOT_STATUS_NOT_EVALUATED As String = "NOT_EVALUATED"
Public Const PIVOT_STATUS_PASS As String = "PASS"
Public Const PIVOT_STATUS_FAIL As String = "FAIL"

Public Sub ValidateAutomationStateConfiguration()

    Dim stateTable As ListObject
    Dim requiredKeys As Variant
    Dim key As Variant

    Set stateTable = GetAutomationStateTable()

    ValidateAutomationStateColumns stateTable

    requiredKeys = Array( _
        STATE_WORKFLOW_STATUS, _
        STATE_LAST_ATTEMPT_STARTED, _
        STATE_LAST_ATTEMPT_COMPLETED, _
        STATE_LAST_SUCCESSFUL_REFRESH, _
        STATE_PIVOT_REFRESH_STATUS, _
        STATE_FAILURE_STEP, _
        STATE_LAST_ERROR_NUMBER, _
        STATE_LAST_ERROR_DESCRIPTION _
    )

    For Each key In requiredKeys
        GetAutomationStateRow stateTable, CStr(key)
    Next key

End Sub

Public Function GetAutomationState(ByVal stateKey As String) As Variant

    Dim stateTable As ListObject
    Dim stateRow As ListRow
    Dim valueColumnIndex As Long

    Set stateTable = GetAutomationStateTable()
    Set stateRow = GetAutomationStateRow(stateTable, stateKey)

    valueColumnIndex = GetRequiredColumnIndex( _
        stateTable, _
        COLUMN_STATE_VALUE _
    )

    GetAutomationState = stateRow.Range.Cells( _
        1, _
        valueColumnIndex _
    ).Value

End Function

Public Sub BeginAutomationAttempt()

    Dim eventTime As Date

    ValidateAutomationStateConfiguration

    eventTime = Now

    WriteAutomationState _
        STATE_WORKFLOW_STATUS, _
        AUTOMATION_STATUS_RUNNING, _
        eventTime

    WriteAutomationState _
        STATE_LAST_ATTEMPT_STARTED, _
        eventTime, _
        eventTime

    WriteAutomationState _
        STATE_LAST_ATTEMPT_COMPLETED, _
        vbNullString, _
        eventTime, _
        True

    WriteAutomationState _
        STATE_PIVOT_REFRESH_STATUS, _
        PIVOT_STATUS_NOT_EVALUATED, _
        eventTime

    ClearAutomationFailure eventTime

End Sub

Public Sub SetPivotRefreshStatus(ByVal pivotStatus As String)

    Dim eventTime As Date

    ValidateAutomationStateConfiguration

    Select Case UCase$(Trim$(pivotStatus))

        Case PIVOT_STATUS_NOT_EVALUATED, _
             PIVOT_STATUS_PASS, _
             PIVOT_STATUS_FAIL

            eventTime = Now

            WriteAutomationState _
                STATE_PIVOT_REFRESH_STATUS, _
                UCase$(Trim$(pivotStatus)), _
                eventTime

        Case Else

            Err.Raise _
                vbObjectError + 903, _
                "modAutomationState.SetPivotRefreshStatus", _
                "Unsupported PivotRefreshStatus: '" & pivotStatus & "'."

    End Select

End Sub

Public Sub CompleteAutomationAttempt(ByVal finalStatus As String)

    Dim normalizedStatus As String
    Dim currentPivotStatus As String
    Dim eventTime As Date

    ValidateAutomationStateConfiguration

    normalizedStatus = UCase$(Trim$(finalStatus))

    If normalizedStatus <> AUTOMATION_STATUS_SUCCESS _
       And normalizedStatus <> AUTOMATION_STATUS_WARNING Then

        Err.Raise _
            vbObjectError + 904, _
            "modAutomationState.CompleteAutomationAttempt", _
            "Final automation status must be SUCCESS or WARNING."

    End If

    currentPivotStatus = UCase$(Trim$(CStr( _
        GetAutomationState(STATE_PIVOT_REFRESH_STATUS) _
    )))

    If currentPivotStatus <> PIVOT_STATUS_PASS Then

        Err.Raise _
            vbObjectError + 905, _
            "modAutomationState.CompleteAutomationAttempt", _
            "The automation workflow cannot be completed successfully because PivotRefreshStatus is not PASS."

    End If

    eventTime = Now

    WriteAutomationState _
        STATE_LAST_ATTEMPT_COMPLETED, _
        eventTime, _
        eventTime

    ClearAutomationFailure eventTime

    WriteAutomationState _
        STATE_LAST_SUCCESSFUL_REFRESH, _
        eventTime, _
        eventTime

    WriteAutomationState _
        STATE_WORKFLOW_STATUS, _
        normalizedStatus, _
        eventTime

End Sub

Public Sub MarkAutomationFailure( _
    ByVal failureStep As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String _
)

    Dim eventTime As Date

    ValidateAutomationStateConfiguration

    eventTime = Now

    WriteAutomationState _
        STATE_LAST_ATTEMPT_COMPLETED, _
        eventTime, _
        eventTime

    WriteAutomationState _
        STATE_FAILURE_STEP, _
        failureStep, _
        eventTime

    WriteAutomationState _
        STATE_LAST_ERROR_NUMBER, _
        errorNumber, _
        eventTime

    WriteAutomationState _
        STATE_LAST_ERROR_DESCRIPTION, _
        errorDescription, _
        eventTime

    WriteAutomationState _
        STATE_WORKFLOW_STATUS, _
        AUTOMATION_STATUS_FAILED, _
        eventTime

End Sub

Public Sub RestoreLastSuccessfulRefresh(ByVal previousValue As Variant)

    Dim eventTime As Date
    Dim previousText As String

    ValidateAutomationStateConfiguration

    eventTime = Now

    If IsEmpty(previousValue) Then

        WriteAutomationState _
            STATE_LAST_SUCCESSFUL_REFRESH, _
            vbNullString, _
            eventTime, _
            True

        Exit Sub

    End If

    previousText = Trim$(CStr(previousValue))

    If previousText = vbNullString Or previousText = "0" Then

        WriteAutomationState _
            STATE_LAST_SUCCESSFUL_REFRESH, _
            vbNullString, _
            eventTime, _
            True

    Else

        WriteAutomationState _
            STATE_LAST_SUCCESSFUL_REFRESH, _
            previousValue, _
            eventTime

    End If

End Sub

Private Function GetAutomationStateTable() As ListObject

    On Error GoTo ErrorHandler

    Set GetAutomationStateTable = _
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME) _
                    .ListObjects(AUTOMATION_STATE_TABLE_NAME)

    Exit Function

ErrorHandler:

    Err.Raise _
        vbObjectError + 900, _
        "modAutomationState.GetAutomationStateTable", _
        "Required table '" & AUTOMATION_STATE_TABLE_NAME & _
        "' was not found on worksheet '" & CONTROL_SHEET_NAME & "'."

End Function

Private Function GetAutomationStateRow( _
    ByVal stateTable As ListObject, _
    ByVal stateKey As String _
) As ListRow

    Dim rowItem As ListRow
    Dim keyColumnIndex As Long
    Dim currentKey As String

    keyColumnIndex = GetRequiredColumnIndex( _
        stateTable, _
        COLUMN_STATE_KEY _
    )

    For Each rowItem In stateTable.ListRows

        currentKey = Trim$(CStr( _
            rowItem.Range.Cells(1, keyColumnIndex).Value _
        ))

        If StrComp( _
            currentKey, _
            Trim$(stateKey), _
            vbTextCompare _
        ) = 0 Then

            Set GetAutomationStateRow = rowItem
            Exit Function

        End If

    Next rowItem

    Err.Raise _
        vbObjectError + 901, _
        "modAutomationState.GetAutomationStateRow", _
        "Required automation state key '" & stateKey & "' was not found."

End Function

Private Function GetRequiredColumnIndex( _
    ByVal stateTable As ListObject, _
    ByVal columnName As String _
) As Long

    On Error GoTo ErrorHandler

    GetRequiredColumnIndex = _
        stateTable.ListColumns(columnName).Index

    Exit Function

ErrorHandler:

    Err.Raise _
        vbObjectError + 902, _
        "modAutomationState.GetRequiredColumnIndex", _
        "Required column '" & columnName & _
        "' was not found in table '" & stateTable.Name & "'."

End Function

Private Sub ValidateAutomationStateColumns( _
    ByVal stateTable As ListObject _
)

    Dim columnIndex As Long

    columnIndex = GetRequiredColumnIndex( _
        stateTable, _
        COLUMN_STATE_KEY _
    )

    columnIndex = GetRequiredColumnIndex( _
        stateTable, _
        COLUMN_STATE_VALUE _
    )

    columnIndex = GetRequiredColumnIndex( _
        stateTable, _
        COLUMN_UPDATED_AT _
    )

    columnIndex = GetRequiredColumnIndex( _
        stateTable, _
        COLUMN_NOTES _
    )

End Sub

Private Sub WriteAutomationState( _
    ByVal stateKey As String, _
    ByVal stateValue As Variant, _
    ByVal eventTime As Date, _
    Optional ByVal clearValue As Boolean = False _
)

    Dim stateTable As ListObject
    Dim stateRow As ListRow
    Dim valueColumnIndex As Long
    Dim updatedAtColumnIndex As Long

    Set stateTable = GetAutomationStateTable()
    Set stateRow = GetAutomationStateRow(stateTable, stateKey)

    valueColumnIndex = GetRequiredColumnIndex( _
        stateTable, _
        COLUMN_STATE_VALUE _
    )

    updatedAtColumnIndex = GetRequiredColumnIndex( _
        stateTable, _
        COLUMN_UPDATED_AT _
    )

    If clearValue Then

        stateRow.Range.Cells( _
            1, _
            valueColumnIndex _
        ).ClearContents

    Else

        stateRow.Range.Cells( _
            1, _
            valueColumnIndex _
        ).Value = stateValue

    End If

    stateRow.Range.Cells( _
        1, _
        updatedAtColumnIndex _
    ).Value = eventTime

End Sub

Private Sub ClearAutomationFailure(ByVal eventTime As Date)

    WriteAutomationState _
        STATE_FAILURE_STEP, _
        vbNullString, _
        eventTime, _
        True

    WriteAutomationState _
        STATE_LAST_ERROR_NUMBER, _
        vbNullString, _
        eventTime, _
        True

    WriteAutomationState _
        STATE_LAST_ERROR_DESCRIPTION, _
        vbNullString, _
        eventTime, _
        True

End Sub