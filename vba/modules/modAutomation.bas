Attribute VB_Name = "modAutomation"
Option Explicit

Private Const STEP_VALIDATION As String = "VALIDATION"
Private Const STEP_BEGIN_ATTEMPT As String = "BEGIN_ATTEMPT"
Private Const STEP_POWER_QUERY As String = "POWER_QUERY"
Private Const STEP_EXCEL_CALCULATION As String = "EXCEL_CALCULATION"
Private Const STEP_PRELIMINARY_QC As String = "PRELIMINARY_QC"
Private Const STEP_PIVOT_REFRESH As String = "PIVOT_REFRESH"
Private Const STEP_POST_PIVOT_CALCULATION As String = "POST_PIVOT_CALCULATION"
Private Const STEP_POST_PIVOT_QC As String = "POST_PIVOT_QC"
Private Const STEP_FINALIZE As String = "FINALIZE"
Private Const STEP_FINAL_CALCULATION As String = "FINAL_CALCULATION"
Private Const STEP_FINAL_QC As String = "FINAL_QC"

Public Sub RefreshProcureFlow()

    Dim currentStep As String

    Dim preliminaryStatus As String
    Dim postPivotStatus As String
    Dim fullQualityStatus As String

    Dim automationStarted As Boolean
    Dim pivotRefreshCompleted As Boolean

    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean

    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    currentStep = STEP_VALIDATION

    Application.StatusBar = _
        "ProcureFlow: validating automation configuration..."

    ValidateAutomationStateConfiguration
    ValidatePowerQueryPipeline
    ValidateQualityControlConfiguration
    ValidatePivotRefreshConfiguration

    currentStep = STEP_BEGIN_ATTEMPT
    automationStarted = True

    Application.StatusBar = _
        "ProcureFlow: starting controlled refresh..."

    BeginAutomationAttempt

    currentStep = STEP_POWER_QUERY

    Application.StatusBar = _
        "ProcureFlow: refreshing Power Query pipeline..."

    RefreshPowerQueryPipeline

    currentStep = STEP_EXCEL_CALCULATION

    Application.StatusBar = _
        "ProcureFlow: calculating workbook..."

    Application.CalculateFull

    currentStep = STEP_PRELIMINARY_QC

    Application.StatusBar = _
        "ProcureFlow: evaluating preliminary Quality Control..."

    preliminaryStatus = GetPreliminaryQualityStatus()

    If preliminaryStatus = QUALITY_STATUS_FAIL Then

        Err.Raise _
            vbObjectError + 990, _
            "modAutomation.RefreshProcureFlow", _
            "Preliminary Quality Control returned FAIL."

    End If

    If preliminaryStatus <> QUALITY_STATUS_PASS _
       And preliminaryStatus <> QUALITY_STATUS_WARNING Then

        Err.Raise _
            vbObjectError + 991, _
            "modAutomation.RefreshProcureFlow", _
            "Preliminary Quality Control returned an unsupported status: '" & _
            preliminaryStatus & "'."

    End If

    currentStep = STEP_PIVOT_REFRESH

    Application.StatusBar = _
        "ProcureFlow: refreshing analysis PivotTables..."

    RefreshAnalysisPivotTables

    SetPivotRefreshStatus PIVOT_STATUS_PASS
    pivotRefreshCompleted = True

    currentStep = STEP_POST_PIVOT_CALCULATION

    Application.StatusBar = _
        "ProcureFlow: recalculating post-Pivot controls..."

    Application.CalculateFull

    currentStep = STEP_POST_PIVOT_QC

    Application.StatusBar = _
        "ProcureFlow: evaluating post-Pivot Quality Control..."

    postPivotStatus = GetPostPivotQualityStatus()

    If postPivotStatus = QUALITY_STATUS_FAIL Then

        Err.Raise _
            vbObjectError + 992, _
            "modAutomation.RefreshProcureFlow", _
            "Post-Pivot Quality Control returned FAIL."

    End If

    If postPivotStatus <> QUALITY_STATUS_PASS _
       And postPivotStatus <> QUALITY_STATUS_WARNING Then

        Err.Raise _
            vbObjectError + 993, _
            "modAutomation.RefreshProcureFlow", _
            "Post-Pivot Quality Control returned an unsupported status: '" & _
            postPivotStatus & "'."

    End If

    currentStep = STEP_FINALIZE

    Application.StatusBar = _
        "ProcureFlow: finalizing accepted refresh..."

    Select Case postPivotStatus

        Case QUALITY_STATUS_PASS
            CompleteAutomationAttempt AUTOMATION_STATUS_SUCCESS

        Case QUALITY_STATUS_WARNING
            CompleteAutomationAttempt AUTOMATION_STATUS_WARNING

        Case Else

            Err.Raise _
                vbObjectError + 996, _
                "modAutomation.RefreshProcureFlow", _
                "Cannot finalize automation from Quality Control status '" & _
                postPivotStatus & "'."

    End Select

    currentStep = STEP_FINAL_CALCULATION

    Application.StatusBar = _
        "ProcureFlow: calculating final Quality Control state..."

    Application.CalculateFull

    currentStep = STEP_FINAL_QC

    fullQualityStatus = GetFullQualityStatus()

    If fullQualityStatus = QUALITY_STATUS_FAIL Then

        Err.Raise _
            vbObjectError + 994, _
            "modAutomation.RefreshProcureFlow", _
            "Final Quality Control returned FAIL after automation completion."

    End If

    If fullQualityStatus <> QUALITY_STATUS_PASS _
       And fullQualityStatus <> QUALITY_STATUS_WARNING Then

        Err.Raise _
            vbObjectError + 995, _
            "modAutomation.RefreshProcureFlow", _
            "Final Quality Control returned an unsupported status: '" & _
            fullQualityStatus & "'."

    End If

    Application.StatusBar = False

    Application.ScreenUpdating = originalScreenUpdating
    Application.EnableEvents = originalEnableEvents

    MsgBox _
        "ProcureFlow refresh completed." & vbCrLf & vbCrLf & _
        "Quality Control: " & fullQualityStatus, _
        vbInformation, _
        "ProcureFlow"

    Exit Sub

ErrorHandler:

    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description

    On Error Resume Next

    If currentStep = STEP_PIVOT_REFRESH _
       And Not pivotRefreshCompleted Then

        SetPivotRefreshStatus PIVOT_STATUS_FAIL

    End If

    If automationStarted Then

        MarkAutomationFailure _
            currentStep, _
            originalErrorNumber, _
            originalErrorDescription

    End If

    Application.StatusBar = False

    Application.ScreenUpdating = originalScreenUpdating
    Application.EnableEvents = originalEnableEvents

    On Error GoTo 0

    MsgBox _
        "ProcureFlow refresh failed." & vbCrLf & vbCrLf & _
        "Step: " & currentStep & vbCrLf & _
        "Error: " & CStr(originalErrorNumber) & vbCrLf & _
        originalErrorDescription, _
        vbCritical, _
        "ProcureFlow"

End Sub