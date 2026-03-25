Attribute VB_Name = "Infra_Error"
Option Explicit

' @Module: Infra_Error
' @Category: Infrastructure
' @Description: Centralized error logging and breadcrumb tracking optimized for AI-assisted development.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Config, Infra_Diagnostics, Infra_OperationContext

Private pStack As Collection

Public Const CATEGORY_RUNTIME As String = "runtime"
Public Const CATEGORY_VALIDATION As String = "validation"
Public Const CATEGORY_TEST As String = "test"
Public Const SEVERITY_FATAL As String = "fatal"
Public Const SEVERITY_WARNING As String = "warning"

' --- PUBLIC PROCEDURES ---

' Factory method to create a new RAII context tracker.
Public Function Track(ByVal procName As String) As Infra_OperationContext
    Dim tracker As New Infra_OperationContext
    PushContext procName
    tracker.Init procName
    Set Track = tracker
End Function

' Pushes a procedure name onto the execution stack.
Public Sub PushContext(ByVal procName As String)
    If pStack Is Nothing Then Set pStack = New Collection
    pStack.Add procName
End Sub

' Pops the last procedure name from the execution stack.
Public Sub PopContext()
    On Error Resume Next
    If pStack Is Nothing Then Exit Sub
    If pStack.Count > 0 Then pStack.Remove pStack.Count
    On Error GoTo 0
End Sub

' The main entry point for all error handling.
' StateInfo: Optional Collection of "Key: Value" strings representing local variable states.
Public Sub HandleError(ByVal procedureName As String, ByRef errSource As ErrObject, Optional ByVal stateInfo As Collection = Nothing)
    HandleErrorDetailed procedureName, errSource, stateInfo, CATEGORY_RUNTIME, SEVERITY_FATAL
End Sub

Public Sub HandleErrorDetailed(ByVal procedureName As String, ByRef errSource As ErrObject, Optional ByVal stateInfo As Collection = Nothing, Optional ByVal category As String = CATEGORY_RUNTIME, Optional ByVal severity As String = SEVERITY_FATAL)
    On Error Resume Next
    
    Dim ctx As Infra_ErrorContext
    Dim errorMsg As String
    Set ctx = BuildErrorContext(procedureName, errSource, stateInfo, category, severity)
    errorMsg = BuildPromptString(ctx)
    
    ' 1. Log to Immediate Window (Developer Console)
    Debug.Print String(50, "-")
    Debug.Print "BEAVER ADD-IN - SYSTEM CRASH (" & Now() & ")"
    Debug.Print errorMsg
    Debug.Print String(50, "-")
    LogErrorContext ctx
    
    ' 2. Optional enrichers
    TryCopyToClipboard errorMsg
    TryShowErrorDialog errorMsg, ctx

    ' 3. Clear Stack
    Set pStack = Nothing
    
    ' 4. Failsafe Reset
    ResetApplicationState
    
    On Error GoTo 0
End Sub


' --- PRIVATE HELPERS ---

Private Function GetBreadcrumbs() As String
    If pStack Is Nothing Then
        GetBreadcrumbs = "Unknown"
        Exit Function
    End If
    
    Dim result As String
    Dim i As Long
    For i = 1 To pStack.Count
        result = result & IIf(i = 1, "", " -> ") & pStack(i)
    Next i
    GetBreadcrumbs = result
End Function

Public Function CurrentBreadcrumbs() As String
    CurrentBreadcrumbs = GetBreadcrumbs()
End Function

Public Function NewStateInfo() As Collection
    Set NewStateInfo = New Collection
End Function

Public Sub AddState(ByRef stateInfo As Collection, ByVal key As String, ByVal value As Variant)
    On Error Resume Next
    If stateInfo Is Nothing Then Set stateInfo = New Collection
    stateInfo.Add key & ": " & SafeVariantToString(value)
    On Error GoTo 0
End Sub

Private Function GetRichContext(Optional ByVal stateInfo As Collection = Nothing) As String
    On Error Resume Next
    Dim context As String
    context = ""
    
    ' State Info (Local Variables)
    If Not stateInfo Is Nothing Then
        context = context & "--- LOCAL STATE ---" & vbCrLf
        Dim item As Variant
        For Each item In stateInfo
            context = context & CStr(item) & vbCrLf
        Next item
        context = context & vbCrLf
    End If

    ' Application State
    context = context & "--- APP STATE ---" & vbCrLf
    context = context & "App.ScreenUpdating: " & CStr(Application.ScreenUpdating) & vbCrLf
    context = context & "App.EnableEvents: " & CStr(Application.EnableEvents) & vbCrLf
    
    Dim calcMode As String
    Select Case Application.Calculation
        Case -4105: calcMode = "Automatic"
        Case -4135: calcMode = "Manual"
        Case 2: calcMode = "SemiAutomatic"
        Case Else: calcMode = CStr(Application.Calculation)
    End Select
    context = context & "App.Calculation: " & calcMode & vbCrLf
    
    ' Workbook / Sheet
    If Not ActiveWorkbook Is Nothing Then
        context = context & "ActiveWorkbook: " & ActiveWorkbook.Name & vbCrLf
    End If
    If Not ActiveSheet Is Nothing Then
        context = context & "ActiveSheet: " & ActiveSheet.Name & vbCrLf
    End If
    
    ' Selection
    If TypeName(Selection) = "Range" Then
        Dim sel As Range
        Set sel = Selection
        context = context & "Selection.Address: " & sel.Address(False, False) & vbCrLf
        
        ' If single cell, get its value/formula to help AI
        If sel.Cells.Count = 1 Then
            Dim cellVal As Variant
            cellVal = sel.Value
            
            Dim cellFormula As String
            If sel.HasFormula Then
                cellFormula = sel.Formula
                context = context & "Selection.Formula: " & cellFormula & vbCrLf
            End If
            
            If IsError(cellVal) Then
                context = context & "Selection.Value: [Error " & CStr(cellVal) & "]" & vbCrLf
            ElseIf IsEmpty(cellVal) Then
                context = context & "Selection.Value: [Empty]" & vbCrLf
            Else
                ' Truncate long strings just in case
                Dim strVal As String
                strVal = CStr(cellVal)
                If Len(strVal) > 100 Then strVal = Left(strVal, 100) & "..."
                context = context & "Selection.Value: " & strVal & vbCrLf
            End If
        Else
            context = context & "Selection.CellsCount: " & sel.Cells.Count & vbCrLf
        End If
    Else
        context = context & "Selection.Type: " & TypeName(Selection) & vbCrLf
    End If
    
    GetRichContext = context
    On Error GoTo 0
End Function

Private Function BuildErrorContext(ByVal procedureName As String, ByRef errSource As ErrObject, Optional ByVal stateInfo As Collection = Nothing, Optional ByVal category As String = CATEGORY_RUNTIME, Optional ByVal severity As String = SEVERITY_FATAL) As Infra_ErrorContext
    Dim ctx As New Infra_ErrorContext
    Dim selectionRange As Range

    ctx.Timestamp = Now
    ctx.ProcedureName = procedureName
    ctx.Severity = severity
    ctx.Category = category
    ctx.ErrorNumber = errSource.Number
    ctx.ErrorDescription = errSource.Description
    ctx.ErrorSource = errSource.Source
    ctx.Username = Environ$("Username")
    ctx.AppVersion = Infra_Config.ADDIN_VERSION
    ctx.Breadcrumbs = GetBreadcrumbs()
    If ctx.Breadcrumbs = "Unknown" Then ctx.Breadcrumbs = procedureName

    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then ctx.WorkbookName = ActiveWorkbook.Name
    If Not ActiveSheet Is Nothing Then ctx.WorksheetName = ActiveSheet.Name
    If TypeName(Selection) = "Range" Then
        Set selectionRange = Selection
        If Not selectionRange Is Nothing Then
            ctx.SelectionAddress = selectionRange.Address(False, False)
        End If
    End If
    ctx.StateDump = FormatStateInfo(stateInfo)
    On Error GoTo 0

    Set BuildErrorContext = ctx
End Function

Private Function BuildPromptString(ByVal ctx As Infra_ErrorContext) As String
    Dim prompt As String
    
    prompt = "[SYSTEM CRASH]" & vbCrLf & _
             "Error Number: " & ctx.ErrorNumber & vbCrLf & _
             "Description: " & ctx.ErrorDescription & vbCrLf & _
             "Source: " & ctx.ErrorSource & vbCrLf & _
             "Severity: " & ctx.Severity & vbCrLf & _
             "Category: " & ctx.Category & vbCrLf & _
             "Path (Breadcrumbs): " & ctx.Breadcrumbs & vbCrLf & _
             "Addin Version: " & ctx.AppVersion & vbCrLf & _
             "Excel Version: " & Application.version & vbCrLf & vbCrLf & _
             "--- ENVIRONMENT SNAPSHOT ---" & vbCrLf & _
             GetRichContext(ParseStateInfoToCollection(ctx.StateDump)) & vbCrLf & _
             "Hey AI, the Beaver Add-in failed with the above error. " & _
             "Based on the Path and the Environment Snapshot, please analyze what might cause this and provide a fix."
             
    BuildPromptString = prompt
End Function

Private Function FormatStateInfo(Optional ByVal stateInfo As Collection = Nothing) As String
    On Error Resume Next
    Dim item As Variant
    Dim result As String

    If stateInfo Is Nothing Then Exit Function
    For Each item In stateInfo
        result = result & CStr(item) & vbCrLf
    Next item
    FormatStateInfo = result
    On Error GoTo 0
End Function

Private Function ParseStateInfoToCollection(ByVal stateDump As String) As Collection
    On Error Resume Next
    Dim lines() As String
    Dim i As Long
    Dim items As New Collection

    If stateDump = "" Then
        Set ParseStateInfoToCollection = items
        Exit Function
    End If

    lines = Split(stateDump, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" Then items.Add lines(i)
    Next i

    Set ParseStateInfoToCollection = items
    On Error GoTo 0
End Function

Private Function SanitizeForLog(ByVal value As String) As String
    SanitizeForLog = Replace(value, vbCrLf, " ")
    SanitizeForLog = Replace(SanitizeForLog, vbCr, " ")
    SanitizeForLog = Replace(SanitizeForLog, vbLf, " ")
    SanitizeForLog = Replace(SanitizeForLog, "|", "/")
End Function

Private Function SafeVariantToString(ByVal value As Variant) As String
    On Error Resume Next
    If IsObject(value) Then
        SafeVariantToString = TypeName(value)
    ElseIf IsError(value) Then
        SafeVariantToString = "#ERROR"
    ElseIf IsNull(value) Then
        SafeVariantToString = "[Null]"
    ElseIf IsEmpty(value) Then
        SafeVariantToString = "[Empty]"
    Else
        SafeVariantToString = CStr(value)
    End If
    On Error GoTo 0
End Function

Private Sub LogErrorContext(ByVal ctx As Infra_ErrorContext)
    If ctx Is Nothing Then Exit Sub
    Infra_Diagnostics.LogError ctx.ProcedureName, _
        "number=" & ctx.ErrorNumber & _
        " | description=" & SanitizeForLog(ctx.ErrorDescription) & _
        " | source=" & SanitizeForLog(ctx.ErrorSource) & _
        " | severity=" & SanitizeForLog(ctx.Severity) & _
        " | category=" & SanitizeForLog(ctx.Category) & _
        " | breadcrumbs=" & SanitizeForLog(ctx.Breadcrumbs)
End Sub

Private Sub ResetApplicationState()
    On Error Resume Next
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
    On Error GoTo 0
End Sub

Private Sub TryCopyToClipboard(ByVal text As String)
    On Error Resume Next
    CopyToClipboard text
    On Error GoTo 0
End Sub

Private Sub TryShowErrorDialog(ByVal errorMsg As String, ByVal ctx As Infra_ErrorContext)
    On Error Resume Next
    If ctx Is Nothing Then Exit Sub
    ShowAIErrorDialog errorMsg, ctx.Severity
    On Error GoTo 0
End Sub

Private Sub ShowAIErrorDialog(ByVal errorMsg As String, Optional ByVal severity As String = SEVERITY_FATAL)
    Dim displayMsg As String
    Dim style As VbMsgBoxStyle
    
    displayMsg = "🚨 ERROR CAUGHT!" & vbCrLf & vbCrLf & _
                 "✅ THE DETAILS HAVE BEEN AUTOMATICALLY COPIED TO YOUR CLIPBOARD." & vbCrLf & _
                 "Just switch to your AI chat and press Ctrl+V to paste the context." & vbCrLf & vbCrLf & _
                 "----- ERROR SUMMARY -----" & vbCrLf & _
                 "Check the Immediate Window (Ctrl+G) for the full log if needed." & vbCrLf & vbCrLf & _
                 errorMsg

    style = IIf(severity = SEVERITY_WARNING, vbExclamation, vbCritical)
    MsgBox displayMsg, style, Infra_Config.ADDIN_NAME & " Error - Copied to Clipboard!"
End Sub

Private Sub CopyToClipboard(ByVal text As String)
    On Error Resume Next
    
    ' Method 1: DataObject (Late bound, doesn't require MSForms reference)
    Dim objData As Object
    Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If Not objData Is Nothing Then
        objData.SetText text
        objData.PutInClipboard
        Exit Sub
    End If
    
    ' Method 2: HTMLFile (Fallback for Windows)
    CreateObject("htmlfile").ParentWindow.ClipboardData.SetData "text", text
    
    On Error GoTo 0
End Sub
