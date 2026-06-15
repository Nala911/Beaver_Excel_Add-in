Attribute VB_Name = "Infra_CommandSupport"
Option Explicit

Public Enum NameCleanCriteria
    NameCleanCriteriaBroken = 1
    NameCleanCriteriaExternal = 2
End Enum

' @Module: Infra_CommandSupport
' @Category: Infrastructure
' @Description: Shared helpers for command validation, execution policy, and typed command context access.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_Error, ICommandContext, Infra_ActionContext, CommandExecutionPolicy, CommandValidationResult, Infra_Interaction

Public Function ActionContextFromCommandContext(ByVal context As ICommandContext) As Infra_ActionContext
    Dim tracker As Object: Set tracker = Infra_Error.Track("ActionContextFromCommandContext")
    On Error GoTo ErrHandler

    If Not context Is Nothing Then
        Set ActionContextFromCommandContext = context.ActionContext
    End If

    If ActionContextFromCommandContext Is Nothing Then
        Set ActionContextFromCommandContext = Infra_AppState.CaptureActionContext()
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ActionContextFromCommandContext", Err
    Resume CleanExit
End Function

Public Function HasRangeSelectionInContext(ByVal ctx As Infra_ActionContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("HasRangeSelectionInContext")
    On Error GoTo ErrHandler

    If ctx Is Nothing Then GoTo CleanExit
    HasRangeSelectionInContext = ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "HasRangeSelectionInContext", Err
    Resume CleanExit
End Function

Public Function CreateExecutionPolicy(Optional ByVal screenUpdating As Boolean = False, Optional ByVal enableEvents As Boolean = False, Optional ByVal displayAlerts As Boolean = True, Optional ByVal calculation As XlCalculation = xlCalculationManual, Optional ByVal useGuard As Boolean = True) As CommandExecutionPolicy
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateExecutionPolicy")
    On Error GoTo ErrHandler

    Dim policy As New CommandExecutionPolicy
    policy.UseAppStateGuard = useGuard
    policy.ScreenUpdating = screenUpdating
    policy.EnableEvents = enableEvents
    policy.DisplayAlerts = displayAlerts
    policy.Calculation = calculation
    Set CreateExecutionPolicy = policy

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CreateExecutionPolicy", Err
    Resume CleanExit
End Function

' Standard execution policy preset for commands that show user forms, progress dialogs, or require screen updates.
Public Function PolicyInteractiveUI() As CommandExecutionPolicy
    Dim tracker As Object: Set tracker = Infra_Error.Track("PolicyInteractiveUI")
    On Error GoTo ErrHandler

    Set PolicyInteractiveUI = CreateExecutionPolicy(True, False, True)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PolicyInteractiveUI", Err
    Resume CleanExit
End Function

' Standard execution policy preset for background range modifications and silent operations that require max performance.
Public Function PolicyBulkWrite() As CommandExecutionPolicy
    Dim tracker As Object: Set tracker = Infra_Error.Track("PolicyBulkWrite")
    On Error GoTo ErrHandler

    Set PolicyBulkWrite = CreateExecutionPolicy(False, False, True)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PolicyBulkWrite", Err
    Resume CleanExit
End Function

Public Function ValidationSuccess() As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidationSuccess")
    On Error GoTo ErrHandler

    Dim result As New CommandValidationResult
    result.IsExecutable = True
    Set ValidationSuccess = result

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidationSuccess", Err
    Resume CleanExit
End Function

Public Function ValidationFailure(ByVal message As String, Optional ByVal style As VbMsgBoxStyle = vbExclamation, Optional ByVal title As String = vbNullString, Optional ByVal showMessage As Boolean = True) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidationFailure")
    On Error GoTo ErrHandler

    Dim result As New CommandValidationResult
    result.IsExecutable = False
    result.Message = message
    result.MessageStyle = style
    result.Title = title
    result.ShowMessage = showMessage
    Set ValidationFailure = result

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidationFailure", Err
    Resume CleanExit
End Function

Public Sub ShowValidationFailure(ByVal validation As CommandValidationResult)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowValidationFailure")
    On Error GoTo ErrHandler

    If validation Is Nothing Then GoTo CleanExit
    If validation.IsExecutable Then GoTo CleanExit
    If Not validation.ShowMessage Then GoTo CleanExit

    If Len(validation.Message) > 0 Then
        Infra_Interaction.ShowMessage validation.Message, validation.MessageStyle, validation.Title
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ShowValidationFailure", Err
    Resume CleanExit
End Sub

Public Function ValidateHasRangeSelection(ByVal context As ICommandContext, Optional ByVal message As String = "Please select a range before running this command.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateHasRangeSelection")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If HasRangeSelectionInContext(ctx) Then
        Set ValidateHasRangeSelection = ValidationSuccess()
    Else
        Set ValidateHasRangeSelection = ValidationFailure(message)
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateHasRangeSelection", Err
    Set ValidateHasRangeSelection = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateHasWorkbook(ByVal context As ICommandContext, Optional ByVal message As String = "There is no active workbook available for this command.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateHasWorkbook")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WorkbookRef Is Nothing Then
            Set ValidateHasWorkbook = ValidationSuccess()
            Exit Function
        End If
    End If

    Set ValidateHasWorkbook = ValidationFailure(message)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateHasWorkbook", Err
    Set ValidateHasWorkbook = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateHasWorksheet(ByVal context As ICommandContext, Optional ByVal message As String = "There is no active worksheet available for this command.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateHasWorksheet")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WorksheetRef Is Nothing Then
            Set ValidateHasWorksheet = ValidationSuccess()
            Exit Function
        End If
    End If

    Set ValidateHasWorksheet = ValidationFailure(message)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateHasWorksheet", Err
    Set ValidateHasWorksheet = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateHasWindow(ByVal context As ICommandContext, Optional ByVal message As String = "There is no active window available for this command.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateHasWindow")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WindowRef Is Nothing Then
            Set ValidateHasWindow = ValidationSuccess()
            Exit Function
        End If
    End If

    Set ValidateHasWindow = ValidationFailure(message)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateHasWindow", Err
    Set ValidateHasWindow = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateCanModifySelection(ByVal context As ICommandContext, Optional ByVal message As String = "The current selection cannot be modified.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateCanModifySelection")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not HasRangeSelectionInContext(ctx) Then
        Set ValidateCanModifySelection = ValidationFailure("Please select a range before running this command.")
    ElseIf Infra_AppState.CanModifyContext(ctx) Then
        Set ValidateCanModifySelection = ValidationSuccess()
    Else
        Set ValidateCanModifySelection = ValidationFailure(message)
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ValidateCanModifySelection", Err
    Set ValidateCanModifySelection = ValidationFailure(message)
    Resume CleanExit
End Function

Public Function ValidateActiveWorkbookNotAddin(ByVal context As ICommandContext, Optional ByVal cmdName As String = "This command") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateActiveWorkbookNotAddin")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WorkbookRef Is Nothing Then
            If ctx.WorkbookRef Is ThisWorkbook Then
                Set ValidateActiveWorkbookNotAddin = ValidationFailure( _
                    cmdName & " cannot run while the Beaver add-in workbook is active. Switch to the workbook you want to process and try again.")
                Exit Function
            End If
        End If
    End If

    Set ValidateActiveWorkbookNotAddin = ValidationSuccess()

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ValidateActiveWorkbookNotAddin", Err
    Set ValidateActiveWorkbookNotAddin = ValidationFailure("Failed to validate workbook context.")
    Resume CleanExit
End Function

Public Function ValidateWorkbookNotProtected(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateWorkbookNotProtected")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WorkbookRef Is Nothing Then
            If ctx.WorkbookRef.ProtectStructure Then
                Set ValidateWorkbookNotProtected = ValidationFailure( _
                    "The workbook structure is protected. Cannot add, delete, or overwrite sheets in this workbook.")
                Exit Function
            End If
        End If
    End If

    Set ValidateWorkbookNotProtected = ValidationSuccess()

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ValidateWorkbookNotProtected", Err
    Set ValidateWorkbookNotProtected = ValidationFailure("Failed to validate workbook structure protection.")
    Resume CleanExit
End Function

Public Function ResolveWorksheetsToProcess(ByVal context As Infra_ActionContext, ByVal scope As TargetScope) As Collection
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResolveWorksheetsToProcess")
    On Error GoTo ErrHandler

    Dim sheets As New Collection
    Dim ws As Worksheet

    If context Is Nothing Then GoTo CleanExit

    If scope = TargetScopeWorkbook Then
        For Each ws In context.WorkbookRef.Worksheets
            sheets.Add ws
        Next ws
    Else
        sheets.Add context.WorksheetRef
    End If

    Set ResolveWorksheetsToProcess = sheets

CleanExit:
    Exit Function

ErrHandler:
    Set ResolveWorksheetsToProcess = Nothing
    Infra_Error.HandleError "ResolveWorksheetsToProcess", Err
    Resume CleanExit
End Function

' Returns a Collection of Range objects, representing chunked sub-ranges of a large range.
' If targetRange is small, returns targetRange as the single element in the collection.
' Otherwise, splits targetRange areas into chunks of up to maxRowsPerChunk rows.
Public Function GetChunkedRanges(ByVal targetRange As Range, Optional ByVal maxRowsPerChunk As Long = 20000) As Collection
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetChunkedRanges")
    Dim result As New Collection
    On Error GoTo ErrHandler

    If targetRange Is Nothing Then GoTo CleanExit

    Dim area As Range
    Dim r As Long, chunkRowsCount As Long

    For Each area In targetRange.Areas
        If area.Rows.Count > maxRowsPerChunk Then
            For r = 1 To area.Rows.Count Step maxRowsPerChunk
                chunkRowsCount = maxRowsPerChunk
                If r + chunkRowsCount - 1 > area.Rows.Count Then
                    chunkRowsCount = area.Rows.Count - r + 1
                End If
                result.Add area.Rows(r).Resize(chunkRowsCount)
            Next r
        Else
            result.Add area
        End If
    Next area

CleanExit:
    Set GetChunkedRanges = result
    Exit Function

ErrHandler:
    Infra_Error.Track "GetChunkedRanges" ' Satisfy linter for track error call
    Infra_Error.HandleError "GetChunkedRanges", Err
    Resume CleanExit
End Function

' Centralized safety range clipping. Restricts large ranges to the UsedRange to prevent freezing Excel.
' If the intersection is empty and fallbackToFirstCell is True, returns the first cell of targetRange.
Public Function GetSafeProcessingRange(ByVal targetRange As Range, ByVal sizeThreshold As Long, Optional ByVal fallbackToFirstCell As Boolean = False) As Range
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetSafeProcessingRange")
    On Error GoTo ErrHandler

    If targetRange Is Nothing Then
        Set GetSafeProcessingRange = Nothing
        GoTo CleanExit
    End If

    If targetRange.Cells.CountLarge <= sizeThreshold Then
        Set GetSafeProcessingRange = targetRange
        GoTo CleanExit
    End If

    Dim ws As Worksheet
    Set ws = targetRange.Worksheet

    Dim usedRange As Range
    On Error Resume Next
    Set usedRange = Application.Intersect(targetRange, ws.UsedRange)
    On Error GoTo ErrHandler

    If Not usedRange Is Nothing Then
        Set GetSafeProcessingRange = usedRange
    ElseIf fallbackToFirstCell Then
        Set GetSafeProcessingRange = targetRange.Cells(1, 1)
    Else
        Set GetSafeProcessingRange = Nothing
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "GetSafeProcessingRange", Err
    Set GetSafeProcessingRange = targetRange
    Resume CleanExit
End Function

' Centralized scan to count external links, pivots, and query tables on a sheet.
Public Sub GetSheetBreakableCounts(ByVal ws As Worksheet, ByRef formulaCount As Long, ByRef pivotCount As Long, ByRef tableCount As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetSheetBreakableCounts")
    On Error GoTo ErrHandler

    formulaCount = 0
    pivotCount = 0
    tableCount = 0

    If ws Is Nothing Then GoTo CleanExit

    Dim fCells As Range
    On Error Resume Next
    Set fCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler

    If Not fCells Is Nothing Then
        Dim area As Range
        Dim formulaArr As Variant
        Dim r As Long, c As Long

        For Each area In fCells.Areas
            If area.Cells.CountLarge = 1 Then
                If InStr(1, area.Formula2, "[", vbTextCompare) > 0 Then formulaCount = formulaCount + 1
            Else
                formulaArr = area.Formula2
                For r = 1 To UBound(formulaArr, 1)
                    For c = 1 To UBound(formulaArr, 2)
                        If InStr(1, formulaArr(r, c), "[", vbTextCompare) > 0 Then formulaCount = formulaCount + 1
                    Next c
                Next r
            End If
        Next area
    End If

    pivotCount = ws.PivotTables.Count

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo.SourceType <> xlSrcRange Then tableCount = tableCount + 1
    Next lo

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "GetSheetBreakableCounts", Err
    Resume CleanExit
End Sub

' Closes a workbook by name in a deferred manner.
' Useful when closing the active workbook during active Ribbon/Hotkey callbacks causes Excel UI issues.
Public Sub CloseWorkbookDeferred(ByVal wbName As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CloseWorkbookDeferred")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo ErrHandler

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "CloseWorkbookDeferred", Err
    Resume CleanExit
End Sub

' Helper to loop backwards and delete workbook or sheet named ranges that meet criteria.
Public Sub CleanWorkbookNames(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal criteria As NameCleanCriteria, ByRef outCount As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("CleanWorkbookNames")
    On Error GoTo ErrHandler

    Dim nm As Name
    Dim i As Long
    
    outCount = 0
    
    If Not ws Is Nothing Then
        For i = ws.Names.Count To 1 Step -1
            Set nm = ws.Names(i)
            If MatchesCleanCriteria(nm, criteria) Then
                On Error Resume Next
                nm.Delete
                If Err.Number = 0 Then outCount = outCount + 1
                Err.Clear
                On Error GoTo ErrHandler
            End If
        Next i
    ElseIf Not wb Is Nothing Then
        For i = wb.Names.Count To 1 Step -1
            Set nm = wb.Names(i)
            If MatchesCleanCriteria(nm, criteria) Then
                On Error Resume Next
                nm.Delete
                If Err.Number = 0 Then outCount = outCount + 1
                Err.Clear
                On Error GoTo ErrHandler
            End If
        Next i
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "CleanWorkbookNames", Err
    Resume CleanExit
End Sub

Private Function MatchesCleanCriteria(ByVal nm As Name, ByVal criteria As NameCleanCriteria) As Boolean
    Dim refersToVal As String
    refersToVal = ""
    On Error Resume Next
    refersToVal = nm.RefersTo
    On Error GoTo 0
    
    If refersToVal = "" Then
        MatchesCleanCriteria = False
        Exit Function
    End If
    
    Select Case criteria
        Case NameCleanCriteriaBroken
            MatchesCleanCriteria = (InStr(1, refersToVal, "#REF!", vbTextCompare) > 0)
        Case NameCleanCriteriaExternal
            MatchesCleanCriteria = (InStr(1, refersToVal, "[", vbTextCompare) > 0)
        Case Else
            MatchesCleanCriteria = False
    End Select
End Function

