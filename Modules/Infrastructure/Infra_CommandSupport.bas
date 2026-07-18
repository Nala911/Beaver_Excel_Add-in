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
' @Dependencies: Infra_AppState, Infra_Error, ICommandContext, ActionContext, CommandExecutionPolicy, CommandValidationResult, Infra_Interaction

Public Function ActionContextFromCommandContext(ByVal context As ICommandContext) As ActionContext
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

Public Function HasRangeSelectionInContext(ByVal ctx As ActionContext) As Boolean
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

    Dim ctx As ActionContext
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

    Dim ctx As ActionContext
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

    Dim ctx As ActionContext
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

    Dim ctx As ActionContext
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

    Dim ctx As ActionContext
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

    Dim ctx As ActionContext
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

    Dim ctx As ActionContext
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

Public Function ValidateWorksheetNotProtected(ByVal context As ICommandContext, Optional ByVal message As String = "The active worksheet is protected. Please unprotect the sheet and try again.") As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ValidateWorksheetNotProtected")
    On Error GoTo ErrHandler

    Dim ctx As ActionContext
    Set ctx = ActionContextFromCommandContext(context)

    If Not ctx Is Nothing Then
        If Not ctx.WorksheetRef Is Nothing Then
            If ctx.WorksheetRef.ProtectContents Then
                Set ValidateWorksheetNotProtected = ValidationFailure(message)
                Exit Function
            End If
        End If
    End If

    Set ValidateWorksheetNotProtected = ValidationSuccess()

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ValidateWorksheetNotProtected", Err
    Set ValidateWorksheetNotProtected = ValidationFailure("Failed to validate worksheet protection.")
    Resume CleanExit
End Function


Public Function ResolveWorksheetsToProcess(ByVal context As ActionContext, ByVal scope As TargetScope) As Collection
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
Public Function GetChunkedRanges(ByVal targetRange As Range, Optional ByVal maxRowsPerChunk As Long = 0) As Collection
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetChunkedRanges")
    Dim result As New Collection
    On Error GoTo ErrHandler

    If targetRange Is Nothing Then GoTo CleanExit

    Dim area As Range
    Dim r As Long, chunkRowsCount As Long
    
    Dim localMaxRows As Long
    If maxRowsPerChunk > 0 Then
        localMaxRows = maxRowsPerChunk
    Else
        localMaxRows = Infra_Config.CHUNK_ROWS_LIMIT
    End If

    Dim maxCellsPerChunk As Long
    maxCellsPerChunk = Infra_Config.CHUNK_CELLS_LIMIT
    
    Dim dynamicMaxRows As Long

    For Each area In targetRange.Areas
        Dim colCount As Long
        colCount = area.Columns.Count
        
        If colCount > 0 Then
            dynamicMaxRows = maxCellsPerChunk \ colCount
            If dynamicMaxRows < 1 Then dynamicMaxRows = 1
            If dynamicMaxRows > localMaxRows Then dynamicMaxRows = localMaxRows
        Else
            dynamicMaxRows = localMaxRows
        End If

        If area.Rows.Count > dynamicMaxRows Then
            For r = 1 To area.Rows.Count Step dynamicMaxRows
                chunkRowsCount = dynamicMaxRows
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

' Converts column indices (e.g. 28) to letters (e.g. "AB") purely in memory.
Public Function GetColLetter(ByVal colNum As Long) As String
    ' [Bypass Lint] PushContext "GetColLetter" / PopContext / Infra_Error.Track (exempt for CPU performance)
    Static cachedLetters(1 To 16384) As String
    
    If colNum < 1 Or colNum > 16384 Then Exit Function
    
    If cachedLetters(colNum) <> vbNullString Then
        GetColLetter = cachedLetters(colNum)
        Exit Function
    End If

    On Error GoTo ErrHandler

    Dim temp As Long
    temp = colNum
    Dim letter As String
    Do While temp > 0
        Dim remainder As Long
        remainder = (temp - 1) Mod 26
        letter = Chr$(65 + remainder) & letter
        temp = (temp - remainder) \ 26
    Loop
    cachedLetters(colNum) = letter
    GetColLetter = letter

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetColLetter", Err
    Resume CleanExit
End Function

' Concatenates column letter and row index to construct A1 address in memory.
Public Function GetA1Address(ByVal rowNum As Long, ByVal colNum As Long) As String
    ' [Bypass Lint] PushContext "GetA1Address" / PopContext / Infra_Error.Track (exempt for CPU performance)
    On Error GoTo ErrHandler

    GetA1Address = GetColLetter(colNum) & rowNum

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetA1Address", Err
    Resume CleanExit
End Function

' Parses refersTo formula like "='Sheet Name'!$A$1" or "=Sheet1!$B$5:$C$10"
' Returns True if successful, and outputs sheetName and rangeAddress.
' Bypasses COM calls to RefersToRange for performance.
Public Function ParseSimpleRefersTo(ByVal refersTo As String, ByRef outSheetName As String, ByRef outRangeAddress As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ParseSimpleRefersTo")
    On Error GoTo ErrHandler

    ParseSimpleRefersTo = False
    outSheetName = vbNullString
    outRangeAddress = vbNullString
    
    If Left$(refersTo, 1) <> "=" Then GoTo CleanExit
    
    Dim posExcl As Long
    posExcl = InStr(1, refersTo, "!")
    If posExcl <= 2 Then GoTo CleanExit
    
    If InStr(refersTo, "(") > 0 Or InStr(refersTo, "#REF!") > 0 Or InStr(refersTo, "[") > 0 Then GoTo CleanExit
    
    Dim sheetPart As String
    sheetPart = Mid$(refersTo, 2, posExcl - 2)
    
    If Left$(sheetPart, 1) = "'" And Right$(sheetPart, 1) = "'" Then
        If Len(sheetPart) > 2 Then
            sheetPart = Mid$(sheetPart, 2, Len(sheetPart) - 2)
        Else
            sheetPart = vbNullString
        End If
    End If
    
    sheetPart = Replace(sheetPart, "''", "'")
    
    Dim addrPart As String
    addrPart = Mid$(refersTo, posExcl + 1)
    
    If InStr(addrPart, "$") = 0 Then GoTo CleanExit
    
    Dim i As Long
    Dim c As String
    For i = 1 To Len(addrPart)
        c = Mid$(addrPart, i, 1)
        If Not (c = "$" Or c = ":" Or c = "," Or (c >= "0" And c <= "9") Or (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z")) Then
            GoTo CleanExit
        End If
    Next i
    
    outSheetName = sheetPart
    outRangeAddress = addrPart
    ParseSimpleRefersTo = True

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ParseSimpleRefersTo", Err
    Resume CleanExit
End Function


' Accumulates a cell or range address into a buffered Union range to optimize performance.
' Flushes the address buffer to Union when the string length exceeds 240 characters.
Public Function AccumulateUnion( _
    ByVal currentUnion As Range, _
    ByVal ws As Worksheet, _
    ByRef addrList As String, _
    ByVal cellAddr As String, _
    Optional ByVal forceFlush As Boolean = False) As Range
    
    ' [Bypass Lint] PushContext "AccumulateUnion" / PopContext / Infra_Error.Track (exempt for CPU performance)
    On Error GoTo ErrHandler

    Set AccumulateUnion = currentUnion

    If cellAddr <> vbNullString Then
        If Len(addrList) + Len(cellAddr) + 1 > 255 Then
            ' Flush current buffer
            If addrList <> vbNullString Then
                If AccumulateUnion Is Nothing Then
                    Set AccumulateUnion = ws.Range(addrList)
                Else
                    Set AccumulateUnion = Application.Union(AccumulateUnion, ws.Range(addrList))
                End If
            End If
            addrList = cellAddr
        Else
            If addrList = vbNullString Then
                addrList = cellAddr
            Else
                addrList = addrList & "," & cellAddr
            End If
        End If
    End If

    If forceFlush And addrList <> vbNullString Then
        If AccumulateUnion Is Nothing Then
            Set AccumulateUnion = ws.Range(addrList)
        Else
            Set AccumulateUnion = Application.Union(AccumulateUnion, ws.Range(addrList))
        End If
        addrList = vbNullString
    End If

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "AccumulateUnion", Err
    Resume CleanExit
End Function

Public Function ProcessRangeUnified( _
    ByVal targetRange As Range, _
    ByVal transformer As ICellTransformer, _
    Optional ByVal valueTypes As Long = 3) As Long ' 3 = xlTextValues Or xlNumbers
    
    Dim tracker As Object: Set tracker = Infra_Error.Track("ProcessRangeUnified")
    On Error GoTo ErrHandler

    If targetRange Is Nothing Then GoTo CleanExit
    If transformer Is Nothing Then GoTo CleanExit

    ' Intersect with UsedRange to prevent scanning millions of cells
    Dim safeRange As Range
    Set safeRange = GetSafeProcessingRange(targetRange, 100000, False)
    If safeRange Is Nothing Then GoTo CleanExit

    Dim targetCells As Range
    If safeRange.CountLarge = 1 Then
        If Not safeRange.HasFormula Then
            Dim vt As Integer
            vt = VarType(safeRange.Value)
            ' Check if matches expected types
            Dim matchesType As Boolean
            matchesType = False
            If (valueTypes And xlTextValues) And (vt = vbString) Then matchesType = True
            If (valueTypes And xlNumbers) And (vt = vbDouble Or vt = vbSingle Or vt = vbInteger Or vt = vbLong Or vt = vbDate) Then matchesType = True
            
            If matchesType Then Set targetCells = safeRange
        End If
    Else
        On Error Resume Next
        Set targetCells = safeRange.SpecialCells(xlCellTypeConstants, valueTypes)
        On Error GoTo ErrHandler
    End If

    If targetCells Is Nothing Then GoTo CleanExit

    ' Pre-fetch and cache column formats for all columns in safeRange to avoid redundant queries in chunks
    Dim colFormatCache As Object
    Set colFormatCache = CreateObject("Scripting.Dictionary")
    
    Dim colIdx As Long
    For colIdx = safeRange.Column To safeRange.Column + safeRange.Columns.Count - 1
        Dim colFmt As Variant
        On Error Resume Next
        colFmt = safeRange.Worksheet.Columns(colIdx).NumberFormat
        On Error GoTo 0
        If IsNull(colFmt) Then
            colFormatCache(colIdx) = "__MIXED__"
        Else
            colFormatCache(colIdx) = CStr(colFmt)
        End If
    Next colIdx

    Dim changeCount As Long
    
    ' Split into contiguous ranges and process in chunks
    Dim chunks As Collection
    Set chunks = GetChunkedRanges(targetCells, Infra_Config.CHUNK_ROWS_LIMIT)
    
    Dim chunkRange As Range
    For Each chunkRange In chunks
        If chunkRange.Cells.CountLarge = 1 Then
            Dim oldVal As Variant, newVal As Variant
            Dim oldFormat As String, newFormat As String
            oldVal = chunkRange.Value
            oldFormat = chunkRange.NumberFormat
            
            If transformer.TransformCell(oldVal, oldFormat, newVal, newFormat) Then
                If newFormat <> vbNullString And newFormat <> oldFormat Then
                    chunkRange.NumberFormat = newFormat
                End If
                chunkRange.Value = newVal
                changeCount = changeCount + 1
            End If
        Else
            changeCount = changeCount + ProcessChunkArea(chunkRange, transformer, colFormatCache)
        End If
        
        DoEvents
        If Infra_Progress.UserCancelled Then Exit For
    Next chunkRange

    ProcessRangeUnified = changeCount

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ProcessRangeUnified", Err
    Resume CleanExit
End Function

Private Function ProcessChunkArea(ByVal area As Range, ByVal transformer As ICellTransformer, ByVal colFormatCache As Object) As Long
    Dim vals As Variant
    Dim fmts As Variant
    vals = area.Value
    fmts = area.NumberFormat
    
    Dim isFmtsArray As Boolean: isFmtsArray = IsArray(fmts)
    Dim isFmtsNull As Boolean: isFmtsNull = IsNull(fmts)
    
    Dim hasChanged As Boolean
    Dim changeCount As Long
    
    Dim r As Long, c As Long
    Dim rowMin As Long, rowMax As Long
    Dim colMin As Long, colMax As Long
    
    rowMin = LBound(vals, 1)
    rowMax = UBound(vals, 1)
    colMin = LBound(vals, 2)
    colMax = UBound(vals, 2)
    
    Dim formatChangeCount As Long
    Dim dictRanges As Object
    Dim dictAddrLists As Object
    Set dictRanges = CreateObject("Scripting.Dictionary")
    Set dictAddrLists = CreateObject("Scripting.Dictionary")
    
    Dim isModifyData As Boolean
    isModifyData = (TypeName(transformer) = "FeatCmd_ModifyData")
    
    ' Resolve column formats using the cached column formats passed from the parent range
    Dim colFormats() As String
    If isFmtsNull Then
        ReDim colFormats(colMin To colMax)
        Dim colIdx As Long
        For colIdx = colMin To colMax
            Dim absCol As Long
            absCol = area.Column + colIdx - colMin
            If colFormatCache.Exists(absCol) Then
                colFormats(colIdx) = colFormatCache(absCol)
            Else
                colFormats(colIdx) = "__MIXED__"
            End If
        Next colIdx
    End If
    
    For r = rowMin To rowMax
        For c = colMin To colMax
            Dim oldVal As Variant
            Dim oldFormat As String
            oldVal = vals(r, c)
            
            If isFmtsArray Then
                oldFormat = CStr(fmts(r, c))
            ElseIf isFmtsNull Then
                Dim colFmtCached As String
                colFmtCached = colFormats(c)
                
                If colFmtCached <> "__MIXED__" Then
                    oldFormat = colFmtCached
                Else
                    Dim needsFormat As Boolean
                    needsFormat = False
                    If isModifyData Then
                        If VarType(oldVal) = vbDate Then needsFormat = True
                    End If
                    
                    If needsFormat Then
                        oldFormat = CStr(area.Cells(r, c).NumberFormat)
                    Else
                        oldFormat = "__MIXED__"
                    End If
                End If
            Else
                oldFormat = CStr(fmts)
            End If
            
            Dim newVal As Variant
            Dim newFormat As String
            
            If transformer.TransformCell(oldVal, oldFormat, newVal, newFormat) Then
                vals(r, c) = newVal
                changeCount = changeCount + 1
                hasChanged = True
                
                ' Resolve lazy/deferred format check if needed
                If newFormat = "__FORCE_GENERAL_IF_TEXT__" Then
                    Dim actualFmt As String
                    If isFmtsArray Then
                        actualFmt = CStr(fmts(r, c))
                    ElseIf isFmtsNull And colFormats(c) <> "__MIXED__" Then
                        actualFmt = colFormats(c)
                    ElseIf Not isFmtsArray And Not isFmtsNull Then
                        actualFmt = CStr(fmts)
                    Else
                        actualFmt = CStr(area.Cells(r, c).NumberFormat)
                    End If
                    
                    If InStr(1, actualFmt, "@") > 0 Or LCase$(actualFmt) = "text" Then
                        newFormat = "General"
                    Else
                        newFormat = vbNullString
                    End If
                End If
                
                If newFormat <> vbNullString And newFormat <> "__MIXED__" And (oldFormat = "__MIXED__" Or newFormat <> oldFormat) Then
                    formatChangeCount = formatChangeCount + 1
                    Dim cellAddr As String
                    cellAddr = GetA1Address(area.Row + r - 1, area.Column + c - 1)
                    
                    Dim curUnion As Range
                    Dim curAddrList As String
                    If dictRanges.Exists(newFormat) Then
                        Set curUnion = dictRanges(newFormat)
                        curAddrList = dictAddrLists(newFormat)
                    Else
                        Set curUnion = Nothing
                        curAddrList = vbNullString
                    End If
                    
                    Set curUnion = AccumulateUnion(curUnion, area.Parent, curAddrList, cellAddr)
                    Set dictRanges(newFormat) = curUnion
                    dictAddrLists(newFormat) = curAddrList
                End If
            End If
        Next c
    Next r
    
    If hasChanged Then
        ' Apply format changes efficiently in batches grouped by format
        If formatChangeCount > 0 Then
            Dim keyFormat As Variant
            For Each keyFormat In dictRanges.Keys
                Dim unionRange As Range
                Dim addrList As String
                Set unionRange = dictRanges(keyFormat)
                addrList = dictAddrLists(keyFormat)
                
                Set unionRange = AccumulateUnion(unionRange, area.Parent, addrList, "", True)
                If Not unionRange Is Nothing Then
                    unionRange.NumberFormat = keyFormat
                End If
            Next keyFormat
        End If
        
        area.Value = vals
    End If
    
    ProcessChunkArea = changeCount
End Function

' Finds the target bottom row index (1-indexed) to fill down from a source cell.
' Returns 0 if no target boundary is found or if the boundary does not exceed the source cell's row.
Public Function GetFillDownBoundary( _
    ByVal sourceCell As Range, _
    Optional ByVal maxSearchCol As Long = 0, _
    Optional ByVal cachedListObject As ListObject = Nothing, _
    Optional ByVal cachedCurrentRegion As Range = Nothing, _
    Optional ByVal cachedRowValues As Variant = Nothing, _
    Optional ByVal cachedRowStartCol As Long = 0 _
) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetFillDownBoundary")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = sourceCell.Worksheet
    Dim startRow As Long: startRow = sourceCell.Row
    Dim colIdx As Long: colIdx = sourceCell.Column
    
    Dim localMaxSearchCol As Long
    If maxSearchCol > 0 Then
        localMaxSearchCol = maxSearchCol
    Else
        localMaxSearchCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    End If
    
    Dim lastRow As Long: lastRow = 0
    Dim foundRef As Boolean: foundRef = False
    
    ' 1. ListObject (Table) Search
    Dim lo As ListObject
    If Not cachedListObject Is Nothing Then
        Set lo = cachedListObject
    Else
        On Error Resume Next
        Set lo = sourceCell.ListObject
        On Error GoTo ErrHandler
    End If
    
    If Not lo Is Nothing Then
        If Not Intersect(sourceCell, lo.DataBodyRange) Is Nothing Then
            lastRow = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
            foundRef = True
        End If
    End If
    
    ' 2. Proximity Search (closest column neighbor)
    If Not foundRef Then
        Dim maxProximityDist As Long: maxProximityDist = Infra_Config.MAX_FILL_PROXIMITY_COLUMNS
        Dim dist As Long
        Dim leftCol As Long, rightCol As Long
        
        Dim useCache As Boolean
        useCache = IsArray(cachedRowValues) And (cachedRowStartCol > 0)
        
        Dim nearRegion As Range
        Dim inCached As Boolean
        Dim candidateLastRow As Long
        
        For dist = 1 To maxProximityDist
            leftCol = colIdx - dist
            If leftCol >= 1 Then
                Dim leftIsEmpty As Boolean
                If useCache Then
                    Dim leftCacheIdx As Long
                    leftCacheIdx = leftCol - cachedRowStartCol + 1
                    If leftCacheIdx >= 1 And leftCacheIdx <= UBound(cachedRowValues, 2) Then
                        leftIsEmpty = IsEmpty(cachedRowValues(1, leftCacheIdx))
                    Else
                        leftIsEmpty = IsEmpty(ws.Cells(startRow, leftCol))
                    End If
                Else
                    leftIsEmpty = IsEmpty(ws.Cells(startRow, leftCol))
                End If
                
                If Not leftIsEmpty Then
                    If Not cachedCurrentRegion Is Nothing Then
                        inCached = (leftCol >= cachedCurrentRegion.Column) And _
                                   (leftCol <= cachedCurrentRegion.Column + cachedCurrentRegion.Columns.Count - 1)
                        If inCached Then
                            Set nearRegion = cachedCurrentRegion
                        Else
                            Set nearRegion = ws.Cells(startRow, leftCol).CurrentRegion
                        End If
                    Else
                        Set nearRegion = ws.Cells(startRow, leftCol).CurrentRegion
                    End If
                    candidateLastRow = nearRegion.Row + nearRegion.Rows.Count - 1
                    If candidateLastRow > startRow Then
                        lastRow = candidateLastRow
                        foundRef = True
                        Exit For
                    End If
                End If
            End If
            
            rightCol = colIdx + dist
            If rightCol <= localMaxSearchCol Then
                Dim rightIsEmpty As Boolean
                If useCache Then
                    Dim rightCacheIdx As Long
                    rightCacheIdx = rightCol - cachedRowStartCol + 1
                    If rightCacheIdx >= 1 And rightCacheIdx <= UBound(cachedRowValues, 2) Then
                        rightIsEmpty = IsEmpty(cachedRowValues(1, rightCacheIdx))
                    Else
                        rightIsEmpty = IsEmpty(ws.Cells(startRow, rightCol))
                    End If
                Else
                    rightIsEmpty = IsEmpty(ws.Cells(startRow, rightCol))
                End If
                
                If Not rightIsEmpty Then
                    If Not cachedCurrentRegion Is Nothing Then
                        inCached = (rightCol >= cachedCurrentRegion.Column) And _
                                   (rightCol <= cachedCurrentRegion.Column + cachedCurrentRegion.Columns.Count - 1)
                        If inCached Then
                            Set nearRegion = cachedCurrentRegion
                        Else
                            Set nearRegion = ws.Cells(startRow, rightCol).CurrentRegion
                        End If
                    Else
                        Set nearRegion = ws.Cells(startRow, rightCol).CurrentRegion
                    End If
                    candidateLastRow = nearRegion.Row + nearRegion.Rows.Count - 1
                    If candidateLastRow > startRow Then
                        lastRow = candidateLastRow
                        foundRef = True
                        Exit For
                    End If
                End If
            End If
        Next dist
    End If
    
    If foundRef And lastRow > startRow Then
        GetFillDownBoundary = lastRow
    Else
        GetFillDownBoundary = 0
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetFillDownBoundary", Err
    Resume CleanExit
End Function

' Finds the target rightmost column index (1-indexed) to fill rightward from a source cell.
' Returns 0 if no target boundary is found or if the boundary does not exceed the source cell's column.
Public Function GetFillRightBoundary( _
    ByVal sourceCell As Range, _
    Optional ByVal maxSearchRow As Long = 0, _
    Optional ByVal cachedListObject As ListObject = Nothing, _
    Optional ByVal cachedCurrentRegion As Range = Nothing, _
    Optional ByVal cachedColValues As Variant = Nothing, _
    Optional ByVal cachedColStartRow As Long = 0 _
) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetFillRightBoundary")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = sourceCell.Worksheet
    Dim startCol As Long: startCol = sourceCell.Column
    Dim rowIdx As Long: rowIdx = sourceCell.Row
    
    Dim localMaxSearchRow As Long
    If maxSearchRow > 0 Then
        localMaxSearchRow = maxSearchRow
    Else
        localMaxSearchRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    End If
    
    Dim lastCol As Long: lastCol = 0
    Dim foundRef As Boolean: foundRef = False
    
    ' 1. ListObject (Table) Search
    Dim lo As ListObject
    If Not cachedListObject Is Nothing Then
        Set lo = cachedListObject
    Else
        On Error Resume Next
        Set lo = sourceCell.ListObject
        On Error GoTo ErrHandler
    End If
    
    If Not lo Is Nothing Then
        If Not Intersect(sourceCell, lo.DataBodyRange) Is Nothing Then
            lastCol = lo.DataBodyRange.Column + lo.DataBodyRange.Columns.Count - 1
            foundRef = True
        End If
    End If
    
    If Not foundRef Then
        Dim maxProximityDist As Long: maxProximityDist = Infra_Config.MAX_FILL_PROXIMITY_COLUMNS
        Dim dist As Long
        Dim upRow As Long, downRow As Long
        
        Dim useCache As Boolean
        useCache = IsArray(cachedColValues) And (cachedColStartRow > 0)
        
        Dim nearRegion As Range
        Dim inCached As Boolean
        Dim candidateLastCol As Long
        
        For dist = 1 To maxProximityDist
            upRow = rowIdx - dist
            If upRow >= 1 Then
                Dim upIsEmpty As Boolean
                If useCache Then
                    Dim upCacheIdx As Long
                    upCacheIdx = upRow - cachedColStartRow + 1
                    If upCacheIdx >= 1 And upCacheIdx <= UBound(cachedColValues, 1) Then
                        upIsEmpty = IsEmpty(cachedColValues(upCacheIdx, 1))
                    Else
                        upIsEmpty = IsEmpty(ws.Cells(upRow, startCol))
                    End If
                Else
                    upIsEmpty = IsEmpty(ws.Cells(upRow, startCol))
                End If
                
                If Not upIsEmpty Then
                    If Not cachedCurrentRegion Is Nothing Then
                        inCached = (upRow >= cachedCurrentRegion.Row) And _
                                   (upRow <= cachedCurrentRegion.Row + cachedCurrentRegion.Rows.Count - 1)
                        If inCached Then
                            Set nearRegion = cachedCurrentRegion
                        Else
                            Set nearRegion = ws.Cells(upRow, startCol).CurrentRegion
                        End If
                    Else
                        Set nearRegion = ws.Cells(upRow, startCol).CurrentRegion
                    End If
                    candidateLastCol = nearRegion.Column + nearRegion.Columns.Count - 1
                    If candidateLastCol > startCol Then
                        lastCol = candidateLastCol
                        foundRef = True
                        Exit For
                    End If
                End If
            End If
            
            downRow = rowIdx + dist
            If downRow <= localMaxSearchRow Then
                Dim downIsEmpty As Boolean
                If useCache Then
                    Dim downCacheIdx As Long
                    downCacheIdx = downRow - cachedColStartRow + 1
                    If downCacheIdx >= 1 And downCacheIdx <= UBound(cachedColValues, 1) Then
                        downIsEmpty = IsEmpty(cachedColValues(downCacheIdx, 1))
                    Else
                        downIsEmpty = IsEmpty(ws.Cells(downRow, startCol))
                    End If
                Else
                    downIsEmpty = IsEmpty(ws.Cells(downRow, startCol))
                End If
                
                If Not downIsEmpty Then
                    If Not cachedCurrentRegion Is Nothing Then
                        inCached = (downRow >= cachedCurrentRegion.Row) And _
                                   (downRow <= cachedCurrentRegion.Row + cachedCurrentRegion.Rows.Count - 1)
                        If inCached Then
                            Set nearRegion = cachedCurrentRegion
                        Else
                            Set nearRegion = ws.Cells(downRow, startCol).CurrentRegion
                        End If
                    Else
                        Set nearRegion = ws.Cells(downRow, startCol).CurrentRegion
                    End If
                    candidateLastCol = nearRegion.Column + nearRegion.Columns.Count - 1
                    If candidateLastCol > startCol Then
                        lastCol = candidateLastCol
                        foundRef = True
                        Exit For
                    End If
                End If
            End If
        Next dist
    End If
    
    If foundRef And lastCol > startCol Then
        GetFillRightBoundary = lastCol
    Else
        GetFillRightBoundary = 0
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetFillRightBoundary", Err
    Resume CleanExit
End Function

' Checks if innerRange is geometrically inside outerRange.
Public Function IsRangeInsideRange(ByVal innerRange As Range, ByVal outerRange As Range) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("IsRangeInsideRange")
    On Error GoTo ErrHandler

    IsRangeInsideRange = False
    
    If innerRange Is Nothing Or outerRange Is Nothing Then Exit Function
    If innerRange.Worksheet.Parent.Name <> outerRange.Worksheet.Parent.Name Then Exit Function
    If innerRange.Worksheet.Name <> outerRange.Worksheet.Name Then Exit Function

    If (innerRange.Row >= outerRange.Row) And _
       (innerRange.Row + innerRange.Rows.Count - 1 <= outerRange.Row + outerRange.Rows.Count - 1) And _
       (innerRange.Column >= outerRange.Column) And _
       (innerRange.Column + innerRange.Columns.Count - 1 <= outerRange.Column + outerRange.Columns.Count - 1) Then
        IsRangeInsideRange = True
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "IsRangeInsideRange", Err
    Resume CleanExit
End Function



