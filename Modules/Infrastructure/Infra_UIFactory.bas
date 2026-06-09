Attribute VB_Name = "Infra_UIFactory"
Option Explicit

' @Module: Infra_UIFactory
' @Category: Infrastructure
' @Description: Centralized factory for creating and displaying standardized user prompts.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Config, Infra_ExportRequest, Infra_ScopedRequest, Infra_ActionContext

' Shows the Clean Data options via UserForm picker and returns a populated Request object.
Public Function ShowCleanDataDialog(ByVal ctx As Infra_ActionContext) As Infra_CleanDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowCleanDataDialog")
    On Error GoTo ErrHandler
    
    Dim promptMsg As String
    Dim normalizedChoice As String
    Dim request As Infra_CleanDataRequest
    Dim options As Variant
    Dim defaultChoice As String
    Dim hasSelection As Boolean
    Dim confirmMsg As String

    hasSelection = HasUsableSelection(ctx)
    If hasSelection Then
        options = BuildChoiceArray("Range", "Sheet", "Workbook")
        defaultChoice = "Range"
    Else
        options = BuildChoiceArray("Sheet", "Workbook")
        defaultChoice = "Sheet"
    End If

    promptMsg = "Clean text with TRIM and CLEAN." & vbCrLf & vbCrLf & _
                BuildCompactContextSummary(ctx, hasSelection) & vbCrLf & vbCrLf & _
                "Scope:" & vbCrLf & _
                "Sheet - Active sheet" & vbCrLf & _
                "Workbook - All sheets"
    If hasSelection Then
        promptMsg = promptMsg & vbCrLf & "Range - Current selection"
    Else
        promptMsg = promptMsg & vbCrLf & vbCrLf & "No selection is required for Sheet or Workbook scope."
    End If

    confirmMsg = "Workbook-wide Clean Data updates every sheet and cannot be restored as a single workbook-wide undo action." & vbCrLf & vbCrLf & _
                 "Continue with workbook-wide cleaning?"

    If Not PromptForScopeSelection(ctx, "Clean Data", promptMsg, defaultChoice, options, confirmMsg, normalizedChoice) Then GoTo CleanExit

    If normalizedChoice = "R" Or normalizedChoice = "RANGE" Or normalizedChoice = "SELECTED" Or normalizedChoice = "SELECTION" Then
        If Not hasSelection Then
            Infra_Interaction.ShowWarning "Select a range first if you want to clean only the current selection.", BuildDialogTitle("Clean Data")
            GoTo CleanExit
        End If
    End If

    Set request = CreateCleanDataRequest(ctx, normalizedChoice)
    If request Is Nothing Then GoTo CleanExit

    Dim cleanOptionsList As Variant
    Dim cleanDefaultsChecked As Variant
    Dim selectedIndices As Variant
    Dim idx As Variant

    cleanOptionsList = Array( _
        "Trim extra spaces & non-breaking spaces", _
        "Remove non-printable characters", _
        "Convert text-formatted numbers to real numbers", _
        "Delete broken named ranges (#REF!)", _
        "Highlight inconsistent formulas (yellow)" _
    )
    cleanDefaultsChecked = Array(True, True, True, True, True)

    If Not Infra_Interaction.PromptMultiOption( _
        "Select the data cleaning options to apply:", _
        "Clean Data Options", _
        cleanOptionsList, _
        cleanDefaultsChecked, _
        selectedIndices) Then
        GoTo CleanExit
    End If

    request.CleanTrimSpaces = False
    request.CleanNonPrintables = False
    request.CleanTextNumbers = False
    request.CleanBrokenNames = False
    request.CleanInconsistentFormulas = False

    For Each idx In selectedIndices
        Select Case idx
            Case 0: request.CleanTrimSpaces = True
            Case 1: request.CleanNonPrintables = True
            Case 2: request.CleanTextNumbers = True
            Case 3: request.CleanBrokenNames = True
            Case 4: request.CleanInconsistentFormulas = True
        End Select
    Next idx

    Set ShowCleanDataDialog = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowCleanDataDialog", Err
    Resume CleanExit
End Function

' Shows the Export options via UserForm picker and returns a populated Request object.
Public Function ShowExportDialog(ByVal ctx As Infra_ActionContext) As Infra_ExportRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowExportDialog")
    On Error GoTo ErrHandler
    
    Dim request As Infra_ExportRequest
    Dim exportChoice As String
    Dim normalizedChoice As String
    
    Set request = New Infra_ExportRequest
    Set request.Context = ctx
    Set request.SourceRange = ResolveExportRange(ctx)
    request.ScaleFactor = Infra_Config.DEFAULT_EXPORT_SCALE
    
    If request.SourceRange Is Nothing Then
        Infra_Interaction.ShowWarning "No data found on the active sheet to export."
        GoTo CleanExit
    End If

    Do
        If Not ShowOptionPicker( _
            "Export the selected content and choose where to save it." & vbCrLf & vbCrLf & _
            BuildExportSummary(request.SourceRange) & vbCrLf & vbCrLf & _
            "Choose a format:" & vbCrLf & _
            "PNG - High-resolution image" & vbCrLf & _
            "PDF - Print-ready document" & vbCrLf & vbCrLf & _
            "Choose PNG or PDF.", _
            BuildDialogTitle("Export"), "PNG", Array("PNG", "PDF"), exportChoice) Then GoTo CleanExit

        normalizedChoice = NormalizeChoiceText(exportChoice)
        Select Case normalizedChoice
            Case "", "PNG", "IMAGE"
                request.ExportAsPng = True
                Exit Do
            Case "PDF"
                request.ExportAsPng = False
                Exit Do
            Case Else
                Infra_Interaction.ShowWarning "Please choose PNG or PDF.", BuildDialogTitle("Export")
        End Select
    Loop

    If request.ExportAsPng Then
        request.ScaleFactor = PromptForExportScale(request.ScaleFactor)
        If request.ScaleFactor = 0 Then GoTo CleanExit
    End If

    request.OutputPath = PromptForOutputPath( _
        ctx, _
        "Export", _
        BuildSuggestedExportBaseName(request.SourceRange, request.ExportAsPng), _
        IIf(request.ExportAsPng, "png", "pdf"), _
        IIf(request.ExportAsPng, "PNG Files (*.png), *.png", "PDF Files (*.pdf), *.pdf"))
    If request.OutputPath = vbNullString Then GoTo CleanExit
    
    Set ShowExportDialog = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowExportDialog", Err
    Resume CleanExit
End Function

' Shows the conversion scope dialog for formula-to-value actions using a UserForm picker.
Public Function ShowStaticConversionDialog(ByVal ctx As Infra_ActionContext) As Infra_ScopedRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowStaticConversionDialog")
    On Error GoTo ErrHandler

    Dim request As Infra_ScopedRequest
    Dim promptMsg As String
    Dim confirmMsg As String
    Dim normalizedChoice As String

    promptMsg = "Convert formulas to values." & vbCrLf & vbCrLf & _
                BuildCompactContextSummary(ctx, False) & vbCrLf & vbCrLf & _
                "Scope:" & vbCrLf & _
                "Sheet - Active sheet" & vbCrLf & _
                "Workbook - All sheets"

    confirmMsg = "You are about to convert formulas on every worksheet in " & SafeWorkbookName(ctx) & "." & vbCrLf & vbCrLf & _
                 "This is not reversible as a single workbook-wide undo action." & vbCrLf & vbCrLf & _
                 "Continue with workbook-wide conversion?"

    If Not PromptForScopeSelection(ctx, "Make Static", promptMsg, "Sheet", Array("Sheet", "Workbook"), confirmMsg, normalizedChoice) Then GoTo CleanExit

    Set ShowStaticConversionDialog = CreateScopedRequest(ctx, normalizedChoice)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowStaticConversionDialog", Err
    Resume CleanExit
End Function

Public Function ShowBreakLinksDialog(ByVal ctx As Infra_ActionContext, ByVal linkInfo As String) As Infra_ScopedRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowBreakLinksDialog")
    On Error GoTo ErrHandler

    Dim normalizedChoice As String
    Dim options As Variant
    Dim defaultChoice As String
    Dim allowSheetScope As Boolean
    Dim request As Infra_ScopedRequest
    Dim promptMsg As String
    Dim confirmMsg As String

    allowSheetScope = ActiveSheetHasBreakableItems(ctx)
    If allowSheetScope Then
        options = BuildChoiceArray("Sheet", "Workbook")
        defaultChoice = "Sheet"
    Else
        options = BuildChoiceArray("Workbook")
        defaultChoice = "Workbook"
    End If

    promptMsg = "External links were found and can be permanently converted to values." & vbCrLf & vbCrLf & _
                BuildContextSummary(ctx, False) & vbCrLf & vbCrLf & _
                "Detected items:" & vbCrLf & linkInfo & vbCrLf & vbCrLf & _
                "Choose a scope:" & vbCrLf & _
                "Sheet     - Converts linked formulas, pivot tables, and external tables only on the active sheet" & vbCrLf & _
                "Workbook  - Also removes workbook-level links, connections, and external names" & vbCrLf & vbCrLf & _
                IIf(allowSheetScope, "Choose Sheet or Workbook.", "Only Workbook scope can remove the detected workbook-level items from this context.")

    confirmMsg = "This will remove workbook-level links and connections and flatten external content." & vbCrLf & vbCrLf & _
                 "Continue with whole-workbook processing?"

    If Not PromptForScopeSelection(ctx, "Break External Links", promptMsg, defaultChoice, options, confirmMsg, normalizedChoice) Then GoTo CleanExit

    If normalizedChoice = "S" Or normalizedChoice = "SHEET" Or normalizedChoice = "ACTIVE SHEET" Or normalizedChoice = "ACTIVESHEET" Then
        If Not allowSheetScope Then
            Infra_Interaction.ShowWarning "The active sheet has no breakable linked formulas, pivots, or tables. Use Workbook scope to remove the remaining workbook-level items.", BuildDialogTitle("Break External Links")
            GoTo CleanExit
        End If
    End If

    Set ShowBreakLinksDialog = CreateScopedRequest(ctx, normalizedChoice)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowBreakLinksDialog", Err
    Resume CleanExit
End Function

Public Function PromptForDuplicateOutputPath(ByVal ctx As Infra_ActionContext, ByVal suggestedBaseName As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForDuplicateOutputPath")
    On Error GoTo ErrHandler

    PromptForDuplicateOutputPath = PromptForOutputPath( _
        ctx, _
        "Create Duplicate", _
        suggestedBaseName, _
        "xlsx", _
        "Excel Workbook (*.xlsx), *.xlsx")

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForDuplicateOutputPath", Err
    PromptForDuplicateOutputPath = vbNullString
    Resume CleanExit
End Function

Public Function PromptForSheetInsertPosition(ByVal ctx As Infra_ActionContext, ByVal sheetName As String) As SheetInsertPosition
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForSheetInsertPosition")
    On Error GoTo ErrHandler

    Dim userChoice As String
    Dim normalizedChoice As String
    Dim defaultChoice As String

    defaultChoice = IIf(LooksLikeFrontLoadedSheet(sheetName), "Before Current", "After Current")

    Do
        If Not ShowOptionPicker( _
            "Create a new worksheet in " & SafeWorkbookName(ctx) & "." & vbCrLf & vbCrLf & _
            "Sheet name: " & sheetName & vbCrLf & vbCrLf & _
            "Where should the new sheet be inserted?" & vbCrLf & _
            "Before Current - Place it before the active sheet" & vbCrLf & _
            "After Current  - Place it after the active sheet", _
            BuildDialogTitle("Create Sheet"), defaultChoice, BuildChoiceArray("Before Current", "After Current"), userChoice) Then GoTo CleanExit

        normalizedChoice = NormalizeChoiceText(userChoice)
        If normalizedChoice = "" Then normalizedChoice = UCase$(defaultChoice)

        Select Case normalizedChoice
            Case "BEFORE CURRENT", "BEFORECURRENT", "BEFORE", "FRONT"
                PromptForSheetInsertPosition = SheetInsertPositionBeforeCurrent
                Exit Do
            Case "AFTER CURRENT", "AFTERCURRENT", "AFTER", "BACK"
                PromptForSheetInsertPosition = SheetInsertPositionAfterCurrent
                Exit Do
            Case Else
                Infra_Interaction.ShowWarning "Please choose Before Current or After Current.", BuildDialogTitle("Create Sheet")
        End Select
    Loop

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "PromptForSheetInsertPosition", Err
    Resume CleanExit
End Function

Public Function PromptForWrapFormulaPattern(ByVal ctx As Infra_ActionContext, ByVal placeholder As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapFormulaPattern")
    On Error GoTo ErrHandler

    Dim userInput As String

    Do
        If Not ShowInputBox( _
            "Wrap selected formulas or values with a new formula pattern." & vbCrLf & vbCrLf & _
            BuildContextSummary(ctx, True) & vbCrLf & vbCrLf & _
            "Use " & placeholder & " where the existing cell content should go." & vbCrLf & _
            "Example: =ROUND(" & placeholder & ", 0)", _
            BuildDialogTitle("Wrap Formula"), placeholder, userInput) Then GoTo CleanExit

        PromptForWrapFormulaPattern = Trim$(CStr(userInput))
        If PromptForWrapFormulaPattern = vbNullString Then GoTo CleanExit

        If InStr(1, PromptForWrapFormulaPattern, placeholder, vbTextCompare) > 0 Then Exit Do

        Infra_Interaction.ShowWarning "Your formula pattern must include the placeholder " & placeholder & ".", _
                                       BuildDialogTitle("Wrap Formula")
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForWrapFormulaPattern", Err
    Resume CleanExit
End Function

Private Function ResolveExportRange(ByVal ctx As Infra_ActionContext) As Range
    If ctx Is Nothing Then Exit Function
    If ctx.WorksheetRef Is Nothing Then Exit Function
    
    If Not ctx.HasRangeSelection Or ctx.SelectionRange Is Nothing Then
        Set ResolveExportRange = ctx.WorksheetRef.UsedRange
    ElseIf ctx.SelectionRange.Cells.CountLarge = 1 Then
        Set ResolveExportRange = ctx.WorksheetRef.UsedRange
    Else
        Set ResolveExportRange = ctx.SelectionRange
    End If
End Function

Private Function NormalizeExportScale(ByVal scaleInput As String) As Long
    If IsNumeric(scaleInput) Then
        NormalizeExportScale = CLng(scaleInput)
    Else
        NormalizeExportScale = Infra_Config.DEFAULT_EXPORT_SCALE
    End If

    If NormalizeExportScale < 1 Then NormalizeExportScale = 1
    If NormalizeExportScale > Infra_Config.MAX_EXPORT_SCALE Then
        NormalizeExportScale = Infra_Config.MAX_EXPORT_SCALE
    End If
End Function

Private Function PromptForExportScale(ByVal defaultScale As Long) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForExportScale")
    On Error GoTo ErrHandler

    Dim scaleInput As String
    Dim normalizedScale As Long

    Do
        If Not ShowInputBox( _
            "Choose the PNG scale factor." & vbCrLf & vbCrLf & _
            "1 = Smaller file" & vbCrLf & _
            CStr(Infra_Config.DEFAULT_EXPORT_SCALE) & " = Balanced default" & vbCrLf & _
            CStr(Infra_Config.MAX_EXPORT_SCALE) & " = Largest supported image" & vbCrLf & vbCrLf & _
            "Enter a number from 1 to " & Infra_Config.MAX_EXPORT_SCALE & ".", _
            BuildDialogTitle("PNG Quality"), CStr(defaultScale), scaleInput) Then GoTo CleanExit

        If Trim$(CStr(scaleInput)) = vbNullString Then scaleInput = CStr(defaultScale)
        If IsNumeric(scaleInput) Then
            normalizedScale = CLng(scaleInput)
            If normalizedScale >= 1 And normalizedScale <= Infra_Config.MAX_EXPORT_SCALE Then
                PromptForExportScale = normalizedScale
                Exit Do
            End If
        End If

        Infra_Interaction.ShowWarning "Please enter a whole number from 1 to " & Infra_Config.MAX_EXPORT_SCALE & ".", _
                                       BuildDialogTitle("PNG Quality")
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForExportScale", Err
    Resume CleanExit
End Function

Private Function BuildDialogTitle(ByVal dialogName As String) As String
    BuildDialogTitle = Infra_Interaction.FormatTitle(dialogName)
End Function

Private Function BuildContextSummary(ByVal ctx As Infra_ActionContext, Optional ByVal includeSelection As Boolean = True) As String
    Dim summary As String

    If ctx Is Nothing Then Exit Function

    summary = "Workbook: " & SafeWorkbookName(ctx) & vbCrLf & _
              "Sheet: " & SafeWorksheetName(ctx)

    If includeSelection Then
        summary = summary & vbCrLf & "Selection: " & SafeSelectionAddress(ctx)
    End If

    BuildContextSummary = summary
End Function

Private Function BuildCompactContextSummary(ByVal ctx As Infra_ActionContext, Optional ByVal includeSelection As Boolean = True) As String
    Dim summary As String

    summary = "Book: " & SafeWorkbookName(ctx) & vbCrLf & _
              "Sheet: " & SafeWorksheetName(ctx)

    If includeSelection Then
        summary = summary & vbCrLf & "Selection: " & SafeSelectionAddress(ctx)
    End If

    BuildCompactContextSummary = summary
End Function

Private Function BuildExportSummary(ByVal sourceRange As Range) As String
    Dim summary As String

    If sourceRange Is Nothing Then Exit Function

    summary = "Range: " & sourceRange.Address(False, False) & vbCrLf & _
              "Sheet: " & sourceRange.Worksheet.Name & vbCrLf & _
              "Size: " & Format(sourceRange.Rows.Count, "#,##0") & " row(s) x " & _
              Format(sourceRange.Columns.Count, "#,##0") & " column(s)"

    BuildExportSummary = summary
End Function

Private Function BuildSuggestedExportBaseName(ByVal sourceRange As Range, ByVal exportAsPng As Boolean) As String
    Dim fileStem As String

    If sourceRange Is Nothing Then
        fileStem = "BeaverExport"
    Else
        fileStem = IIf(exportAsPng, "RangeImage", "RangePDF") & "_" & _
                   Infra_AppState.SanitizeFileNameStem(sourceRange.Worksheet.Name) & "_" & _
                   BuildRangeFileLabel(sourceRange)
    End If

    BuildSuggestedExportBaseName = fileStem & "_" & Format(Now, "yyyymmdd_hhnnss")
End Function

Private Function BuildRangeFileLabel(ByVal sourceRange As Range) As String
    If sourceRange Is Nothing Then
        BuildRangeFileLabel = "Selection"
    Else
        BuildRangeFileLabel = Replace(sourceRange.Address(False, False), ":", "-")
        BuildRangeFileLabel = Replace(BuildRangeFileLabel, "$", "")
    End If
End Function

Private Function PromptForOutputPath(ByVal ctx As Infra_ActionContext, ByVal taskName As String, ByVal suggestedBaseName As String, ByVal extensionWithoutDot As String, ByVal fileFilter As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForOutputPath")
    On Error GoTo ErrHandler

    Dim desktopPath As String
    Dim initialPath As String
    Dim selectedPath As String
    Dim normalizedBaseName As String

    desktopPath = Infra_AppState.GetDesktopPath()
    If desktopPath = vbNullString Then
        Infra_Interaction.ShowCritical "Could not locate a default save location."
        GoTo CleanExit
    End If

    normalizedBaseName = Infra_AppState.SanitizeFileNameStem(suggestedBaseName)
    If normalizedBaseName = vbNullString Then normalizedBaseName = "BeaverOutput"

    initialPath = Infra_AppState.CombinePath(desktopPath, normalizedBaseName & "." & LCase$(extensionWithoutDot))
    If Not Infra_Interaction.PromptSaveAsPath(BuildDialogTitle(taskName), initialPath, fileFilter, selectedPath) Then GoTo CleanExit

    selectedPath = Infra_AppState.EnsureExtension(selectedPath, extensionWithoutDot)
    If Infra_AppState.FileExists(selectedPath) Then
        If Not Infra_Interaction.Confirm( _
            "A file with this name already exists:" & vbCrLf & selectedPath & vbCrLf & vbCrLf & _
            "Do you want to replace it?", _
            BuildDialogTitle(taskName), vbDefaultButton2) Then
            GoTo CleanExit
        End If
    End If

    PromptForOutputPath = selectedPath

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "PromptForOutputPath", Err
    Resume CleanExit
End Function

Private Function NormalizeChoiceText(ByVal rawValue As Variant) As String
    NormalizeChoiceText = UCase$(Trim$(CStr(rawValue)))
End Function

Private Function BuildChoiceArray(ParamArray values() As Variant) As Variant
    BuildChoiceArray = values
End Function

Private Function ShowInputBox(ByVal promptMsg As String, ByVal title As String, ByVal defaultText As String, ByRef outResult As String) As Boolean
    ShowInputBox = Infra_Interaction.PromptText(promptMsg, title, defaultText, outResult)
End Function

Private Function ShowOptionPicker(ByVal promptMsg As String, ByVal title As String, ByVal defaultChoice As String, ByVal options As Variant, ByRef outResult As String) As Boolean
    ShowOptionPicker = Infra_Interaction.PromptOption(promptMsg, title, defaultChoice, options, outResult)
End Function

Public Function PromptForWrapMode(ByVal ctx As Infra_ActionContext) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapMode")
    On Error GoTo ErrHandler

    Dim userInput As String
    Dim normalizedChoice As String

    Do
        If Not ShowOptionPicker( _
            "Choose how you want to wrap the current selection." & vbCrLf & vbCrLf & _
            BuildContextSummary(ctx, True) & vbCrLf & vbCrLf & _
            "Choose one of these options:" & vbCrLf & _
            "Cell  - Reuse a wrapper formula from another cell" & vbCrLf & _
            "Type  - Enter a formula pattern manually using [value]" & vbCrLf & vbCrLf & _
            "Choose Cell or Type.", _
            BuildDialogTitle("Wrap"), "Type", Array("Cell", "Type"), userInput) Then GoTo CleanExit

        normalizedChoice = NormalizeChoiceText(userInput)
        If normalizedChoice = "" Then normalizedChoice = "TYPE"

        Select Case normalizedChoice
            Case "C", "CELL", "WRAPPER CELL", "WRAPPERCELL"
                PromptForWrapMode = WrapModeCell
                Exit Do
            Case "T", "TYPE", "TYPED", "PATTERN", "MANUAL"
                PromptForWrapMode = WrapModeTyped
                Exit Do
            Case Else
                Infra_Interaction.ShowWarning "Please choose Cell or Type.", BuildDialogTitle("Wrap")
        End Select
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForWrapMode", Err
    Resume CleanExit
End Function

Public Function PromptForSheetName(ByVal ctx As Infra_ActionContext) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForSheetName")
    On Error GoTo ErrHandler

    Dim userInput As String

    Do
        If Not ShowInputBox( _
            "Create a new worksheet in " & SafeWorkbookName(ctx) & "." & vbCrLf & vbCrLf & _
            "Enter the name for the new sheet.", _
            BuildDialogTitle("Create Sheet"), vbNullString, userInput) Then GoTo CleanExit

        userInput = Trim$(userInput)
        If userInput = vbNullString Then
            Infra_Interaction.ShowWarning "Sheet name cannot be blank.", BuildDialogTitle("Create Sheet")
        Else
            PromptForSheetName = userInput
            Exit Do
        End If
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForSheetName", Err
    Resume CleanExit
End Function

Public Function PromptForRelatedCell(ByVal sourceCell As Range) As Range
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForRelatedCell")
    On Error GoTo ErrHandler

    Dim selectedRange As Range
    Dim promptText As String

    If sourceCell Is Nothing Then GoTo CleanExit

    promptText = "You selected " & sourceCell.Address(False, False) & "." & vbCrLf & _
                 "Please select the wrapper cell that contains the formula to apply."

    If Not Infra_Interaction.PromptRange(promptText, BuildDialogTitle("Select Wrapper Cell"), selectedRange) Then GoTo CleanExit
    Set PromptForRelatedCell = selectedRange

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForRelatedCell", Err
    Resume CleanExit
End Function


Private Function LooksLikeFrontLoadedSheet(ByVal sheetName As String) As Boolean
    Dim nameLower As String

    nameLower = LCase$(Trim$(sheetName))
    LooksLikeFrontLoadedSheet = (Left$(nameLower, 7) = "summary" Or Left$(nameLower, 5) = "recon")
End Function

Private Function SafeWorkbookName(ByVal ctx As Infra_ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.WorkbookRef Is Nothing Then Exit Function
    SafeWorkbookName = ctx.WorkbookRef.Name
End Function

Private Function SafeWorksheetName(ByVal ctx As Infra_ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.WorksheetRef Is Nothing Then Exit Function
    SafeWorksheetName = ctx.WorksheetRef.Name
End Function

Private Function SafeSelectionAddress(ByVal ctx As Infra_ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.SelectionRange Is Nothing Then
        SafeSelectionAddress = "(none)"
    Else
        SafeSelectionAddress = ctx.SelectionRange.Address(False, False)
    End If
End Function

Private Function HasUsableSelection(ByVal ctx As Infra_ActionContext) As Boolean
    If ctx Is Nothing Then Exit Function
    HasUsableSelection = ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing
End Function

Private Function ActiveSheetHasBreakableItems(ByVal ctx As Infra_ActionContext) As Boolean
    Dim ws As Worksheet
    Dim formulaCells As Range
    Dim area As Range
    Dim lo As ListObject
    Dim pvt As PivotTable
    Dim formulaArr As Variant
    Dim r As Long, c As Long

    On Error GoTo CleanExit

    If ctx Is Nothing Then GoTo CleanExit
    Set ws = ctx.WorksheetRef
    If ws Is Nothing Then GoTo CleanExit

    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo CleanExit

    If Not formulaCells Is Nothing Then
        For Each area In formulaCells.Areas
            If area.Cells.CountLarge = 1 Then
                If InStr(1, area.Formula, "[", vbTextCompare) > 0 Then
                    ActiveSheetHasBreakableItems = True
                    Exit Function
                End If
            Else
                formulaArr = area.Formula
                For r = 1 To UBound(formulaArr, 1)
                    For c = 1 To UBound(formulaArr, 2)
                        If InStr(1, formulaArr(r, c), "[", vbTextCompare) > 0 Then
                            ActiveSheetHasBreakableItems = True
                            Exit Function
                        End If
                    Next c
                Next r
            End If
        Next area
    End If

    For Each pvt In ws.PivotTables
        ActiveSheetHasBreakableItems = True
        Exit Function
    Next pvt

    For Each lo In ws.ListObjects
        If lo.SourceType <> xlSrcRange Then
            ActiveSheetHasBreakableItems = True
            Exit Function
        End If
    Next lo

CleanExit:
    Exit Function
End Function

Private Function PromptForScopeSelection( _
    ByVal ctx As Infra_ActionContext, _
    ByVal dialogName As String, _
    ByVal promptMsg As String, _
    ByVal defaultChoice As String, _
    ByVal options As Variant, _
    ByVal confirmMessage As String, _
    ByRef outChoice As String) As Boolean
    
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForScopeSelection")
    On Error GoTo ErrHandler

    Dim userChoice As String
    Dim normalizedChoice As String

    Do
        If Not ShowOptionPicker(promptMsg, BuildDialogTitle(dialogName), defaultChoice, options, userChoice) Then
            PromptForScopeSelection = False
            GoTo CleanExit
        End If

        normalizedChoice = NormalizeChoiceText(userChoice)
        If normalizedChoice = "" Then normalizedChoice = UCase$(defaultChoice)

        If (normalizedChoice = "W" Or normalizedChoice = "WB" Or normalizedChoice = "WORKBOOK" Or _
            normalizedChoice = "WHOLE WORKBOOK" Or normalizedChoice = "WHOLEWORKBOOK") And confirmMessage <> "" Then
            If Not Infra_Interaction.Confirm(confirmMessage, BuildDialogTitle("Confirm Workbook Scope"), vbDefaultButton2) Then
                GoTo ContinueLoop
            End If
        End If

        outChoice = normalizedChoice
        PromptForScopeSelection = True
        Exit Do

ContinueLoop:
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForScopeSelection", Err
    PromptForScopeSelection = False
    Resume CleanExit
End Function

Private Function CreateScopedRequest(ByVal ctx As Infra_ActionContext, ByVal choiceText As String) As Infra_ScopedRequest
    Dim request As New Infra_ScopedRequest
    Set request.Context = ctx
    Select Case NormalizeChoiceText(choiceText)
        Case "R", "RANGE", "SELECTED", "SELECTION"
            request.Scope = TargetScopeSelection
        Case "S", "SHEET", "ACTIVE SHEET", "ACTIVESHEET"
            request.Scope = TargetScopeActiveSheet
        Case "W", "WB", "WORKBOOK", "WHOLE WORKBOOK", "WHOLEWORKBOOK"
            request.Scope = TargetScopeWorkbook
        Case Else
            Set CreateScopedRequest = Nothing
            Exit Function
    End Select
    Set CreateScopedRequest = request
End Function

Private Function CreateCleanDataRequest(ByVal ctx As Infra_ActionContext, ByVal choiceText As String) As Infra_CleanDataRequest
    Dim request As New Infra_CleanDataRequest
    Set request.Context = ctx
    Select Case NormalizeChoiceText(choiceText)
        Case "R", "RANGE", "SELECTED", "SELECTION"
            request.Scope = TargetScopeSelection
        Case "S", "SHEET", "ACTIVE SHEET", "ACTIVESHEET"
            request.Scope = TargetScopeActiveSheet
        Case "W", "WB", "WORKBOOK", "WHOLE WORKBOOK", "WHOLEWORKBOOK"
            request.Scope = TargetScopeWorkbook
        Case Else
            Set CreateCleanDataRequest = Nothing
            Exit Function
    End Select
    Set CreateCleanDataRequest = request
End Function
