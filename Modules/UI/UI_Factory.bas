Attribute VB_Name = "UI_Factory"
Option Explicit

' @Module: UI_Factory
' @Category: UI
' @Description: Centralized factory for creating and displaying standardized user prompts.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Config, ExportRequest, ScopedRequest, ActionContext, HighlightDataRequest, ModifyDataRequest

' Shows the Clean Data options via UserForm picker and returns a populated Request object.
Public Function ShowCleanDataDialog(ByVal ctx As ActionContext) As CleanDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowCleanDataDialog")
    On Error GoTo ErrHandler
    
    Dim promptMsg As String
    Dim normalizedChoice As String
    Dim request As CleanDataRequest
    Dim options As Variant
    Dim defaultChoice As String
    Dim hasSelection As Boolean
    Dim confirmMsg As String
    Dim bypassScopePrompt As Boolean

    hasSelection = HasUsableSelection(ctx)
    bypassScopePrompt = False

    If hasSelection Then
        If ctx.SelectionRange.Cells.CountLarge > 1 Then
            normalizedChoice = "RANGE"
            bypassScopePrompt = True
        Else
            options = BuildChoiceArray("Sheet", "Workbook")
            defaultChoice = "Sheet"
        End If
    Else
        options = BuildChoiceArray("Sheet", "Workbook")
        defaultChoice = "Sheet"
    End If

    promptMsg = BuildScopePromptMsg("Clean text with TRIM and CLEAN.", hasSelection, bypassScopePrompt)
    confirmMsg = BuildScopeConfirmMsg("Clean Data", SafeWorkbookName(ctx))

    If Not bypassScopePrompt Then
        If Not PromptForScopeSelection(ctx, "Clean Data", promptMsg, defaultChoice, options, confirmMsg, normalizedChoice) Then GoTo CleanExit
    End If

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
        ChrW$(9670) & "  TEXT CLEANING", _
        "  Trim extra spaces & non-breaking spaces", _
        "  Remove non-printable characters", _
        "  Standardize invisible chars (zero-width, BOM, thin spaces)", _
        "  Standardize dashes (convert – — − to -)", _
        "  Remove accents (convert diacritics to standard letters)", _
        ChrW$(9670) & "  LINE BREAKS", _
        "  Line breaks: Replace with space", _
        "  Line breaks: Remove entirely", _
        "  Line breaks: Standardize to single LF", _
        ChrW$(9670) & "  HYGIENE & FORMATS", _
        "  Convert numeric text to numbers", _
        "  Delete broken named ranges (#REF!)", _
        "  Remove special symbols (tabs, bullets, ™, ®, ©)", _
        "  Remove comments & notes", _
        "  Remove data validation rules", _
        "  Remove conditional formatting", _
        "  Clear cell formatting (keep values)", _
        "  Remove shapes & images (keeps charts)", _
        "  Remove named ranges scoped to this sheet" _
    )
    cleanDefaultsChecked = Array( _
        False, _
        True, True, True, False, False, _
        False, _
        False, False, False, _
        False, _
        False, True, False, False, _
        False, False, False, False, False _
    )

    If Not Infra_Interaction.PromptMultiOption( _
        "Select the data cleaning options to apply:", _
        "Clean Data Options", _
        cleanOptionsList, _
        cleanDefaultsChecked, _
        selectedIndices, _
        "CleanDataOptions") Then
        GoTo CleanExit
    End If

    request.CleanTrimSpaces = False
    request.CleanNonPrintables = False
    request.CleanInvisibleChars = False
    request.CleanReplaceLineBreaksWithSpace = False
    request.CleanRemoveLineBreaks = False
    request.CleanStandardizeLineBreaks = False
    request.CleanConvertNumbers = False
    request.CleanBrokenNames = False
    request.CleanSpecialSymbols = False
    request.CleanStandardizeDashes = False
    request.CleanRemoveAccents = False
    request.CleanComments = False
    request.CleanValidation = False
    request.CleanConditionalFormatting = False
    request.CleanFormats = False
    request.CleanShapes = False
    request.CleanSheetNames = False

    For Each idx In selectedIndices
        Select Case idx
            Case 1: request.CleanTrimSpaces = True
            Case 2: request.CleanNonPrintables = True
            Case 3: request.CleanInvisibleChars = True
            Case 4: request.CleanStandardizeDashes = True
            Case 5: request.CleanRemoveAccents = True
            Case 7: request.CleanReplaceLineBreaksWithSpace = True
            Case 8: request.CleanRemoveLineBreaks = True
            Case 9: request.CleanStandardizeLineBreaks = True
            Case 11: request.CleanConvertNumbers = True
            Case 12: request.CleanBrokenNames = True
            Case 13: request.CleanSpecialSymbols = True
            Case 14: request.CleanComments = True
            Case 15: request.CleanValidation = True
            Case 16: request.CleanConditionalFormatting = True
            Case 17: request.CleanFormats = True
            Case 18: request.CleanShapes = True
            Case 19: request.CleanSheetNames = True
        End Select
    Next idx
    Set ShowCleanDataDialog = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowCleanDataDialog", Err
    Resume CleanExit
End Function

' Shows the Modify Data options via UserForm picker and returns a populated Request object.
Public Function ShowModifyDataDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As ModifyDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowModifyDataDialog")
    On Error GoTo ErrHandler

    Dim request As ModifyDataRequest
    Dim hasSelection As Boolean

    hasSelection = HasUsableSelection(ctx)
    If Not hasSelection Then
        Infra_Interaction.ShowWarning "Please select a range of cells to modify first.", "Modify Data"
        GoTo CleanExit
    End If

    Dim selectedTool As String
    If LCase$(commandName) = "datefixer" Then
        selectedTool = "Date Fixer"
    ElseIf LCase$(commandName) = "casefixer" Then
        selectedTool = "Case Fixer"
    Else
        Dim toolOptions As Variant
        toolOptions = Array("Date Fixer", "Case Fixer")

        If Not Infra_Interaction.PromptOption( _
            "Select the modification tool to apply to the selection:", _
            "Modify Data Options", _
            "Date Fixer", _
            toolOptions, _
            selectedTool, _
            "ModifyDataToolOptions") Then
            GoTo CleanExit
        End If
    End If

    Dim selectedOp As String
    Dim datePattern As String
    datePattern = vbNullString

    If selectedTool = "Date Fixer" Then
        selectedOp = "Date Standardization"
        If Not Infra_Interaction.PromptText( _
            "Enter the date pattern format used in the source data (e.g. DD/MM/YYYY, MM/DD/YYYY, YYYYMMDD):", _
            "Date Standardization", _
            "DD/MM/YYYY", _
            datePattern) Then
            GoTo CleanExit
        End If
        
        datePattern = Trim$(datePattern)
        If datePattern = vbNullString Then
            Infra_Interaction.ShowWarning "A date pattern is required for Date Standardization.", "Modify Data"
            GoTo CleanExit
        End If
        
        Dim patternLower As String
        patternLower = LCase$(datePattern)
        If InStr(1, patternLower, "d") = 0 Or InStr(1, patternLower, "m") = 0 Or InStr(1, patternLower, "y") = 0 Then
            Infra_Interaction.ShowWarning "Invalid date pattern. The pattern must specify Day (d), Month (m), and Year (y).", "Modify Data"
            GoTo CleanExit
        End If

    ElseIf selectedTool = "Case Fixer" Then
        Dim caseOptions As Variant
        caseOptions = Array("UPPERCASE", "lowercase", "Proper Case")
        
        Dim selectedCase As String
        If Not Infra_Interaction.PromptOption( _
            "Select the text casing standard to apply:", _
            "Case Fixer Options", _
            "UPPERCASE", _
            caseOptions, _
            selectedCase, _
            "CaseFixerCasingOptions") Then
            GoTo CleanExit
        End If
        
        Select Case selectedCase
            Case "UPPERCASE"
                selectedOp = "Case: UPPERCASE"
            Case "lowercase"
                selectedOp = "Case: lowercase"
            Case "Proper Case"
                selectedOp = "Case: Proper Case"
            Case Else
                GoTo CleanExit
        End Select
    End If

    Set request = New ModifyDataRequest
    Set request.Context = ctx
    request.Operation = selectedOp
    request.DatePattern = datePattern

    Set ShowModifyDataDialog = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowModifyDataDialog", Err
    Resume CleanExit
End Function

' Shows the Highlight Data options via UserForm picker and returns a populated Request object.
Public Function ShowHighlightDataDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As HighlightDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowHighlightDataDialog")
    On Error GoTo ErrHandler
    
    Dim promptMsg As String
    Dim normalizedChoice As String
    Dim request As HighlightDataRequest
    Dim options As Variant
    Dim defaultChoice As String
    Dim hasSelection As Boolean
    Dim confirmMsg As String
    Dim bypassScopePrompt As Boolean

    hasSelection = HasUsableSelection(ctx)
    bypassScopePrompt = False

    If hasSelection Then
        If ctx.SelectionRange.Cells.CountLarge > 1 Then
            normalizedChoice = "RANGE"
            bypassScopePrompt = True
        Else
            options = BuildChoiceArray("Sheet", "Workbook")
            defaultChoice = "Sheet"
        End If
    Else
        options = BuildChoiceArray("Sheet", "Workbook")
        defaultChoice = "Sheet"
    End If

    promptMsg = BuildScopePromptMsg("Highlight key data patterns (Inconsistent Formulas, Duplicates, Errors, Hardcoded Values).", hasSelection, bypassScopePrompt)
    confirmMsg = BuildScopeConfirmMsg("Highlight Data", SafeWorkbookName(ctx))

    If Not bypassScopePrompt Then
        If Not PromptForScopeSelection(ctx, "Highlight Data", promptMsg, defaultChoice, options, confirmMsg, normalizedChoice) Then GoTo CleanExit
    End If

    If normalizedChoice = "R" Or normalizedChoice = "RANGE" Or normalizedChoice = "SELECTED" Or normalizedChoice = "SELECTION" Then
        If Not hasSelection Then
            Infra_Interaction.ShowWarning "Select a range first if you want to highlight only the current selection.", BuildDialogTitle("Highlight Data")
            GoTo CleanExit
        End If
    End If

    Set request = CreateHighlightDataRequest(ctx, normalizedChoice)
    If request Is Nothing Then GoTo CleanExit

    Dim chosenOption As String
    Select Case LCase$(commandName)
        Case "highlightinconsistentformulas"
            chosenOption = "Highlight Inconsistent Formulas (yellow)"
        Case "highlightduplicates"
            chosenOption = "Highlight Duplicates (soft red)"
        Case "highlighterrors"
            chosenOption = "Highlight Errors (orange)"
        Case "highlighthardcodedvalues"
            chosenOption = "Highlight Hardcoded Values in Formulas (lavender)"
        Case "highlightdatavalidations"
            chosenOption = "Highlight Data Validations (soft green)"
        Case "highlightconditionalformatting"
            chosenOption = "Highlight Conditional Formatting (soft blue)"
        Case Else
            Dim highlightOptionsList As Variant
            highlightOptionsList = Array( _
                "Highlight Inconsistent Formulas (yellow)", _
                "Highlight Duplicates (soft red)", _
                "Highlight Errors (orange)", _
                "Highlight Hardcoded Values in Formulas (lavender)", _
                "Highlight Data Validations (soft green)", _
                "Highlight Conditional Formatting (soft blue)" _
            )

            If Not Infra_Interaction.PromptOption( _
                "Select the data highlighting option to apply:", _
                "Highlight Data Option", _
                "Highlight Inconsistent Formulas (yellow)", _
                highlightOptionsList, _
                chosenOption, _
                "HighlightDataOptions") Then
                GoTo CleanExit
            End If
    End Select

    request.HighlightInconsistentFormulas = False
    request.HighlightDuplicates = False
    request.HighlightErrors = False
    request.HighlightHardcodedValues = False
    request.HighlightDataValidations = False
    request.HighlightConditionalFormatting = False

    If chosenOption = "Highlight Inconsistent Formulas (yellow)" Then
        request.HighlightInconsistentFormulas = True
    ElseIf chosenOption = "Highlight Duplicates (soft red)" Then
        request.HighlightDuplicates = True
    ElseIf chosenOption = "Highlight Errors (orange)" Then
        request.HighlightErrors = True
    ElseIf chosenOption = "Highlight Hardcoded Values in Formulas (lavender)" Then
        request.HighlightHardcodedValues = True
    ElseIf chosenOption = "Highlight Data Validations (soft green)" Then
        request.HighlightDataValidations = True
    ElseIf chosenOption = "Highlight Conditional Formatting (soft blue)" Then
        request.HighlightConditionalFormatting = True
    End If

    Set ShowHighlightDataDialog = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowHighlightDataDialog", Err
    Resume CleanExit
End Function

' Shows the Export options via UserForm picker and returns a populated Request object.
Public Function ShowExportDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As ExportRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowExportDialog")
    On Error GoTo ErrHandler
    
    Dim request As ExportRequest
    Dim exportChoice As String
    Dim normalizedChoice As String
    
    Set request = New ExportRequest
    Set request.Context = ctx
    Set request.SourceRange = ResolveExportRange(ctx)
    request.ScaleFactor = Infra_Config.DEFAULT_EXPORT_SCALE
    
    If request.SourceRange Is Nothing Then
        Infra_Interaction.ShowWarning "No data found on the active sheet to export."
        GoTo CleanExit
    End If

    Dim skipFormatPrompt As Boolean
    skipFormatPrompt = False
    
    If LCase$(commandName) = "exportpng" Then
        request.ExportAsPng = True
        skipFormatPrompt = True
    ElseIf LCase$(commandName) = "exportpdf" Then
        request.ExportAsPng = False
        skipFormatPrompt = True
    End If

    If Not skipFormatPrompt Then
        Do
            If Not ShowOptionPicker( _
                "Export the selected content and choose where to save it." & vbCrLf & vbCrLf & _
                BuildExportSummary(request.SourceRange) & vbCrLf & vbCrLf & _
                "Choose a format:" & vbCrLf & _
                "PNG - High-resolution image" & vbCrLf & _
                "PDF - Print-ready document" & vbCrLf & vbCrLf & _
                "Choose PNG or PDF.", _
                BuildDialogTitle("Export"), "PNG", Array("PNG", "PDF"), exportChoice, "ExportOptions") Then GoTo CleanExit

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
    End If

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
Public Function ShowStaticConversionDialog(ByVal ctx As ActionContext) As ScopedRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowStaticConversionDialog")
    On Error GoTo ErrHandler

    Dim request As ScopedRequest
    Dim promptMsg As String
    Dim confirmMsg As String
    Dim normalizedChoice As String

    promptMsg = BuildScopePromptMsg("Convert formulas to values.", HasUsableSelection(ctx), False)
    confirmMsg = BuildScopeConfirmMsg("Make Static", SafeWorkbookName(ctx), "You are about to convert formulas on every worksheet in " & SafeWorkbookName(ctx) & "." & vbCrLf & vbCrLf & "This is not reversible as a single workbook-wide undo action.")

    If Not PromptForScopeSelection(ctx, "Make Static", promptMsg, "Sheet", Array("Sheet", "Workbook"), confirmMsg, normalizedChoice) Then GoTo CleanExit

    Set ShowStaticConversionDialog = CreateScopedRequest(ctx, normalizedChoice)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowStaticConversionDialog", Err
    Resume CleanExit
End Function

Public Function ShowBreakLinksDialog(ByVal ctx As ActionContext, ByVal linkInfo As String) As ScopedRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowBreakLinksDialog")
    On Error GoTo ErrHandler

    Dim normalizedChoice As String
    Dim options As Variant
    Dim defaultChoice As String
    Dim allowSheetScope As Boolean
    Dim request As ScopedRequest
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
                "Detected items:" & vbCrLf & linkInfo
    promptMsg = BuildScopePromptMsg(promptMsg, HasUsableSelection(ctx), False)

    confirmMsg = BuildScopeConfirmMsg("Break External Links", SafeWorkbookName(ctx), _
        "This will remove workbook-level links and connections and flatten external content.")

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

Public Function PromptForSheetInsertPosition(ByVal ctx As ActionContext, ByVal sheetName As String) As SheetInsertPosition
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
            BuildDialogTitle("Create Sheet"), defaultChoice, BuildChoiceArray("Before Current", "After Current"), userChoice, "CreateSheetInsertPositionOptions") Then GoTo CleanExit

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

Public Function PromptForWrapFormulaPattern(ByVal ctx As ActionContext, ByVal placeholder As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapFormulaPattern")
    On Error GoTo ErrHandler

    Dim userInput As String

    Do
        If Not ShowInputBox( _
            "Wrap selected formulas or values with a new formula pattern." & vbCrLf & vbCrLf & _
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

Private Function ResolveExportRange(ByVal ctx As ActionContext) As Range
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

Private Function PromptForOutputPath(ByVal ctx As ActionContext, ByVal taskName As String, ByVal suggestedBaseName As String, ByVal extensionWithoutDot As String, ByVal fileFilter As String) As String
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

Private Function ShowOptionPicker(ByVal promptMsg As String, ByVal title As String, ByVal defaultChoice As String, ByVal options As Variant, ByRef outResult As String, Optional ByVal prefNamespace As String = vbNullString) As Boolean
    ShowOptionPicker = Infra_Interaction.PromptOption(promptMsg, title, defaultChoice, options, outResult, prefNamespace)
End Function

Public Function PromptForWrapMode(ByVal ctx As ActionContext) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapMode")
    On Error GoTo ErrHandler

    Dim userInput As String
    Dim normalizedChoice As String

    Do
        If Not ShowOptionPicker( _
            "Choose how you want to wrap the current selection." & vbCrLf & vbCrLf & _
            "Choose one of these options:" & vbCrLf & _
            "Cell  - Reuse a wrapper formula from another cell" & vbCrLf & _
            "Type  - Enter a formula pattern manually using [value]" & vbCrLf & vbCrLf & _
            "Choose Cell or Type.", _
            BuildDialogTitle("Wrap"), "Type", Array("Cell", "Type"), userInput, "WrapOptions") Then GoTo CleanExit

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

Public Function PromptForSheetName(ByVal ctx As ActionContext) As String
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

Private Function SafeWorkbookName(ByVal ctx As ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.WorkbookRef Is Nothing Then Exit Function
    SafeWorkbookName = ctx.WorkbookRef.Name
End Function

Private Function SafeWorksheetName(ByVal ctx As ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.WorksheetRef Is Nothing Then Exit Function
    SafeWorksheetName = ctx.WorksheetRef.Name
End Function

Private Function SafeSelectionAddress(ByVal ctx As ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.SelectionRange Is Nothing Then
        SafeSelectionAddress = "(none)"
    Else
        SafeSelectionAddress = ctx.SelectionRange.Address(False, False)
    End If
End Function

Private Function HasUsableSelection(ByVal ctx As ActionContext) As Boolean
    If ctx Is Nothing Then Exit Function
    HasUsableSelection = ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing
End Function

Private Function ActiveSheetHasBreakableItems(ByVal ctx As ActionContext) As Boolean
    Dim ws As Worksheet
    Dim formulaCount As Long
    Dim pivotCount As Long
    Dim tableCount As Long

    On Error GoTo CleanExit

    If ctx Is Nothing Then GoTo CleanExit
    Set ws = ctx.WorksheetRef
    If ws Is Nothing Then GoTo CleanExit

    Infra_CommandSupport.GetSheetBreakableCounts ws, formulaCount, pivotCount, tableCount
    ActiveSheetHasBreakableItems = (formulaCount > 0 Or pivotCount > 0 Or tableCount > 0)

CleanExit:
    Exit Function
End Function

Private Function BuildScopePromptMsg(ByVal description As String, ByVal hasSelection As Boolean, ByVal bypassScopePrompt As Boolean) As String
    Dim msg As String
    msg = description & vbCrLf & vbCrLf & _
          "Scope:" & vbCrLf & _
          "Sheet - Active sheet" & vbCrLf & _
          "Workbook - All sheets"
    If hasSelection And Not bypassScopePrompt Then
        msg = msg & vbCrLf & vbCrLf & "Note: The active selection is a single cell. Choose Sheet or Workbook scope to proceed."
    ElseIf Not hasSelection Then
        msg = msg & vbCrLf & vbCrLf & "No selection is required for Sheet or Workbook scope."
    End If
    BuildScopePromptMsg = msg
End Function

Private Function BuildScopeConfirmMsg(ByVal taskName As String, ByVal workbookName As String, Optional ByVal customDetail As String = vbNullString) As String
    Dim detail As String
    If customDetail <> vbNullString Then
        detail = customDetail
    Else
        detail = "Workbook-wide " & taskName & " updates every sheet in '" & workbookName & "' and cannot be restored as a single workbook-wide undo action."
    End If
    BuildScopeConfirmMsg = detail & vbCrLf & vbCrLf & "Continue with workbook-wide processing?"
End Function

Private Function PromptForScopeSelection( _
    ByVal ctx As ActionContext, _
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
        If Not ShowOptionPicker(promptMsg, BuildDialogTitle(dialogName), defaultChoice, options, userChoice, Replace(dialogName, " ", "") & "ScopeOptions") Then
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

Private Function ResolveScopeFromText(ByVal choiceText As String, ByRef outScope As TargetScope) As Boolean
    ResolveScopeFromText = True
    Select Case NormalizeChoiceText(choiceText)
        Case "R", "RANGE", "SELECTED", "SELECTION"
            outScope = TargetScopeSelection
        Case "S", "SHEET", "ACTIVE SHEET", "ACTIVESHEET"
            outScope = TargetScopeActiveSheet
        Case "W", "WB", "WORKBOOK", "WHOLE WORKBOOK", "WHOLEWORKBOOK"
            outScope = TargetScopeWorkbook
        Case Else
            ResolveScopeFromText = False
    End Select
End Function

Private Function CreateScopedRequest(ByVal ctx As ActionContext, ByVal choiceText As String) As ScopedRequest
    Dim scopeVal As TargetScope
    if Not ResolveScopeFromText(choiceText, scopeVal) Then
        Set CreateScopedRequest = Nothing
        Exit Function
    End If

    Dim request As New ScopedRequest
    Set request.Context = ctx
    request.Scope = scopeVal
    Set CreateScopedRequest = request
End Function

Private Function CreateCleanDataRequest(ByVal ctx As ActionContext, ByVal choiceText As String) As CleanDataRequest
    Dim scopeVal As TargetScope
    if Not ResolveScopeFromText(choiceText, scopeVal) Then
        Set CreateCleanDataRequest = Nothing
        Exit Function
    End If

    Dim request As New CleanDataRequest
    Set request.Context = ctx
    request.Scope = scopeVal
    Set CreateCleanDataRequest = request
End Function

Private Function CreateHighlightDataRequest(ByVal ctx As ActionContext, ByVal choiceText As String) As HighlightDataRequest
    Dim scopeVal As TargetScope
    if Not ResolveScopeFromText(choiceText, scopeVal) Then
        Set CreateHighlightDataRequest = Nothing
        Exit Function
    End If

    Dim request As New HighlightDataRequest
    Set request.Context = ctx
    request.Scope = scopeVal
    Set CreateHighlightDataRequest = request
End Function
