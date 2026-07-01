Attribute VB_Name = "UI_DialogCleanData"
Option Explicit

' @Module: UI_DialogCleanData
' @Category: UI
' @Description: Option selection dialog for text and formatting cleanup (Clean Data).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Interaction, ActionContext, CleanDataRequest, UI_DialogShared

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

    hasSelection = UI_DialogShared.HasUsableSelection(ctx)
    bypassScopePrompt = False

    If hasSelection Then
        If ctx.SelectionRange.Cells.CountLarge > 1 Then
            normalizedChoice = "RANGE"
            bypassScopePrompt = True
        Else
            options = UI_DialogShared.BuildChoiceArray("Sheet", "Workbook")
            defaultChoice = "Sheet"
        End If
    Else
        options = UI_DialogShared.BuildChoiceArray("Sheet", "Workbook")
        defaultChoice = "Sheet"
    End If

    promptMsg = UI_DialogShared.BuildScopePromptMsg("Clean text with TRIM and CLEAN.", hasSelection, bypassScopePrompt)
    confirmMsg = UI_DialogShared.BuildScopeConfirmMsg("Clean Data", UI_DialogShared.SafeWorkbookName(ctx))

    If Not bypassScopePrompt Then
        If Not UI_DialogShared.PromptForScopeSelection(ctx, "Clean Data", promptMsg, defaultChoice, options, confirmMsg, normalizedChoice) Then GoTo CleanExit
    End If

    If normalizedChoice = "R" Or normalizedChoice = "RANGE" Or normalizedChoice = "SELECTED" Or normalizedChoice = "SELECTION" Then
        If Not hasSelection Then
            Infra_Interaction.ShowWarning "Select a range first if you want to clean only the current selection.", UI_DialogShared.BuildDialogTitle("Clean Data")
            GoTo CleanExit
        End If
    End If

    Set request = UI_DialogShared.CreateCleanDataRequest(ctx, normalizedChoice)
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
