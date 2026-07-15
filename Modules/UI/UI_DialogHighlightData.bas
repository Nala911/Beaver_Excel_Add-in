Attribute VB_Name = "UI_DialogHighlightData"
Option Explicit

' @Module: UI_DialogHighlightData
' @Category: UI
' @Description: Option selection dialog for data highlighting (Highlight Data).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Interaction, ActionContext, HighlightDataRequest, UI_DialogShared

' Shows the Highlight Data options via UserForm picker and returns a populated Request object.
Public Function ShowHighlightDataDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As HighlightDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowHighlightDataDialog")
    On Error GoTo ErrHandler
    
    Dim normalizedChoice As String
    Dim request As HighlightDataRequest
    Dim hasSelection As Boolean

    hasSelection = UI_DialogShared.HasUsableSelection(ctx)

    If hasSelection Then
        If ctx.SelectionRange.Cells.CountLarge > 1 Then
            normalizedChoice = "RANGE"
        Else
            normalizedChoice = "WORKBOOK"
        End If
    Else
        normalizedChoice = "WORKBOOK"
    End If

    Set request = UI_DialogShared.CreateHighlightDataRequest(ctx, normalizedChoice)
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
        Case "highlightnamedranges"
            chosenOption = "Highlight Named Ranges (light blue)"
        Case Else
            Dim highlightOptionsList As Variant
            highlightOptionsList = Array( _
                "Highlight Inconsistent Formulas (yellow)", _
                "Highlight Duplicates (soft red)", _
                "Highlight Errors (orange)", _
                "Highlight Hardcoded Values in Formulas (lavender)", _
                "Highlight Data Validations (soft green)", _
                "Highlight Conditional Formatting (soft blue)", _
                "Highlight Named Ranges (light blue)" _
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
    request.HighlightNamedRanges = False

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
    ElseIf chosenOption = "Highlight Named Ranges (light blue)" Then
        request.HighlightNamedRanges = True
    End If

    Set ShowHighlightDataDialog = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowHighlightDataDialog", Err
    Resume CleanExit
End Function
