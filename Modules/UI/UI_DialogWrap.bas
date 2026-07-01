Attribute VB_Name = "UI_DialogWrap"
Option Explicit

' @Module: UI_DialogWrap
' @Category: UI
' @Description: Prompt dialogs and cell selection inputs for the formula wrapping feature.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Interaction, ActionContext, UI_DialogShared, Enums

Public Function PromptForWrapMode(ByVal ctx As ActionContext) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapMode")
    On Error GoTo ErrHandler

    Dim userInput As String
    Dim normalizedChoice As String

    Do
        If Not Infra_Interaction.PromptOption( _
            "Choose how you want to wrap the current selection." & vbCrLf & vbCrLf & _
            "Choose one of these options:" & vbCrLf & _
            "Cell  - Reuse a wrapper formula from another cell" & vbCrLf & _
            "Type  - Enter a formula pattern manually using [value]" & vbCrLf & vbCrLf & _
            "Choose Cell or Type.", _
            UI_DialogShared.BuildDialogTitle("Wrap"), "Type", UI_DialogShared.BuildChoiceArray("Cell", "Type"), userInput, "WrapOptions") Then GoTo CleanExit

        normalizedChoice = UI_DialogShared.NormalizeChoiceText(userInput)
        If normalizedChoice = "" Then normalizedChoice = "TYPE"

        Select Case normalizedChoice
            Case "C", "CELL", "WRAPPER CELL", "WRAPPERCELL"
                PromptForWrapMode = WrapModeCell
                Exit Do
            Case "T", "TYPE", "TYPED", "PATTERN", "MANUAL"
                PromptForWrapMode = WrapModeTyped
                Exit Do
            Case Else
                Infra_Interaction.ShowWarning "Please choose Cell or Type.", UI_DialogShared.BuildDialogTitle("Wrap")
        End Select
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForWrapMode", Err
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

    If Not Infra_Interaction.PromptRange(promptText, UI_DialogShared.BuildDialogTitle("Select Wrapper Cell"), selectedRange) Then GoTo CleanExit
    Set PromptForRelatedCell = selectedRange

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForRelatedCell", Err
    Resume CleanExit
End Function

Public Function PromptForWrapFormulaPattern(ByVal ctx As ActionContext, ByVal placeholder As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapFormulaPattern")
    On Error GoTo ErrHandler

    Dim userInput As String

    Do
        If Not Infra_Interaction.PromptText( _
            "Wrap selected formulas or values with a new formula pattern." & vbCrLf & vbCrLf & _
            "Use " & placeholder & " where the existing cell content should go." & vbCrLf & _
            "Example: =ROUND(" & placeholder & ", 0)", _
            UI_DialogShared.BuildDialogTitle("Wrap Formula"), placeholder, userInput) Then GoTo CleanExit

        PromptForWrapFormulaPattern = Trim$(CStr(userInput))
        If PromptForWrapFormulaPattern = vbNullString Then GoTo CleanExit

        If InStr(1, PromptForWrapFormulaPattern, placeholder, vbTextCompare) > 0 Then Exit Do

        Infra_Interaction.ShowWarning "Your formula pattern must include the placeholder " & placeholder & ".", _
                                       UI_DialogShared.BuildDialogTitle("Wrap Formula")
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForWrapFormulaPattern", Err
    Resume CleanExit
End Function
