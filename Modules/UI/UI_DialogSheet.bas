Attribute VB_Name = "UI_DialogSheet"
Option Explicit

' @Module: UI_DialogSheet
' @Category: UI
' @Description: Prompt dialogs for worksheet naming and insertion placement in workbook.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Interaction, ActionContext, UI_DialogShared, Enums

Public Function PromptForSheetInsertPosition(ByVal ctx As ActionContext, ByVal sheetName As String) As SheetInsertPosition
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForSheetInsertPosition")
    On Error GoTo ErrHandler

    Dim userChoice As String
    Dim normalizedChoice As String
    Dim defaultChoice As String

    defaultChoice = IIf(LooksLikeFrontLoadedSheet(sheetName), "Before Current", "After Current")

    Do
        If Not Infra_Interaction.PromptOption( _
            "Create a new worksheet in " & UI_DialogShared.SafeWorkbookName(ctx) & "." & vbCrLf & vbCrLf & _
            "Sheet name: " & sheetName & vbCrLf & vbCrLf & _
            "Where should the new sheet be inserted?" & vbCrLf & _
            "Before Current - Place it before the active sheet" & vbCrLf & _
            "After Current  - Place it after the active sheet", _
            UI_DialogShared.BuildDialogTitle("Create Sheet"), defaultChoice, UI_DialogShared.BuildChoiceArray("Before Current", "After Current"), userChoice, "CreateSheetInsertPositionOptions") Then GoTo CleanExit

        normalizedChoice = UI_DialogShared.NormalizeChoiceText(userChoice)
        If normalizedChoice = "" Then normalizedChoice = UCase$(defaultChoice)

        Select Case normalizedChoice
            Case "BEFORE CURRENT", "BEFORECURRENT", "BEFORE", "FRONT"
                PromptForSheetInsertPosition = SheetInsertPositionBeforeCurrent
                Exit Do
            Case "AFTER CURRENT", "AFTERCURRENT", "AFTER", "BACK"
                PromptForSheetInsertPosition = SheetInsertPositionAfterCurrent
                Exit Do
            Case Else
                Infra_Interaction.ShowWarning "Please choose Before Current or After Current.", UI_DialogShared.BuildDialogTitle("Create Sheet")
        End Select
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForSheetInsertPosition", Err
    Resume CleanExit
End Function

Public Function PromptForSheetName(ByVal ctx As ActionContext) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForSheetName")
    On Error GoTo ErrHandler

    Dim userInput As String

    Do
        If Not Infra_Interaction.PromptText( _
            "Create a new worksheet in " & UI_DialogShared.SafeWorkbookName(ctx) & "." & vbCrLf & vbCrLf & _
            "Enter the name for the new sheet.", _
            UI_DialogShared.BuildDialogTitle("Create Sheet"), vbNullString, userInput) Then GoTo CleanExit

        userInput = Trim$(userInput)
        If userInput = vbNullString Then
            Infra_Interaction.ShowWarning "Sheet name cannot be blank.", UI_DialogShared.BuildDialogTitle("Create Sheet")
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

' --- PRIVATE HELPERS ---

Private Function LooksLikeFrontLoadedSheet(ByVal sheetName As String) As Boolean
    Dim nameLower As String

    nameLower = LCase$(Trim$(sheetName))
    LooksLikeFrontLoadedSheet = (Left$(nameLower, 7) = "summary" Or Left$(nameLower, 5) = "recon")
End Function
