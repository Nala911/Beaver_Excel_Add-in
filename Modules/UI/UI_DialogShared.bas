Attribute VB_Name = "UI_DialogShared"
Option Explicit

' @Module: UI_DialogShared
' @Category: UI
' @Description: Shared helper functions and scope-based dialogs for UI prompts.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Interaction, Infra_Config, ActionContext, ScopedRequest, CleanDataRequest, HighlightDataRequest

' --- PUBLIC DIALOGS ---

' Shows the conversion scope dialog for formula-to-value actions using a UserForm picker.
Public Function ShowStaticConversionDialog(ByVal ctx As ActionContext) As ScopedRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowStaticConversionDialog")
    On Error GoTo ErrHandler

    Dim promptMsg As String
    Dim confirmMsg As String
    Dim normalizedChoice As String

    promptMsg = BuildScopePromptMsg("Convert formulas to values.", HasUsableSelection(ctx), False)
    confirmMsg = BuildScopeConfirmMsg("Make Static", SafeWorkbookName(ctx), "You are about to convert formulas on every worksheet in " & SafeWorkbookName(ctx) & "." & vbCrLf & vbCrLf & "This is not reversible as a single workbook-wide undo action.")

    If Not PromptForScopeSelection(ctx, "Make Static", promptMsg, "Sheet", BuildChoiceArray("Sheet", "Workbook"), confirmMsg, normalizedChoice) Then GoTo CleanExit

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

' --- PUBLIC HELPERS ---

Public Function BuildDialogTitle(ByVal dialogName As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("BuildDialogTitle")
    On Error GoTo ErrHandler

    BuildDialogTitle = Infra_Interaction.FormatTitle(dialogName)

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "BuildDialogTitle", Err
    Resume CleanExit
End Function

Public Function SafeWorkbookName(ByVal ctx As ActionContext) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("SafeWorkbookName")
    On Error GoTo ErrHandler

    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorkbookRef Is Nothing Then GoTo CleanExit
    SafeWorkbookName = ctx.WorkbookRef.Name

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "SafeWorkbookName", Err
    Resume CleanExit
End Function

Public Function SafeWorksheetName(ByVal ctx As ActionContext) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("SafeWorksheetName")
    On Error GoTo ErrHandler

    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Then GoTo CleanExit
    SafeWorksheetName = ctx.WorksheetRef.Name

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "SafeWorksheetName", Err
    Resume CleanExit
End Function

Public Function SafeSelectionAddress(ByVal ctx As ActionContext) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("SafeSelectionAddress")
    On Error GoTo ErrHandler

    If ctx Is Nothing Then
        SafeSelectionAddress = "(none)"
        GoTo CleanExit
    End If
    
    If ctx.SelectionRange Is Nothing Then
        SafeSelectionAddress = "(none)"
    Else
        SafeSelectionAddress = ctx.SelectionRange.Address(False, False)
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "SafeSelectionAddress", Err
    Resume CleanExit
End Function

Public Function HasUsableSelection(ByVal ctx As ActionContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("HasUsableSelection")
    On Error GoTo ErrHandler

    If ctx Is Nothing Then
        HasUsableSelection = False
        GoTo CleanExit
    End If
    
    HasUsableSelection = ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "HasUsableSelection", Err
    Resume CleanExit
End Function

Public Function BuildScopePromptMsg(ByVal description As String, ByVal hasSelection As Boolean, ByVal bypassScopePrompt As Boolean) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("BuildScopePromptMsg")
    On Error GoTo ErrHandler

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

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "BuildScopePromptMsg", Err
    Resume CleanExit
End Function

Public Function BuildScopeConfirmMsg(ByVal taskName As String, ByVal workbookName As String, Optional ByVal customDetail As String = vbNullString) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("BuildScopeConfirmMsg")
    On Error GoTo ErrHandler

    Dim detail As String
    If customDetail <> vbNullString Then
        detail = customDetail
    Else
        detail = "Workbook-wide " & taskName & " updates every sheet in '" & workbookName & "' and cannot be restored as a single workbook-wide undo action."
    End If
    BuildScopeConfirmMsg = detail & vbCrLf & vbCrLf & "Continue with workbook-wide processing?"

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "BuildScopeConfirmMsg", Err
    Resume CleanExit
End Function

Public Function PromptForScopeSelection( _
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
        If Not Infra_Interaction.PromptOption(promptMsg, BuildDialogTitle(dialogName), defaultChoice, options, userChoice, Replace(dialogName, " ", "") & "ScopeOptions") Then
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

Public Function CreateCleanDataRequest(ByVal ctx As ActionContext, ByVal choiceText As String) As CleanDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateCleanDataRequest")
    On Error GoTo ErrHandler

    Dim scopeVal As TargetScope
    If Not ResolveScopeFromText(choiceText, scopeVal) Then
        Set CreateCleanDataRequest = Nothing
        GoTo CleanExit
    End If

    Dim request As New CleanDataRequest
    Set request.Context = ctx
    request.Scope = scopeVal
    Set CreateCleanDataRequest = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "CreateCleanDataRequest", Err
    Resume CleanExit
End Function

Public Function CreateHighlightDataRequest(ByVal ctx As ActionContext, ByVal choiceText As String) As HighlightDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateHighlightDataRequest")
    On Error GoTo ErrHandler

    Dim scopeVal As TargetScope
    If Not ResolveScopeFromText(choiceText, scopeVal) Then
        Set CreateHighlightDataRequest = Nothing
        GoTo CleanExit
    End If

    Dim request As New HighlightDataRequest
    Set request.Context = ctx
    request.Scope = scopeVal
    Set CreateHighlightDataRequest = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "CreateHighlightDataRequest", Err
    Resume CleanExit
End Function

Public Function NormalizeChoiceText(ByVal rawValue As Variant) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("NormalizeChoiceText")
    On Error GoTo ErrHandler

    NormalizeChoiceText = UCase$(Trim$(CStr(rawValue)))

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "NormalizeChoiceText", Err
    Resume CleanExit
End Function

Public Function BuildChoiceArray(ParamArray values() As Variant) As Variant
    Dim tracker As Object: Set tracker = Infra_Error.Track("BuildChoiceArray")
    On Error GoTo ErrHandler

    BuildChoiceArray = values

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "BuildChoiceArray", Err
    Resume CleanExit
End Function

' --- PRIVATE HELPERS ---

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
    If Not ResolveScopeFromText(choiceText, scopeVal) Then
        Set CreateScopedRequest = Nothing
        Exit Function
    End If

    Dim request As New ScopedRequest
    Set request.Context = ctx
    request.Scope = scopeVal
    Set CreateScopedRequest = request
End Function
