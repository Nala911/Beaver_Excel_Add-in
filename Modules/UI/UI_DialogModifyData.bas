Attribute VB_Name = "UI_DialogModifyData"
Option Explicit

' @Module: UI_DialogModifyData
' @Category: UI
' @Description: Option selection dialog for modifying cell values (Date Fixer and Case Fixer).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Interaction, ActionContext, ModifyDataRequest, UI_DialogShared

' Shows the Modify Data options via UserForm picker and returns a populated Request object.
Public Function ShowModifyDataDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As ModifyDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowModifyDataDialog")
    On Error GoTo ErrHandler

    Dim request As ModifyDataRequest
    Dim hasSelection As Boolean

    hasSelection = UI_DialogShared.HasUsableSelection(ctx)
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
