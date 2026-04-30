VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_MergeFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_MergeFormulas
' @Category: Feature
' @Description: Merges directly related formulas by inlining a precedent into its dependent.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy(True, False, True)
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("MergeFormulas.Validate")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)

    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then
        Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("Please select a single cell first.")
    ElseIf ctx.SelectionRange.Cells.CountLarge <> 1 Then
        Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("Please select a single cell first.")
    Else
        Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "MergeFormulas.Validate", Err
    Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("The selected formula cells cannot be merged right now.")
    Resume CleanExit
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("MergeFormulas.Execute")
    On Error GoTo ErrHandler

    Dim userSelection As Range
    Dim otherSelection As Range
    Dim precCell As Range
    Dim depCell As Range
    Dim precFormula As String
    Dim depFormula As String
    Dim finalFormula As String
    Dim targetCell As Range
    Dim ctx As Infra_ActionContext
    
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then
        GoTo CleanExit
    End If
    Set userSelection = ctx.SelectionRange.Cells(1, 1)
    
    Set otherSelection = Infra_UIFactory.PromptForRelatedCell(userSelection)
    If otherSelection Is Nothing Then GoTo CleanExit
    Set otherSelection = otherSelection.Cells(1, 1)
    
    If otherSelection.Address = userSelection.Address Then
        Infra_Interaction.ShowWarning "You selected the same cell twice."
        GoTo CleanExit
    End If

    Dim isUserPrec As Boolean
    Dim foundRelationship As Boolean
    foundRelationship = False
    
    Dim uAddr As String, oAddr As String
    uAddr = userSelection.Address(False, False)
    oAddr = otherSelection.Address(False, False)
    
    If InStr(1, userSelection.Formula2, otherSelection.Address) > 0 Or _
       InStr(1, userSelection.Formula2, oAddr) > 0 Then
        Set depCell = userSelection
        Set precCell = otherSelection
        isUserPrec = False
        foundRelationship = True
    End If
    
    If Not foundRelationship Then
        If InStr(1, otherSelection.Formula2, userSelection.Address) > 0 Or _
           InStr(1, otherSelection.Formula2, uAddr) > 0 Then
            Set depCell = otherSelection
            Set precCell = userSelection
            isUserPrec = True
            foundRelationship = True
        End If
    End If
    
    If Not foundRelationship Then
        Infra_Interaction.ShowCritical "Could not detect a direct dependency between these two cells." & vbCrLf & _
                                       "Ensure one cell's formula directly references the other."
        GoTo CleanExit
    End If

    If precCell.HasFormula Then
        precFormula = Right(precCell.Formula2, Len(precCell.Formula2) - 1)
    Else
        If Application.WorksheetFunction.IsText(precCell) Then
            precFormula = """" & Replace(precCell.Value, """", """""") & """"
        Else
            precFormula = CStr(precCell.Value)
        End If
    End If
    
    precFormula = "(" & precFormula & ")"
    depFormula = depCell.Formula2
    
    Dim addrAbs As String, addrRel As String
    Dim addrAbsSpill As String, addrRelSpill As String
    
    addrAbs = precCell.Address
    addrRel = precCell.Address(False, False)
    addrAbsSpill = addrAbs & "#"
    addrRelSpill = addrRel & "#"
    
    finalFormula = depFormula
    
    If InStr(1, finalFormula, addrAbsSpill) > 0 Then
        finalFormula = Replace(finalFormula, addrAbsSpill, precFormula)
    End If
    If InStr(1, finalFormula, addrRelSpill) > 0 Then
        finalFormula = Replace(finalFormula, addrRelSpill, precFormula)
    End If
    
    If InStr(1, finalFormula, addrAbs) > 0 Then
        finalFormula = Replace(finalFormula, addrAbs, precFormula)
    End If
    If InStr(1, finalFormula, addrRel) > 0 Then
        finalFormula = Replace(finalFormula, addrRel, precFormula)
    End If
    
    Set targetCell = depCell
    If targetCell Is Nothing Then GoTo CleanExit
    Infra_Undo.SaveState targetCell, "Merge Formulas"

    On Error Resume Next
    targetCell.Formula2 = finalFormula
    
    If Err.Number <> 0 Then
        Err.Clear
        Infra_Interaction.ShowWarning "Error applying merged formula. The result might be too long or invalid." & vbCrLf & _
                                      "Attempted Formula: " & finalFormula
        On Error GoTo ErrHandler
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "MergeFormulas.Execute", Err
    Resume CleanExit
End Sub
