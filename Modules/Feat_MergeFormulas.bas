Attribute VB_Name = "Feat_MergeFormulas"
Option Explicit

' @Module: Feat_MergeFormulas
' @Category: Feature
' @Description: Merges directly related formulas by inlining a precedent into its dependent.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_AppStateGuard, Infra_Config, Infra_Error

Public Sub MergeFormulas()
    PushContext "MergeFormulas"
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim userSelection As Range
    Dim otherSelection As Range
    Dim precCell As Range
    Dim depCell As Range
    Dim precFormula As String
    Dim depFormula As String
    Dim finalFormula As String
    
    If Not Infra_AppState.IsRangeSelected() Then
        MsgBox "Please select a single cell first.", vbExclamation, Infra_Config.ADDIN_NAME
        GoTo CleanExit
    End If
    Set userSelection = Selection.Cells(1, 1)
    
    On Error Resume Next
    Set otherSelection = Application.InputBox( _
        Prompt:="You selected " & userSelection.Address(False, False) & "." & vbCrLf & _
                "Please select the related cell (Precedent or Dependent):", _
        Title:="Select Related Cell", _
        Type:=8)
    On Error GoTo ErrHandler
    
    If otherSelection Is Nothing Then GoTo CleanExit
    Set otherSelection = otherSelection.Cells(1, 1)
    
    If otherSelection.Address = userSelection.Address Then
        MsgBox "You selected the same cell twice.", vbExclamation, Infra_Config.ADDIN_NAME
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
        MsgBox "Could not detect a direct dependency between these two cells." & vbCrLf & _
               "Ensure one cell's formula directly references the other.", vbCritical, Infra_Config.ADDIN_NAME
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
    
    On Error Resume Next
    userSelection.Formula2 = finalFormula
    
    If Err.Number <> 0 Then
        MsgBox "Error applying merged formula. The result might be too long or invalid." & vbCrLf & _
               "Attempted Formula: " & finalFormula, vbExclamation, Infra_Config.ADDIN_NAME
        On Error GoTo ErrHandler
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler
    
CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "MergeFormulas", Err
End Sub
