VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_BatchProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_BatchProcessor
' @Category: Infrastructure
' @Description: Centralized engine for executing scoped workbook/sheet/range batch operations with unified progress, undo, and state management.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: IBatchProcessor, ActionContext, Infra_Progress, Infra_Undo, Infra_Error, Infra_CommandSupport, Infra_Interaction

Public Function Execute( _
    ByVal ctx As ActionContext, _
    ByVal scope As TargetScope, _
    ByVal processor As IBatchProcessor, _
    ByVal operationName As String, _
    Optional ByVal showCompletionMsg As Boolean = True, _
    Optional ByVal captureMode As UndoCaptureMode = UndoCaptureFull) As Long

    Dim tracker As Object: Set tracker = Infra_Error.Track("Infra_BatchProcessor.Execute")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim totalProcessed As Long
    Dim sheetsToProcess As Collection
    Dim wsItem As Variant
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim index As Long
    Dim scopeLabel As String

    If ctx Is Nothing Then GoTo CleanExit
    If processor Is Nothing Then GoTo CleanExit

    ' --- 1. Check and save Undo state ---
    Select Case scope
        Case TargetScopeSelection
            Set targetRange = ctx.SelectionRange
            scopeLabel = "selection"
        Case TargetScopeActiveSheet
            Set targetRange = ctx.WorksheetRef.UsedRange
            scopeLabel = "active sheet"
        Case TargetScopeWorkbook
            Set targetRange = Nothing
            scopeLabel = "workbook"
    End Select

    If scope = TargetScopeWorkbook Then
        If Not Infra_Interaction.Confirm( _
            "This operation will process the entire workbook and cannot be undone. Do you want to proceed?", _
            operationName & " - Workbook Operation Warning", vbDefaultButton2) Then
            GoTo CleanExit
        End If
    End If

    If Not targetRange Is Nothing Then
        If Not Infra_Undo.SaveStateOrConfirm(targetRange, operationName, captureMode) Then GoTo CleanExit
    End If

    ' --- 2. Configure app state ---
    guard.ScreenUpdating = False
    guard.EnableEvents = False
    guard.DisplayAlerts = True
    guard.Calculation = xlCalculationManual
    guard.UpdateState

    ' --- 3. Resolve worksheets ---
    Set sheetsToProcess = Infra_CommandSupport.ResolveWorksheetsToProcess(ctx, scope)
    If sheetsToProcess Is Nothing Then GoTo CleanExit

    ' --- 3b. Check worksheet protection in target scope ---
    Dim wsCheck As Worksheet
    Dim protectedSheetsList As String
    Dim wsItemCheck As Variant
    For Each wsItemCheck In sheetsToProcess
        Set wsCheck = wsItemCheck
        If wsCheck.ProtectContents Then
            If protectedSheetsList = "" Then
                protectedSheetsList = wsCheck.Name
            Else
                protectedSheetsList = protectedSheetsList & ", " & wsCheck.Name
            End If
        End If
    Next wsItemCheck
    
    If protectedSheetsList <> "" Then
        Infra_Interaction.ShowWarning "Cannot execute the operation because the following worksheet(s) in the target scope are protected: " & protectedSheetsList & ". Please unprotect them and try again.", operationName
        GoTo CleanExit
    End If

    ' --- 4. Start progress bar ---
    Infra_Progress.StartProgress operationName, sheetsToProcess.Count

    ' --- 5. Execute processing loop ---
    For Each wsItem In sheetsToProcess
        index = index + 1
        Set ws = wsItem

        ' Run worksheet level task
        totalProcessed = totalProcessed + processor.ProcessWorksheet(ws)

        ' Run range level task if applicable
        If scope = TargetScopeSelection Then
            totalProcessed = totalProcessed + processor.ProcessRange(ctx.SelectionRange)
        Else
            totalProcessed = totalProcessed + processor.ProcessRange(ws.UsedRange)
        End If

        Infra_Progress.UpdateProgress index, True
        If Infra_Progress.UserCancelled Then Exit For
    Next wsItem

    ' --- 6. Show completion message ---
    If showCompletionMsg Then
        Dim msg As String
        msg = "Done! Processed " & totalProcessed & " item(s) in the " & scopeLabel & "."
        Infra_Interaction.ShowInfo msg, operationName
    End If

    Execute = totalProcessed

CleanExit:
    Infra_Progress.EndProgress
    Exit Function

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "Infra_BatchProcessor.Execute", Err
    Resume CleanExit
End Function
