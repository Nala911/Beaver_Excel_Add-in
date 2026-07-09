VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_FinancialModelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand
Implements IBatchProcessor

' @Module: FeatCmd_FinancialModelling
' @Category: Feature
' @Description: Auto-formatter applying the industry standard financial modeling coloring scheme to cells.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, IBatchProcessor, ActionContext, ScopedRequest, Infra_Error, Infra_Undo, Infra_Progress, Infra_Interaction, Infra_CommandSupport, Infra_BatchProcessor

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorksheet(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("FinancialModelling.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As ActionContext
    Dim stateInfo As Collection

    Set stateInfo = Infra_Error.NewStateInfo()

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Or ctx.WorkbookRef Is Nothing Then GoTo CleanExit
    
    Infra_Error.AddState stateInfo, "Workbook", ctx.WorkbookRef.Name
    Infra_Error.AddState stateInfo, "Worksheet", ctx.WorksheetRef.Name
    If ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing Then
        Infra_Error.AddState stateInfo, "Selection", ctx.SelectionRange.Address(False, False)
    End If

    ' Scope determination: If selection has > 1 cell, use TargetScopeSelection. Else TargetScopeActiveSheet.
    Dim scopeVal As TargetScope
    If ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing Then
        If ctx.SelectionRange.Cells.CountLarge > 1 Then
            scopeVal = TargetScopeSelection
        Else
            scopeVal = TargetScopeActiveSheet
        End If
    Else
        scopeVal = TargetScopeActiveSheet
    End If

    ' Run the batch processor
    Dim totalProcessed As Long
    totalProcessed = Infra_BatchProcessor.Execute(ctx, scopeVal, Me, "Financial Modelling", False)

    Dim scopeLabel As String
    Select Case scopeVal
        Case TargetScopeSelection: scopeLabel = "selection"
        Case TargetScopeActiveSheet: scopeLabel = "active sheet"
    End Select

    Dim msg As String
    msg = "Done! Formatted cells in the " & scopeLabel & " using the financial modelling color scheme."
    Infra_Interaction.ShowInfo msg, "Financial Modelling"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "FinancialModelling.Execute", Err, stateInfo
    Resume CleanExit
End Sub

Private Function IBatchProcessor_ProcessWorksheet(ByVal ws As Worksheet) As Long
    IBatchProcessor_ProcessWorksheet = 0
End Function

Private Function IBatchProcessor_ProcessRange(ByVal targetRange As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("IBatchProcessor_ProcessRange")
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then
        IBatchProcessor_ProcessRange = 0
        GoTo CleanExit
    End If
    
    ' Intersect with UsedRange if targetRange is large to prevent scanning millions of empty cells
    Dim processingRange As Range
    Set processingRange = Infra_CommandSupport.GetSafeProcessingRange(targetRange, 100000, False)
    If processingRange Is Nothing Then
        IBatchProcessor_ProcessRange = 0
        GoTo CleanExit
    End If
    
    IBatchProcessor_ProcessRange = FormatFinancialModelling(processingRange)
    
CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "IBatchProcessor_ProcessRange", Err
    Resume CleanExit
End Function

Private Function FormatFinancialModelling(ByVal area As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("FormatFinancialModelling")
    On Error GoTo ErrHandler

    Dim totalFormatted As Long
    Dim subArea As Range
    
    For Each subArea In area.Areas
        ' 1. Process Formula cells in the subArea
        Dim formulaCells As Range
        On Error Resume Next
        If subArea.Cells.CountLarge = 1 Then
            If subArea.HasFormula Then Set formulaCells = subArea
        Else
            Set formulaCells = subArea.SpecialCells(xlCellTypeFormulas)
        End If
        On Error GoTo ErrHandler
        
        If Not formulaCells Is Nothing Then
            totalFormatted = totalFormatted + FormatFormulasInArea(formulaCells)
        End If
        
        ' 2. Process Constant cells in the subArea
        Dim constantCells As Range
        On Error Resume Next
        If subArea.Cells.CountLarge = 1 Then
            If Not subArea.HasFormula And Not IsEmpty(subArea.Value) Then Set constantCells = subArea
        Else
            Set constantCells = subArea.SpecialCells(xlCellTypeConstants)
        End If
        On Error GoTo ErrHandler
        
        If Not constantCells Is Nothing Then
            totalFormatted = totalFormatted + FormatConstantsInArea(constantCells)
        End If
        
        DoEvents
        If Infra_Progress.UserCancelled Then Exit For
    Next subArea

    FormatFinancialModelling = totalFormatted

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "FormatFinancialModelling", Err
    Resume CleanExit
End Function

Private Function FormatFormulasInArea(ByVal formulaCells As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("FormatFormulasInArea")
    On Error GoTo ErrHandler

    Dim parentSheet As Worksheet
    Set parentSheet = formulaCells.Parent
    
    Dim redUnion As Range, greenUnion As Range, blackUnion As Range
    Dim redAddrList As String, greenAddrList As String, blackAddrList As String
    Dim cellAddr As String
    
    Dim countProcessed As Long
    Dim area As Range
    
    For Each area In formulaCells.Areas
        If area.Cells.CountLarge = 1 Then
            Dim fVal As Variant
            fVal = area.Formula2
            
            If VarType(fVal) = vbString Then
                Dim formulaText As String
                formulaText = fVal
                
                If formulaText <> vbNullString Then
                    cellAddr = Infra_CommandSupport.GetA1Address(area.Row, area.Column)
                    If InStr(1, formulaText, "[", vbTextCompare) > 0 Then
                        ' External Reference
                        Set redUnion = Infra_CommandSupport.AccumulateUnion(redUnion, parentSheet, redAddrList, cellAddr)
                        If Not redUnion Is Nothing Then
                            If redUnion.Areas.Count >= 500 Then
                                redUnion.Font.Color = RGB(180, 0, 0)
                                Set redUnion = Nothing
                            End If
                        End If
                    ElseIf InStr(1, formulaText, "!", vbTextCompare) > 0 Then
                        ' Internal Worksheet Reference
                        Set greenUnion = Infra_CommandSupport.AccumulateUnion(greenUnion, parentSheet, greenAddrList, cellAddr)
                        If Not greenUnion Is Nothing Then
                            If greenUnion.Areas.Count >= 500 Then
                                greenUnion.Font.Color = RGB(0, 128, 0)
                                Set greenUnion = Nothing
                            End If
                        End If
                    Else
                        ' Local Formula
                        Set blackUnion = Infra_CommandSupport.AccumulateUnion(blackUnion, parentSheet, blackAddrList, cellAddr)
                        If Not blackUnion Is Nothing Then
                            If blackUnion.Areas.Count >= 500 Then
                                blackUnion.Font.Color = RGB(0, 0, 0)
                                Set blackUnion = Nothing
                            End If
                        End If
                    End If
                    countProcessed = countProcessed + 1
                End If
            End If
        Else
            Dim formulas As Variant
            formulas = area.Formula2
            
            Dim r As Long, c As Long
            Dim rMin As Long, rMax As Long
            Dim cMin As Long, cMax As Long
            
            rMin = LBound(formulas, 1)
            rMax = UBound(formulas, 1)
            cMin = LBound(formulas, 2)
            cMax = UBound(formulas, 2)
            
            For r = rMin To rMax
                For c = cMin To cMax
                    Dim cellFormula As Variant
                    cellFormula = formulas(r, c)
                    
                    If VarType(cellFormula) = vbString Then
                        Dim fText As String
                        fText = cellFormula
                        
                        If fText <> vbNullString Then
                            cellAddr = Infra_CommandSupport.GetA1Address(area.Row + r - 1, area.Column + c - 1)
                            
                            If InStr(1, fText, "[", vbTextCompare) > 0 Then
                                Set redUnion = Infra_CommandSupport.AccumulateUnion(redUnion, parentSheet, redAddrList, cellAddr)
                                If Not redUnion Is Nothing Then
                                    If redUnion.Areas.Count >= 500 Then
                                        redUnion.Font.Color = RGB(180, 0, 0)
                                        Set redUnion = Nothing
                                    End If
                                End If
                            ElseIf InStr(1, fText, "!", vbTextCompare) > 0 Then
                                Set greenUnion = Infra_CommandSupport.AccumulateUnion(greenUnion, parentSheet, greenAddrList, cellAddr)
                                If Not greenUnion Is Nothing Then
                                    If greenUnion.Areas.Count >= 500 Then
                                        greenUnion.Font.Color = RGB(0, 128, 0)
                                        Set greenUnion = Nothing
                                    End If
                                End If
                            Else
                                Set blackUnion = Infra_CommandSupport.AccumulateUnion(blackUnion, parentSheet, blackAddrList, cellAddr)
                                If Not blackUnion Is Nothing Then
                                    If blackUnion.Areas.Count >= 500 Then
                                        blackUnion.Font.Color = RGB(0, 0, 0)
                                        Set blackUnion = Nothing
                                    End If
                                End If
                            End If
                            countProcessed = countProcessed + 1
                        End If
                    End If
                Next c
            Next r
        End If
    Next area
    
    ' Flush the remaining unions
    Set redUnion = Infra_CommandSupport.AccumulateUnion(redUnion, parentSheet, redAddrList, "", True)
    If Not redUnion Is Nothing Then
        redUnion.Font.Color = RGB(180, 0, 0)
    End If
    
    Set greenUnion = Infra_CommandSupport.AccumulateUnion(greenUnion, parentSheet, greenAddrList, "", True)
    If Not greenUnion Is Nothing Then
        greenUnion.Font.Color = RGB(0, 128, 0)
    End If
    
    Set blackUnion = Infra_CommandSupport.AccumulateUnion(blackUnion, parentSheet, blackAddrList, "", True)
    If Not blackUnion Is Nothing Then
        blackUnion.Font.Color = RGB(0, 0, 0)
    End If

    FormatFormulasInArea = countProcessed

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "FormatFormulasInArea", Err
    Resume CleanExit
End Function

Private Function FormatConstantsInArea(ByVal constantCells As Range) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("FormatConstantsInArea")
    On Error GoTo ErrHandler

    Dim parentSheet As Worksheet
    Set parentSheet = constantCells.Parent
    
    Dim blueUnion As Range, blackUnion As Range
    Dim blueAddrList As String, blackAddrList As String
    Dim cellAddr As String
    
    Dim countProcessed As Long
    Dim area As Range
    
    For Each area In constantCells.Areas
        If area.Cells.CountLarge = 1 Then
            Dim val As Variant
            val = area.Value2
            
            If Not IsEmpty(val) And Not IsError(val) Then
                cellAddr = Infra_CommandSupport.GetA1Address(area.Row, area.Column)
                If VarType(val) <> vbBoolean And IsNumeric(val) Then
                    ' Hardcoded Numeric Value (Blue)
                    Set blueUnion = Infra_CommandSupport.AccumulateUnion(blueUnion, parentSheet, blueAddrList, cellAddr)
                    If Not blueUnion Is Nothing Then
                        If blueUnion.Areas.Count >= 500 Then
                            blueUnion.Font.Color = RGB(0, 0, 255)
                            Set blueUnion = Nothing
                        End If
                    End If
                Else
                    ' Hardcoded Text/Labels (Black)
                    Set blackUnion = Infra_CommandSupport.AccumulateUnion(blackUnion, parentSheet, blackAddrList, cellAddr)
                    If Not blackUnion Is Nothing Then
                        If blackUnion.Areas.Count >= 500 Then
                            blackUnion.Font.Color = RGB(0, 0, 0)
                            Set blackUnion = Nothing
                        End If
                    End If
                End If
                countProcessed = countProcessed + 1
            End If
        Else
            Dim values As Variant
            values = area.Value2
            
            Dim r As Long, c As Long
            Dim rMin As Long, rMax As Long
            Dim cMin As Long, cMax As Long
            
            rMin = LBound(values, 1)
            rMax = UBound(values, 1)
            cMin = LBound(values, 2)
            cMax = UBound(values, 2)
            
            For r = rMin To rMax
                For c = cMin To cMax
                    Dim itemVal As Variant
                    itemVal = values(r, c)
                    
                    If Not IsEmpty(itemVal) And Not IsError(itemVal) Then
                        cellAddr = Infra_CommandSupport.GetA1Address(area.Row + r - 1, area.Column + c - 1)
                        
                        If VarType(itemVal) <> vbBoolean And IsNumeric(itemVal) Then
                            ' Hardcoded Numeric Value (Blue)
                            Set blueUnion = Infra_CommandSupport.AccumulateUnion(blueUnion, parentSheet, blueAddrList, cellAddr)
                            If Not blueUnion Is Nothing Then
                                If blueUnion.Areas.Count >= 500 Then
                                    blueUnion.Font.Color = RGB(0, 0, 255)
                                    Set blueUnion = Nothing
                                End If
                            End If
                        Else
                            ' Hardcoded Text/Labels (Black)
                            Set blackUnion = Infra_CommandSupport.AccumulateUnion(blackUnion, parentSheet, blackAddrList, cellAddr)
                            If Not blackUnion Is Nothing Then
                                If blackUnion.Areas.Count >= 500 Then
                                    blackUnion.Font.Color = RGB(0, 0, 0)
                                    Set blackUnion = Nothing
                                End If
                            End If
                        End If
                        countProcessed = countProcessed + 1
                    End If
                Next c
            Next r
        End If
    Next area
    
    ' Flush remaining unions
    Set blueUnion = Infra_CommandSupport.AccumulateUnion(blueUnion, parentSheet, blueAddrList, "", True)
    If Not blueUnion Is Nothing Then
        blueUnion.Font.Color = RGB(0, 0, 255)
    End If
    
    Set blackUnion = Infra_CommandSupport.AccumulateUnion(blackUnion, parentSheet, blackAddrList, "", True)
    If Not blackUnion Is Nothing Then
        blackUnion.Font.Color = RGB(0, 0, 0)
    End If

    FormatConstantsInArea = countProcessed

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "FormatConstantsInArea", Err
    Resume CleanExit
End Function
