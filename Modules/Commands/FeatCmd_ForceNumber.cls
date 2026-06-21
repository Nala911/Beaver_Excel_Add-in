VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ForceNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_ForceNumber
' @Category: Feature
' @Description: Forces text-formatted numbers in the selection to become actual numeric values.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, Infra_ActionContext, Infra_Error, Infra_Undo, Infra_Progress, Infra_Interaction, Infra_Config

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyBulkWrite()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ForceNumber.Execute")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Dim selRange As Range
    Dim ws As Worksheet
    Dim textCells As Range
    Dim cell As Range
    Dim changeCount As Long
    Dim area As Range
    Dim vals As Variant
    Dim r As Long, c As Long
    Dim areaRows As Long, areaCols As Long
    Dim cellVal As Variant
    Dim formatRange As Range
    Dim valChanged As Boolean
    Dim outVal As Variant
    Dim cellAddr As String
    Dim isUniform As Boolean
    Dim areaFmtVal As Variant
    Dim areaNeedsGeneral As Boolean
    Dim addrList As String
    Dim startRow As Long
    Dim startCol As Long

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit

    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then
        GoTo CleanExit
    End If

    Set selRange = ctx.SelectionRange
    Set ws = ctx.WorksheetRef
    If ws Is Nothing Or selRange Is Nothing Then GoTo CleanExit

    ' Restrict selection to UsedRange for safety
    Set selRange = Infra_CommandSupport.GetSafeProcessingRange(selRange, 50000, True)
    If selRange Is Nothing Then GoTo CleanExit

    ' Identify cells with text constants
    If selRange.CountLarge = 1 Then
        If Not selRange.HasFormula And VarType(selRange.Value) = vbString Then
            Set textCells = selRange
        End If
    Else
        On Error Resume Next
        Set textCells = selRange.SpecialCells(xlCellTypeConstants, xlTextValues)
        On Error GoTo 0
    End If

    If textCells Is Nothing Then
        Infra_Interaction.ShowInfo "No text cells found in the selection.", "Convert Text to Number"
        GoTo CleanExit
    End If

    ' Save state for Undo
    If Not Infra_Undo.SaveStateOrConfirm(selRange, "Convert Text to Number") Then GoTo CleanExit

    Infra_Progress.StartProgress "Converting Text to Numbers", 100
    Infra_Progress.UpdateProgress 10

    changeCount = 0
    For Each area In textCells.Areas
        areaRows = area.Rows.Count
        areaCols = area.Columns.Count
        
        If areaRows = 1 And areaCols = 1 Then
            If Infra_ValueConversion.TryConvertToNumber(area.Value2, outVal) Then
                If InStr(1, area.NumberFormat, "@") > 0 Or LCase$(area.NumberFormat) = "text" Then
                    area.NumberFormat = "General"
                End If
                area.Value2 = outVal
                changeCount = changeCount + 1
            End If
        Else
            vals = area.Value2
            valChanged = False
            
            areaFmtVal = area.NumberFormat
            If Not IsNull(areaFmtVal) Then
                isUniform = True
                Dim areaFmtStr As String
                areaFmtStr = CStr(areaFmtVal)
                areaNeedsGeneral = (InStr(1, areaFmtStr, "@") > 0 Or LCase$(areaFmtStr) = "text")
            Else
                isUniform = False
                areaNeedsGeneral = False
            End If
            
            addrList = vbNullString
            Set formatRange = Nothing
            startRow = area.Row
            startCol = area.Column
            
            For r = 1 To areaRows
                For c = 1 To areaCols
                    cellVal = vals(r, c)
                    If Infra_ValueConversion.TryConvertToNumber(cellVal, outVal) Then
                        If Not isUniform Then
                            Set cell = area.Cells(r, c)
                            If InStr(1, cell.NumberFormat, "@") > 0 Or LCase$(cell.NumberFormat) = "text" Then
                                cellAddr = Infra_CommandSupport.GetA1Address(startRow + r - 1, startCol + c - 1)
                                Set formatRange = Infra_CommandSupport.AccumulateUnion(formatRange, area.Parent, addrList, cellAddr)
                            End If
                        End If
                        vals(r, c) = outVal
                        valChanged = True
                        changeCount = changeCount + 1
                    End If
                Next c
            Next r
            
            If valChanged Then
                If isUniform Then
                    If areaNeedsGeneral Then
                        area.NumberFormat = "General"
                    End If
                Else
                    Set formatRange = Infra_CommandSupport.AccumulateUnion(formatRange, area.Parent, addrList, "", True)
                    If Not formatRange Is Nothing Then
                        formatRange.NumberFormat = "General"
                    End If
                End If
                area.Value2 = vals
            End If
        End If
    Next area

    Infra_Progress.UpdateProgress 100
    Infra_Progress.EndProgress

    If changeCount > 0 Then
        Infra_Interaction.ShowInfo "Successfully converted " & changeCount & " cells to numbers.", "Convert Text to Number"
    Else
        Infra_Interaction.ShowInfo "No cells required conversion.", "Convert Text to Number"
    End If

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Progress.EndProgress
    Infra_Error.HandleError "ForceNumber.Execute", Err
    Resume CleanExit
End Sub

