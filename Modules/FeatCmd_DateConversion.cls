VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_DateConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_DateConversion
' @Category: Feature
' @Description: Converts text strings to proper Excel Date values with ambiguous format resolution.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.CreateExecutionPolicy(True, False, True)
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("DateConversion.Validate")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)

    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then
        Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("Please select a range before running this command.")
    ElseIf ctx.SelectionRange.Columns.Count > 1 Then
        Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("Please select only one column of dates at a time.")
    Else
        Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()
    End If

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "DateConversion.Validate", Err
    Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("The selected dates cannot be processed right now.")
    Resume CleanExit
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("DateConversion.Execute")
    On Error GoTo ErrHandler

    Dim rng As Range
    Dim dataArr As Variant
    Dim i As Long
    Dim txt As String
    Dim targetMonth As Integer
    Dim ctx As Infra_ActionContext
    Dim convertedCount As Long

    ' 1. Validate selection
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set rng = ctx.SelectionRange

    ' 2. Ask for the month used to resolve ambiguous dates
    targetMonth = Infra_UIFactory.PromptForDateConversionMonth(ctx)
    If targetMonth = 0 Then GoTo CleanExit

    ' 3. Read data into memory array
    If rng.Cells.Count = 1 Then
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = rng.Value
    Else
        dataArr = rng.Value
    End If

    ' 4. Loop through the array (process in memory)
    For i = LBound(dataArr, 1) To UBound(dataArr, 1)
        txt = Trim(CStr(dataArr(i, 1)))

        If Len(txt) > 0 Then
            Dim resolvedDate As Date
            If Lib_DateUtils.TryResolveDateWithMonth(txt, CInt(targetMonth), resolvedDate) Then
                dataArr(i, 1) = resolvedDate
                convertedCount = convertedCount + 1
            End If
        End If
    Next i

    ' 5. Dump array back to Excel (one single write operation)
    rng.Value = dataArr
    rng.NumberFormat = Infra_Config.DEFAULT_DATE_FORMAT

    Infra_Interaction.ShowInfo "Conversion Complete! " & convertedCount & " cell(s) updated."

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "DateConversion.Execute", Err
    Resume CleanExit
End Sub
