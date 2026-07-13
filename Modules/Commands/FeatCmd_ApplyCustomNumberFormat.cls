VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ApplyCustomNumberFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_ApplyCustomNumberFormat
' @Category: Feature
' @Description: Prompts for a custom number format code and applies it to the selected range.

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyBulkWrite()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateCanModifySelection(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ApplyCustomNumberFormat.Execute")
    On Error GoTo ErrHandler
    
    Dim targetRange As Range
    Dim safeRange As Range
    Dim ctx As ActionContext

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If Not Infra_CommandSupport.HasRangeSelectionInContext(ctx) Then GoTo CleanExit
    Set targetRange = ctx.SelectionRange
    If targetRange Is Nothing Then GoTo CleanExit

    Set safeRange = Infra_CommandSupport.GetSafeProcessingRange(targetRange, 50000, True)
    If safeRange Is Nothing Then GoTo CleanExit

    ' Determine default pre-populated format
    Dim defaultFormat As String
    defaultFormat = Infra_Config.Model.DefaultNumberFormat

    Dim selectionFormat As Variant
    selectionFormat = safeRange.NumberFormat
    
    If Not IsNull(selectionFormat) Then
        defaultFormat = CStr(selectionFormat)
    Else
        ' If the range has mixed formats, fall back to the first cell's format
        Dim firstCellFormat As Variant
        firstCellFormat = safeRange.Cells(1, 1).NumberFormat
        If Not IsNull(firstCellFormat) Then
            defaultFormat = CStr(firstCellFormat)
        End If
    End If

    ' Determine custom format code
    Dim formatCode As String
    Dim isTestMock As Boolean
    
    isTestMock = False
    If Not context Is Nothing Then
        If context.SourceName Like "TEST_FORMAT:*" Then
            formatCode = Mid$(context.SourceName, 13)
            isTestMock = True
        End If
    End If

    If Not isTestMock Then
        If Not Infra_Interaction.PromptText( _
            "Enter custom number format code (e.g. #,##0.0 "" kwh""):", _
            "Custom Number Format", _
            defaultFormat, _
            formatCode) Then
            GoTo CleanExit
        End If
    End If

    ' Apply format code with undo support
    If Not Infra_Undo.SaveStateOrConfirm(safeRange, "Apply Custom Format", UndoCaptureFormatOnly) Then GoTo CleanExit
    safeRange.NumberFormat = formatCode

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ApplyCustomNumberFormat.Execute", Err
    Resume CleanExit
End Sub
