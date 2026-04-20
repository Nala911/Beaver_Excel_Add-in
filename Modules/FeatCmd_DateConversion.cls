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

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("DateConversion.CanExecute")
    On Error GoTo ErrHandler
    ICommand_CanExecute = Infra_AppState.IsRangeSelected()
    Exit Function
ErrHandler:
    Infra_Error.HandleError "DateConversion.CanExecute", Err
    ICommand_CanExecute = False
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
    If Not Infra_AppState.IsRangeSelected() Then GoTo CleanExit
    Set rng = Selection
    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit

    ' 2. Safety Check: single column only
    If rng.Columns.Count > 1 Then
        MsgBox "Please select only one column of dates at a time.", vbExclamation, Infra_Config.ADDIN_NAME
        GoTo CleanExit
    End If

    ' 3. Show screen so user can see context, then ask for the month
    Application.ScreenUpdating = True
    targetMonth = Infra_UIFactory.PromptForDateConversionMonth(ctx)
    If targetMonth = 0 Then GoTo CleanExit
    Application.ScreenUpdating = False

    ' 4. Read data into memory array
    If rng.Cells.Count = 1 Then
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = rng.Value
    Else
        dataArr = rng.Value
    End If

    ' 5. Loop through the array (process in memory)
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

    ' 6. Dump array back to Excel (one single write operation)
    rng.Value = dataArr
    rng.NumberFormat = Infra_Config.DEFAULT_DATE_FORMAT

    MsgBox "Conversion Complete! " & convertedCount & " cell(s) updated.", vbInformation, Infra_Config.ADDIN_NAME

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "DateConversion.Execute", Err
    Resume CleanExit
End Sub
