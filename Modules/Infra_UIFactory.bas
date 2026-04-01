Attribute VB_Name = "Infra_UIFactory"
Option Explicit

' @Module: Infra_UIFactory
' @Category: Infrastructure
' @Description: Centralized factory for creating and displaying standardized user prompts.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Config, Infra_CleanDataRequest, Infra_ExportRequest, Infra_StaticRequest, Infra_ActionContext

' Shows the Clean Data options via InputBox and returns a populated Request object.
Public Function ShowCleanDataDialog(ByVal ctx As Infra_ActionContext) As Infra_CleanDataRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowCleanDataDialog")
    On Error GoTo ErrHandler
    
    Dim promptMsg As String
    Dim userChoice As Variant
    Dim normalizedChoice As String
    Dim request As Infra_CleanDataRequest

    Do
        promptMsg = "Clean text values using Excel TRIM and CLEAN." & vbCrLf & vbCrLf & _
                    BuildContextSummary(ctx, True) & vbCrLf & vbCrLf & _
                    "Choose a scope:" & vbCrLf & _
                    "Range     - Clean only the current selection" & vbCrLf & _
                    "Sheet     - Clean the active worksheet" & vbCrLf & _
                    "Workbook  - Clean every worksheet" & vbCrLf & vbCrLf & _
                    "Type Range, Sheet, or Workbook."

        userChoice = Application.InputBox(promptMsg, BuildDialogTitle("Clean Data"), "Range", Type:=2)
        If IsDialogCancelled(userChoice) Then GoTo CleanExit

        normalizedChoice = NormalizeChoiceText(userChoice)
        Select Case normalizedChoice
            Case "", "R", "RANGE", "SELECTED", "SELECTION"
                Set request = New Infra_CleanDataRequest
                Set request.Context = ctx
                request.Scope = 1
                Set ShowCleanDataDialog = request
                Exit Do
            Case "S", "SHEET", "ACTIVE SHEET", "ACTIVESHEET"
                Set request = New Infra_CleanDataRequest
                Set request.Context = ctx
                request.Scope = 2
                Set ShowCleanDataDialog = request
                Exit Do
            Case "W", "WB", "WORKBOOK", "WHOLE WORKBOOK", "WHOLEWORKBOOK"
                Set request = New Infra_CleanDataRequest
                Set request.Context = ctx
                request.Scope = 3
                Set ShowCleanDataDialog = request
                Exit Do
            Case Else
                MsgBox "Please type Range, Sheet, or Workbook.", vbExclamation, BuildDialogTitle("Clean Data")
        End Select
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowCleanDataDialog", Err
    Resume CleanExit
End Function

' Shows the Export options via MsgBox/InputBox and returns a populated Request object.
Public Function ShowExportDialog(ByVal ctx As Infra_ActionContext) As Infra_ExportRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowExportDialog")
    On Error GoTo ErrHandler
    
    Dim request As Infra_ExportRequest
    Dim exportChoice As Variant
    Dim normalizedChoice As String
    Dim scaleInput As Variant
    
    Set request = New Infra_ExportRequest
    Set request.Context = ctx
    Set request.SourceRange = ResolveExportRange(ctx)
    request.ScaleFactor = Infra_Config.DEFAULT_EXPORT_SCALE
    
    If request.SourceRange Is Nothing Then
        MsgBox "No data found on the active sheet to export.", vbExclamation, Infra_Config.ADDIN_NAME
        GoTo CleanExit
    End If
    
    Application.ScreenUpdating = True

    Do
        exportChoice = Application.InputBox( _
            "Export the selected content to your Desktop." & vbCrLf & vbCrLf & _
            BuildExportSummary(request.SourceRange) & vbCrLf & vbCrLf & _
            "Choose a format:" & vbCrLf & _
            "PNG - High-resolution image" & vbCrLf & _
            "PDF - Print-ready document" & vbCrLf & vbCrLf & _
            "Type PNG or PDF.", _
            BuildDialogTitle("Export"), "PNG", Type:=2)
        If IsDialogCancelled(exportChoice) Then GoTo CleanExit

        normalizedChoice = NormalizeChoiceText(exportChoice)
        Select Case normalizedChoice
            Case "", "PNG", "IMAGE"
                request.ExportAsPng = True
                Exit Do
            Case "PDF"
                request.ExportAsPng = False
                Exit Do
            Case Else
                MsgBox "Please type PNG or PDF.", vbExclamation, BuildDialogTitle("Export")
        End Select
    Loop

    If request.ExportAsPng Then
        request.ScaleFactor = PromptForExportScale(request.ScaleFactor)
        If request.ScaleFactor = 0 Then GoTo CleanExit
    End If
    
    Set ShowExportDialog = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowExportDialog", Err
    Resume CleanExit
End Function

' Shows the conversion scope dialog for formula-to-value actions.
Public Function ShowStaticConversionDialog(ByVal ctx As Infra_ActionContext) As Infra_StaticRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowStaticConversionDialog")
    On Error GoTo ErrHandler

    Dim request As Infra_StaticRequest
    Dim scopeChoice As Variant
    Dim normalizedChoice As String

    Do
        scopeChoice = Application.InputBox( _
            "Convert formulas into their current values." & vbCrLf & vbCrLf & _
            BuildContextSummary(ctx, False) & vbCrLf & vbCrLf & _
            "This is intended for permanent conversion." & vbCrLf & _
            "Choose a scope:" & vbCrLf & _
            "Sheet     - Convert formulas on the active sheet" & vbCrLf & _
            "Workbook  - Convert formulas on every worksheet" & vbCrLf & vbCrLf & _
            "Type Sheet or Workbook.", _
            BuildDialogTitle("Make Static"), "Sheet", Type:=2)
        If IsDialogCancelled(scopeChoice) Then GoTo CleanExit

        normalizedChoice = NormalizeChoiceText(scopeChoice)
        If normalizedChoice = "" Then normalizedChoice = "SHEET"

        Select Case normalizedChoice
            Case "S", "SHEET", "ACTIVE SHEET", "ACTIVESHEET"
                Set request = New Infra_StaticRequest
                Set request.Context = ctx
                request.Scope = 1
                Set ShowStaticConversionDialog = request
                Exit Do
            Case "W", "WB", "WORKBOOK", "WHOLE WORKBOOK", "WHOLEWORKBOOK"
                If MsgBox( _
                    "You are about to convert formulas on every worksheet in " & SafeWorkbookName(ctx) & "." & vbCrLf & vbCrLf & _
                    "This is not reversible as a single workbook-wide undo action." & vbCrLf & vbCrLf & _
                    "Continue with workbook-wide conversion?", _
                    vbOKCancel + vbExclamation + vbDefaultButton2, BuildDialogTitle("Confirm Workbook Scope")) <> vbOK Then
                    GoTo CleanExit
                End If
                Set request = New Infra_StaticRequest
                Set request.Context = ctx
                request.Scope = 2
                Set ShowStaticConversionDialog = request
                Exit Do
            Case Else
                MsgBox "Please type Sheet or Workbook.", vbExclamation, BuildDialogTitle("Make Static")
        End Select
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowStaticConversionDialog", Err
    Resume CleanExit
End Function

Public Function ShowBreakLinksScopeDialog(ByVal ctx As Infra_ActionContext, ByVal linkInfo As String) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowBreakLinksScopeDialog")
    On Error GoTo ErrHandler

    Dim userChoice As Variant
    Dim normalizedChoice As String

    Do
        userChoice = Application.InputBox( _
            "External links were found and can be permanently converted to values." & vbCrLf & vbCrLf & _
            BuildContextSummary(ctx, False) & vbCrLf & vbCrLf & _
            "Detected items:" & vbCrLf & linkInfo & vbCrLf & vbCrLf & _
            "Choose a scope:" & vbCrLf & _
            "Sheet     - Process only the active sheet" & vbCrLf & _
            "Workbook  - Process the whole workbook" & vbCrLf & vbCrLf & _
            "Type Sheet or Workbook.", _
            BuildDialogTitle("Break External Links"), "Sheet", Type:=2)
        If IsDialogCancelled(userChoice) Then GoTo CleanExit

        normalizedChoice = NormalizeChoiceText(userChoice)
        If normalizedChoice = "" Then normalizedChoice = "SHEET"

        Select Case normalizedChoice
            Case "S", "SHEET", "ACTIVE SHEET", "ACTIVESHEET"
                ShowBreakLinksScopeDialog = 1
                Exit Do
            Case "W", "WB", "WORKBOOK", "WHOLE WORKBOOK", "WHOLEWORKBOOK"
                If MsgBox( _
                    "This will remove workbook-level links and connections and flatten external content." & vbCrLf & vbCrLf & _
                    "Continue with whole-workbook processing?", _
                    vbOKCancel + vbExclamation + vbDefaultButton2, BuildDialogTitle("Confirm Workbook Scope")) <> vbOK Then
                    GoTo CleanExit
                End If
                ShowBreakLinksScopeDialog = 2
                Exit Do
            Case Else
                MsgBox "Please type Sheet or Workbook.", vbExclamation, BuildDialogTitle("Break External Links")
        End Select
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowBreakLinksScopeDialog", Err
    Resume CleanExit
End Function

Public Function PromptForDateConversionMonth(ByVal ctx As Infra_ActionContext) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForDateConversionMonth")
    On Error GoTo ErrHandler

    Dim userInput As Variant
    Dim monthValue As Long

    Do
        userInput = Application.InputBox( _
            "Convert text dates in the selected column into real Excel dates." & vbCrLf & vbCrLf & _
            BuildContextSummary(ctx, True) & vbCrLf & vbCrLf & _
            "Enter the month to use when dates are ambiguous." & vbCrLf & _
            "Examples: 9, 09, Sep, September", _
            BuildDialogTitle("Date Conversion"), MonthName(Month(Date), True), Type:=2)
        If IsDialogCancelled(userInput) Then GoTo CleanExit

        monthValue = ParseMonthValue(userInput)
        If monthValue >= 1 And monthValue <= 12 Then
            PromptForDateConversionMonth = monthValue
            Exit Do
        End If

        MsgBox "Please enter a valid month name or number from 1 to 12.", vbExclamation, BuildDialogTitle("Date Conversion")
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForDateConversionMonth", Err
    Resume CleanExit
End Function

Public Function PromptForDuplicateName(ByVal ctx As Infra_ActionContext, ByVal suggestedBaseName As String) As Variant
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForDuplicateName")
    On Error GoTo ErrHandler

    Dim userInput As Variant
    Dim fileName As String

    Do
        userInput = Application.InputBox( _
            "Create a macro-free copy of the current workbook on your Desktop." & vbCrLf & vbCrLf & _
            BuildContextSummary(ctx, False) & vbCrLf & vbCrLf & _
            "Enter the new file name." & vbCrLf & _
            "The .xlsx extension will be added automatically.", _
            BuildDialogTitle("Create Duplicate"), suggestedBaseName, Type:=2)
        If IsDialogCancelled(userInput) Then
            PromptForDuplicateName = False
            GoTo CleanExit
        End If

        fileName = Trim$(CStr(userInput))
        If fileName = vbNullString Then fileName = suggestedBaseName

        If IsValidWindowsFileName(fileName) Then
            PromptForDuplicateName = fileName
            Exit Do
        End If

        MsgBox "The file name contains characters Windows cannot save. Avoid: \ / : * ? "" < > |", _
               vbExclamation, BuildDialogTitle("Create Duplicate")
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForDuplicateName", Err
    PromptForDuplicateName = False
    Resume CleanExit
End Function

Public Function PromptForWrapFormulaPattern(ByVal ctx As Infra_ActionContext, ByVal lastPattern As String, ByVal placeholder As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForWrapFormulaPattern")
    On Error GoTo ErrHandler

    Dim userInput As Variant

    Do
        userInput = Application.InputBox( _
            "Wrap selected formulas or values with a new formula pattern." & vbCrLf & vbCrLf & _
            BuildContextSummary(ctx, True) & vbCrLf & vbCrLf & _
            "Use " & placeholder & " where the existing cell content should go." & vbCrLf & _
            "Example: =ROUND(" & placeholder & ", 0)", _
            BuildDialogTitle("Wrap Formula"), lastPattern, Type:=2)
        If IsDialogCancelled(userInput) Then GoTo CleanExit

        PromptForWrapFormulaPattern = Trim$(CStr(userInput))
        If PromptForWrapFormulaPattern = vbNullString Then GoTo CleanExit

        If InStr(1, PromptForWrapFormulaPattern, placeholder, vbTextCompare) > 0 Then Exit Do

        MsgBox "Your formula pattern must include the placeholder " & placeholder & ".", _
               vbExclamation, BuildDialogTitle("Wrap Formula")
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForWrapFormulaPattern", Err
    Resume CleanExit
End Function

Private Function ResolveExportRange(ByVal ctx As Infra_ActionContext) As Range
    If ctx Is Nothing Then Exit Function
    If ctx.WorksheetRef Is Nothing Then Exit Function
    
    If Not ctx.HasRangeSelection Or ctx.SelectionRange Is Nothing Then
        Set ResolveExportRange = ctx.WorksheetRef.UsedRange
    ElseIf ctx.SelectionRange.Cells.Count = 1 Or ctx.SelectionRange.Width = 0 Or ctx.SelectionRange.Height = 0 Then
        Set ResolveExportRange = ctx.WorksheetRef.UsedRange
    Else
        Set ResolveExportRange = ctx.SelectionRange
    End If
End Function

Private Function NormalizeExportScale(ByVal scaleInput As String) As Long
    If IsNumeric(scaleInput) Then
        NormalizeExportScale = CLng(scaleInput)
    Else
        NormalizeExportScale = Infra_Config.DEFAULT_EXPORT_SCALE
    End If

    If NormalizeExportScale < 1 Then NormalizeExportScale = 1
    If NormalizeExportScale > Infra_Config.MAX_EXPORT_SCALE Then
        NormalizeExportScale = Infra_Config.MAX_EXPORT_SCALE
    End If
End Function

Private Function PromptForExportScale(ByVal defaultScale As Long) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("PromptForExportScale")
    On Error GoTo ErrHandler

    Dim scaleInput As Variant
    Dim normalizedScale As Long

    Do
        scaleInput = Application.InputBox( _
            "Choose the PNG scale factor." & vbCrLf & vbCrLf & _
            "1 = Smaller file" & vbCrLf & _
            CStr(Infra_Config.DEFAULT_EXPORT_SCALE) & " = Balanced default" & vbCrLf & _
            CStr(Infra_Config.MAX_EXPORT_SCALE) & " = Largest supported image" & vbCrLf & vbCrLf & _
            "Enter a number from 1 to " & Infra_Config.MAX_EXPORT_SCALE & ".", _
            BuildDialogTitle("PNG Quality"), CStr(defaultScale), Type:=2)
        If IsDialogCancelled(scaleInput) Then GoTo CleanExit

        If Trim$(CStr(scaleInput)) = vbNullString Then scaleInput = CStr(defaultScale)
        If IsNumeric(scaleInput) Then
            normalizedScale = CLng(scaleInput)
            If normalizedScale >= 1 And normalizedScale <= Infra_Config.MAX_EXPORT_SCALE Then
                PromptForExportScale = normalizedScale
                Exit Do
            End If
        End If

        MsgBox "Please enter a whole number from 1 to " & Infra_Config.MAX_EXPORT_SCALE & ".", _
               vbExclamation, BuildDialogTitle("PNG Quality")
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForExportScale", Err
    Resume CleanExit
End Function

Private Function BuildDialogTitle(ByVal dialogName As String) As String
    BuildDialogTitle = Infra_Config.ADDIN_NAME & " - " & dialogName
End Function

Private Function BuildContextSummary(ByVal ctx As Infra_ActionContext, Optional ByVal includeSelection As Boolean = True) As String
    Dim summary As String

    If ctx Is Nothing Then Exit Function

    summary = "Workbook: " & SafeWorkbookName(ctx) & vbCrLf & _
              "Sheet: " & SafeWorksheetName(ctx)

    If includeSelection Then
        summary = summary & vbCrLf & "Selection: " & SafeSelectionAddress(ctx)
    End If

    BuildContextSummary = summary
End Function

Private Function BuildExportSummary(ByVal sourceRange As Range) As String
    Dim summary As String

    If sourceRange Is Nothing Then Exit Function

    summary = "Range: " & sourceRange.Address(False, False) & vbCrLf & _
              "Sheet: " & sourceRange.Worksheet.Name & vbCrLf & _
              "Size: " & Format(sourceRange.Rows.Count, "#,##0") & " row(s) x " & _
              Format(sourceRange.Columns.Count, "#,##0") & " column(s)"

    BuildExportSummary = summary
End Function

Private Function NormalizeChoiceText(ByVal rawValue As Variant) As String
    NormalizeChoiceText = UCase$(Trim$(CStr(rawValue)))
End Function

Private Function IsDialogCancelled(ByVal response As Variant) As Boolean
    IsDialogCancelled = (VarType(response) = vbBoolean And response = False)
End Function

Private Function ParseMonthValue(ByVal rawValue As Variant) As Long
    Dim textValue As String
    Dim monthIndex As Long

    textValue = Trim$(CStr(rawValue))
    If textValue = vbNullString Then Exit Function

    If IsNumeric(textValue) Then
        ParseMonthValue = CLng(textValue)
        Exit Function
    End If

    textValue = UCase$(Left$(textValue, 3))
    For monthIndex = 1 To 12
        If UCase$(MonthName(monthIndex, True)) = textValue Then
            ParseMonthValue = monthIndex
            Exit Function
        End If
    Next monthIndex
End Function

Private Function IsValidWindowsFileName(ByVal fileName As String) As Boolean
    Dim invalidChars As Variant
    Dim item As Variant

    fileName = Trim$(fileName)
    If fileName = vbNullString Then Exit Function
    If Right$(fileName, 1) = "." Then Exit Function

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each item In invalidChars
        If InStr(1, fileName, CStr(item), vbBinaryCompare) > 0 Then Exit Function
    Next item

    IsValidWindowsFileName = True
End Function

Private Function SafeWorkbookName(ByVal ctx As Infra_ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.WorkbookRef Is Nothing Then Exit Function
    SafeWorkbookName = ctx.WorkbookRef.Name
End Function

Private Function SafeWorksheetName(ByVal ctx As Infra_ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.WorksheetRef Is Nothing Then Exit Function
    SafeWorksheetName = ctx.WorksheetRef.Name
End Function

Private Function SafeSelectionAddress(ByVal ctx As Infra_ActionContext) As String
    If ctx Is Nothing Then Exit Function
    If ctx.SelectionRange Is Nothing Then
        SafeSelectionAddress = "(none)"
    Else
        SafeSelectionAddress = ctx.SelectionRange.Address(False, False)
    End If
End Function
