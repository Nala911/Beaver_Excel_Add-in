Attribute VB_Name = "UI_DialogExport"
Option Explicit

' @Module: UI_DialogExport
' @Category: UI
' @Description: Dialog options and output file selection prompts for exporting range data (Image or PDF).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Infra_Interaction, Infra_Config, Infra_AppState, ActionContext, ExportRequest, UI_DialogShared

' Shows the Export options via UserForm picker and returns a populated Request object.
Public Function ShowExportDialog(ByVal ctx As ActionContext, Optional ByVal commandName As String = vbNullString) As ExportRequest
    Dim tracker As Object: Set tracker = Infra_Error.Track("ShowExportDialog")
    On Error GoTo ErrHandler
    
    Dim request As ExportRequest
    Dim exportChoice As String
    Dim normalizedChoice As String
    Dim ws As Worksheet
    Dim selRng As Range

    If ctx Is Nothing Then GoTo CleanExit
    Set ws = ctx.WorksheetRef
    If ws Is Nothing Then GoTo CleanExit
    
    Set request = New ExportRequest
    Set request.Context = ctx
    request.ScaleFactor = Infra_Config.DEFAULT_EXPORT_SCALE

    ' Step 1: Prompt Scope (Sheet, Workbook, or Selection)
    Dim hasSelection As Boolean
    hasSelection = False
    
    If TypeName(Selection) = "Range" Then
        Set selRng = Selection
        If selRng.Cells.CountLarge > 1 Or selRng.Areas.Count > 1 Then
            hasSelection = True
        End If
    End If

    Dim scopeChoice As String
    Dim scopeOptions As Variant
    Dim defaultScope As String

    If hasSelection Then
        scopeOptions = UI_DialogShared.BuildChoiceArray("Selection", "Sheet", "Workbook")
        defaultScope = "Selection"
    Else
        scopeOptions = UI_DialogShared.BuildChoiceArray("Sheet", "Workbook")
        defaultScope = "Sheet"
    End If

    If Not Infra_Interaction.PromptOption( _
        "What content would you like to export?", _
        UI_DialogShared.BuildDialogTitle("Export Scope"), defaultScope, scopeOptions, scopeChoice, "ExportScopeOptions") Then GoTo CleanExit

    Dim normalizedScope As String
    normalizedScope = UI_DialogShared.NormalizeChoiceText(scopeChoice)
    If normalizedScope = "SELECTION" Or normalizedScope = "RANGE" Then
        request.ExportScope = "Selection"
        Set request.SourceRange = selRng
    ElseIf normalizedScope = "WORKBOOK" Or normalizedScope = "WB" Then
        request.ExportScope = "Workbook"
        Set request.SourceRange = ws.UsedRange
    Else
        request.ExportScope = "Sheet"
        Set request.SourceRange = ws.UsedRange
    End If

    If request.SourceRange Is Nothing Then
        Set request.SourceRange = ws.UsedRange
    End If

    If request.SourceRange Is Nothing Then
        Infra_Interaction.ShowWarning "No data found on the active sheet to export."
        GoTo CleanExit
    End If

    ' Step 2: Prompt Destination (Copy vs Save to Desktop)
    Dim destChoice As String
    Dim normalizedDest As String

    If Not Infra_Interaction.PromptOption( _
        "Where would you like to export?" & vbCrLf & vbCrLf & _
        BuildExportSummary(request.SourceRange) & vbCrLf & vbCrLf & _
        "Copy - Copy directly to clipboard to paste into Gmail, WhatsApp, etc." & vbCrLf & _
        "Save to Desktop - Export as a file on Desktop", _
        UI_DialogShared.BuildDialogTitle("Export Destination"), "Save to Desktop", UI_DialogShared.BuildChoiceArray("Copy", "Save to Desktop"), destChoice, "ExportDestinationOptions") Then GoTo CleanExit

    normalizedDest = UI_DialogShared.NormalizeChoiceText(destChoice)
    If normalizedDest = "COPY" Or normalizedDest = "CLIPBOARD" Or normalizedDest = "C" Then
        request.CopyToClipboard = True
        request.ExportAsPng = True
        Set ShowExportDialog = request
        GoTo CleanExit
    Else
        request.CopyToClipboard = False
    End If

    ' Step 3: Prompt Format (PNG vs PDF) & Output Path
    Dim skipFormatPrompt As Boolean
    skipFormatPrompt = False
    
    If LCase$(commandName) = "exportpng" Then
        request.ExportAsPng = True
        skipFormatPrompt = True
    ElseIf LCase$(commandName) = "exportpdf" Then
        request.ExportAsPng = False
        skipFormatPrompt = True
    End If

    If Not skipFormatPrompt Then
        Do
            If Not Infra_Interaction.PromptOption( _
                "Choose an export file format:" & vbCrLf & vbCrLf & _
                "PNG - High-resolution image" & vbCrLf & _
                "PDF - Print-ready document", _
                UI_DialogShared.BuildDialogTitle("Export Format"), "PNG", UI_DialogShared.BuildChoiceArray("PNG", "PDF"), exportChoice, "ExportOptions") Then GoTo CleanExit

            normalizedChoice = UI_DialogShared.NormalizeChoiceText(exportChoice)
            Select Case normalizedChoice
                Case "", "PNG", "IMAGE"
                    request.ExportAsPng = True
                    Exit Do
                Case "PDF"
                    request.ExportAsPng = False
                    Exit Do
                Case Else
                    Infra_Interaction.ShowWarning "Please choose PNG or PDF.", UI_DialogShared.BuildDialogTitle("Export Format")
            End Select
        Loop
    End If

    If request.ExportAsPng Then
        request.ScaleFactor = PromptForExportScale(request.ScaleFactor)
        If request.ScaleFactor = 0 Then GoTo CleanExit
    End If

    request.OutputPath = ResolveDesktopOutputPath( _
        ctx, _
        "Export", _
        BuildSuggestedExportBaseName(request.SourceRange, request.ExportAsPng), _
        IIf(request.ExportAsPng, "png", "pdf"))
    If request.OutputPath = vbNullString Then GoTo CleanExit
    
    Set ShowExportDialog = request

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ShowExportDialog", Err
    Resume CleanExit
End Function

' --- PRIVATE HELPERS ---

Private Function ResolveExportRange(ByVal ctx As ActionContext) As Range
    If ctx Is Nothing Then Exit Function
    If ctx.WorksheetRef Is Nothing Then Exit Function
    
    If Not ctx.HasRangeSelection Or ctx.SelectionRange Is Nothing Then
        Set ResolveExportRange = ctx.WorksheetRef.UsedRange
    ElseIf ctx.SelectionRange.Cells.CountLarge = 1 Then
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

    Dim scaleInput As String
    Dim normalizedScale As Long

    Do
        If Not Infra_Interaction.PromptText( _
            "Choose the PNG scale factor." & vbCrLf & vbCrLf & _
            "1 = Smaller file" & vbCrLf & _
            CStr(Infra_Config.DEFAULT_EXPORT_SCALE) & " = Balanced default" & vbCrLf & _
            CStr(Infra_Config.MAX_EXPORT_SCALE) & " = Largest supported image" & vbCrLf & vbCrLf & _
            "Enter a number from 1 to " & Infra_Config.MAX_EXPORT_SCALE & ".", _
            UI_DialogShared.BuildDialogTitle("PNG Quality"), CStr(defaultScale), scaleInput) Then GoTo CleanExit

        If Trim$(CStr(scaleInput)) = vbNullString Then scaleInput = CStr(defaultScale)
        If IsNumeric(scaleInput) Then
            normalizedScale = CLng(scaleInput)
            If normalizedScale >= 1 And normalizedScale <= Infra_Config.MAX_EXPORT_SCALE Then
                PromptForExportScale = normalizedScale
                Exit Do
            End If
        End If

        Infra_Interaction.ShowWarning "Please enter a whole number from 1 to " & Infra_Config.MAX_EXPORT_SCALE & ".", _
                                       UI_DialogShared.BuildDialogTitle("PNG Quality")
    Loop

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "PromptForExportScale", Err
    Resume CleanExit
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

Private Function BuildSuggestedExportBaseName(ByVal sourceRange As Range, ByVal exportAsPng As Boolean) As String
    Dim fileStem As String

    If sourceRange Is Nothing Then
        fileStem = "BeaverExport"
    Else
        fileStem = IIf(exportAsPng, "RangeImage", "RangePDF") & "_" & _
                   Infra_AppState.SanitizeFileNameStem(sourceRange.Worksheet.Name) & "_" & _
                   BuildRangeFileLabel(sourceRange)
    End If

    BuildSuggestedExportBaseName = fileStem & "_" & Format(Now, "yyyymmdd_hhnnss")
End Function

Private Function BuildRangeFileLabel(ByVal sourceRange As Range) As String
    If sourceRange Is Nothing Then
        BuildRangeFileLabel = "Selection"
    Else
        BuildRangeFileLabel = Replace(sourceRange.Address(False, False), ":", "-")
        BuildRangeFileLabel = Replace(BuildRangeFileLabel, "$", "")
    End If
End Function

Private Function ResolveDesktopOutputPath(ByVal ctx As ActionContext, ByVal taskName As String, ByVal suggestedBaseName As String, ByVal extensionWithoutDot As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResolveDesktopOutputPath")
    On Error GoTo ErrHandler

    Dim desktopPath As String
    Dim normalizedBaseName As String

    desktopPath = Infra_AppState.GetDesktopPath()
    If desktopPath = vbNullString Then
        Infra_Interaction.ShowCritical "Could not locate a default save location."
        GoTo CleanExit
    End If

    normalizedBaseName = Infra_AppState.SanitizeFileNameStem(suggestedBaseName)
    If normalizedBaseName = vbNullString Then normalizedBaseName = "BeaverOutput"

    ResolveDesktopOutputPath = Infra_AppState.CombinePath(desktopPath, normalizedBaseName & "." & LCase$(extensionWithoutDot))

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ResolveDesktopOutputPath", Err
    Resume CleanExit
End Function
