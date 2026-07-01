VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ExportImageOrPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_ExportImageOrPdf
' @Category: Feature
' @Description: Exports selected range as a PNG image or high-quality PDF.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_PageSetupBackup

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorksheet(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Export.Execute")
    On Error GoTo ErrHandler

    Dim ctx As ActionContext
    Dim stateInfo As Collection
    Set stateInfo = Infra_Error.NewStateInfo()
    
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Then GoTo CleanExit
    Infra_Error.AddState stateInfo, "Worksheet", ctx.WorksheetRef.Name

    Dim isPdf As Boolean
    isPdf = (LCase$(context.CommandName) = "exportpdf")

    If isPdf Then
        Dim exportMode As String
        Dim selRng As Range
        Dim ws As Worksheet
        Set ws = ctx.WorksheetRef

        ' 1. Check if we have an active range selection (> 1 cell or multi-area)
        Dim hasRangeSelection As Boolean
        hasRangeSelection = False
        If TypeName(Selection) = "Range" Then
            Set selRng = Selection
            If selRng.Cells.CountLarge > 1 Or selRng.Areas.Count > 1 Then
                hasRangeSelection = True
            End If
        End If

        If hasRangeSelection Then
            exportMode = "Range"
        Else
            ' Prompt options: Sheet or Workbook
            Dim choice As String
            If Not Infra_Interaction.PromptOption( _
                "What would you like to export as PDF?", _
                "Export PDF", "Sheet", Array("Sheet", "Workbook"), choice, _
                "ExportPdfOptions") Then
                GoTo CleanExit
            End If
            exportMode = choice
        End If

        ' 2. Determine Output File Name & Path (Save directly to Desktop)
        Dim desktopPath As String
        desktopPath = Infra_AppState.GetDesktopPath()
        If desktopPath = vbNullString Then
            Infra_Interaction.ShowCritical "Could not locate the Desktop path.", "Export PDF"
            GoTo CleanExit
        End If

        Dim wb As Workbook
        Set wb = ws.Parent

        Dim wbName As String
        If InStr(wb.Name, ".") > 0 Then
            wbName = Left$(wb.Name, InStrRev(wb.Name, ".") - 1)
        Else
            wbName = wb.Name
        End If
        wbName = Infra_AppState.SanitizeFileNameStem(wbName)
        If wbName = vbNullString Then wbName = "Workbook"

        Dim fileLabel As String
        If exportMode = "Range" Then
            fileLabel = Infra_AppState.SanitizeFileNameStem(ws.Name) & "_Range"
        ElseIf exportMode = "Sheet" Then
            fileLabel = Infra_AppState.SanitizeFileNameStem(ws.Name) & "_Sheet"
        Else
            fileLabel = "Workbook"
        End If

        Dim fileName As String
        fileName = wbName & "_" & fileLabel & "_Export_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"

        Dim outputPath As String
        outputPath = Infra_AppState.CombinePath(desktopPath, fileName)

        ' 3. Perform the PDF Export based on Mode
        Application.StatusBar = "Beaver: Exporting PDF - please wait..."

        Dim success As Boolean
        success = False

        If exportMode = "Range" Then
            success = ExportRangesPdf(ws, selRng, outputPath)
        ElseIf exportMode = "Sheet" Then
            success = ExportSheetPdf(ws, outputPath)
        ElseIf exportMode = "Workbook" Then
            success = ExportWorkbookPdf(wb, outputPath)
        End If

        If success Then
            Infra_Interaction.ShowInfo "PDF exported successfully to Desktop:" & vbCrLf & outputPath, Infra_Config.ADDIN_NAME & " - Export PDF"
        End If
    Else
        ' PNG / dialog flow
        Dim request As ExportRequest
        Set request = UI_Factory.ShowExportDialog(ctx, context.CommandName)
        If request Is Nothing Then GoTo CleanExit
        
        If request.SourceRange Is Nothing Then GoTo CleanExit
        If request.OutputPath = vbNullString Then GoTo CleanExit

        If request.ExportAsPng Then
            Application.StatusBar = "Beaver: Exporting image - please wait..."
            If Not ExportAsPngImpl(request.SourceRange, request.ScaleFactor, request.OutputPath) Then GoTo CleanExit
            Infra_Interaction.ShowInfo "Export saved:" & vbCrLf & request.OutputPath, Infra_Config.ADDIN_NAME & " - Export"
        Else
            ' Fallback PDF export
            Application.StatusBar = "Beaver: Exporting PDF - please wait..."
            If Not ExportAsPdfImpl(request.SourceRange, request.OutputPath) Then GoTo CleanExit
            Infra_Interaction.ShowInfo "Export saved:" & vbCrLf & request.OutputPath, Infra_Config.ADDIN_NAME & " - Export"
        End If
    End If

CleanExit:
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Infra_Error.HandleError "Export.Execute", Err, stateInfo
    Resume CleanExit
End Sub

Private Function ExportAsPngImpl(ByVal selRng As Range, ByVal scaleFactor As Long, ByVal filePath As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExportAsPngImpl")
    On Error GoTo ErrHandler
    
    Dim srcWS As Worksheet
    Dim tmpWS As Worksheet
    Dim chtObj As ChartObject
    Dim userWB As Workbook

    If selRng Is Nothing Then GoTo CleanExit

    Set srcWS = selRng.Worksheet
    ' Copy range as picture
    On Error Resume Next
    selRng.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
    If Err.Number <> 0 Then
        Err.Clear
        Infra_Interaction.ShowCritical "Failed to copy range as picture."
        On Error GoTo ErrHandler
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler

    ' Temp sheet on the add-in workbook (ThisWorkbook) to preserve user's Undo stack
    Set tmpWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    tmpWS.Name = "tmpExport_" & Format(Now, "hhmmss")

    ' Temp chart
    Set chtObj = tmpWS.ChartObjects.Add(10, 10, Application.CentimetersToPoints(2), Application.CentimetersToPoints(2))

    ' Paste and resize
    On Error Resume Next
    chtObj.Activate
    chtObj.Chart.Paste
    If Err.Number <> 0 Then
        Err.Clear
        Infra_Interaction.ShowCritical "Failed to paste picture into chart."
        On Error GoTo ErrHandler
        GoTo PngCleanup
    End If
    On Error GoTo ErrHandler

    With chtObj
        .Width = selRng.Width * scaleFactor
        .Height = selRng.Height * scaleFactor
    End With

    ' Wait for chart to render (polling loop instead of hardcoded 1s wait)
    Dim i As Long
    For i = 1 To 20
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 0) + (1 / 86400) / 20 ' ~50ms
    Next i

    If Not chtObj.Chart.Export(filePath) Then
        Infra_Interaction.ShowCritical "Export failed. Try a lower scale factor or check Desktop permissions.", _
                                       Infra_Config.ADDIN_NAME & " - Export Error"
        GoTo PngCleanup
    End If

    ExportAsPngImpl = True

PngCleanup:
CleanExit:
    On Error Resume Next
    If Not chtObj Is Nothing Then chtObj.Delete
    If Not tmpWS Is Nothing Then
        Application.DisplayAlerts = False
        tmpWS.Delete
        Application.DisplayAlerts = True
    End If
    Application.CutCopyMode = False
    On Error GoTo 0
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ExportAsPngImpl", Err
    Resume PngCleanup
End Function

Private Sub ApplyStandardPageSetup(ByVal ws As Worksheet, ByVal printAreaAddress As String)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ApplyStandardPageSetup")
    On Error GoTo ErrHandler

    With ws.PageSetup
        .PrintArea = printAreaAddress
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .CenterVertically = False
    End With

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "ApplyStandardPageSetup", Err
    Resume CleanExit
End Sub

Private Function ExportAsPdfImpl(ByVal exportRng As Range, ByVal filePath As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExportAsPdfImpl")
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = exportRng.Worksheet

    Dim backupObj As New Infra_PageSetupBackup
    backupObj.Backup ws

    ApplyStandardPageSetup ws, exportRng.Address

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ExportAsPdfImpl = True

CleanExit:
    On Error Resume Next
    backupObj.Restore
    On Error GoTo 0
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ExportAsPdfImpl", Err
    Resume CleanExit
End Function

Private Function ExportRangesPdf(ByVal ws As Worksheet, ByVal selRng As Range, ByVal outputPath As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExportRangesPdf")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler
    
    Dim backupObj As New Infra_PageSetupBackup
    backupObj.Backup ws
    
    ' Intersect selected areas with the used range to prevent massive blank grids
    Dim area As Range
    Dim intersectRange As Range
    Dim printAreaAddress As String
    printAreaAddress = ""
    
    For Each area In selRng.Areas
        Set intersectRange = Intersect(area, ws.UsedRange)
        If Not intersectRange Is Nothing Then
            If printAreaAddress <> "" Then
                printAreaAddress = printAreaAddress & ","
            End If
            printAreaAddress = printAreaAddress & intersectRange.Address
        End If
    Next area
    
    ' Fallback to UsedRange if no areas intersect
    If printAreaAddress = "" Then
        printAreaAddress = ws.UsedRange.Address
    End If
    
    ' Apply standard PDF settings
    ApplyStandardPageSetup ws, printAreaAddress
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    ExportRangesPdf = True

CleanExit:
    On Error Resume Next
    backupObj.Restore
    On Error GoTo 0
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ExportRangesPdf", Err
    Resume CleanExit
End Function

Private Function ExportSheetPdf(ByVal ws As Worksheet, ByVal outputPath As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExportSheetPdf")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler
    
    Dim backupObj As New Infra_PageSetupBackup
    backupObj.Backup ws
    
    ' Apply standard PDF settings
    ApplyStandardPageSetup ws, ws.UsedRange.Address
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    ExportSheetPdf = True

CleanExit:
    On Error Resume Next
    backupObj.Restore
    On Error GoTo 0
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ExportSheetPdf", Err
    Resume CleanExit
End Function

Private Function ExportWorkbookPdf(ByVal wb As Workbook, ByVal outputPath As String) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExportWorkbookPdf")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler
    
    Dim backups As New Collection
    Dim ws As Worksheet
    Dim backupObj As Infra_PageSetupBackup
    
    ' First, verify there is at least one visible worksheet with data
    Dim visibleWithDataCount As Long
    visibleWithDataCount = 0
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Not ws.UsedRange Is Nothing Then
                If Application.WorksheetFunction.CountA(ws.UsedRange) > 0 Then
                    visibleWithDataCount = visibleWithDataCount + 1
                End If
            End If
        End If
    Next ws
    
    If visibleWithDataCount = 0 Then
        Infra_Interaction.ShowWarning "No visible worksheets with data were found in the active workbook.", "Export PDF"
        ExportWorkbookPdf = False
        GoTo CleanExit
    End If
    
    ' Proceed with setting up layout and visibility
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Dim hasData As Boolean
            hasData = False
            If Not ws.UsedRange Is Nothing Then
                If Application.WorksheetFunction.CountA(ws.UsedRange) > 0 Then
                    hasData = True
                End If
            End If
            
            Set backupObj = New Infra_PageSetupBackup
            backupObj.Backup ws
            backups.Add backupObj
            
            If hasData Then
                ' Set up print area and layout on visible sheets with data
                ApplyStandardPageSetup ws, ws.UsedRange.Address
            Else
                ' Temporarily hide empty sheets so they aren't printed
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    
    ' Export the workbook
    wb.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    ExportWorkbookPdf = True

CleanExit:
    On Error Resume Next
    ' Restore all sheet settings and visibilities in reverse order
    If backups.Count > 0 Then
        Dim i As Long
        For i = backups.Count To 1 Step -1
            Set backupObj = backups.Item(i)
            backupObj.Restore
        Next i
    End If
    On Error GoTo 0
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ExportWorkbookPdf", Err
    Resume CleanExit
End Function
