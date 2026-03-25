Attribute VB_Name = "Feat_ExportImageOrPdf"
Option Explicit

' @Module: Feat_ExportImageOrPdf
' @Category: Feature
' @Description: Exports selected range as a PNG image or high-quality PDF.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_AppStateGuard, Infra_Config, Infra_Error, Infra_ExportRequest

' Main entry point: resolves selection, prompts PNG vs PDF via Factory, delegates to helpers.
Public Sub Export()
    Dim tracker As Object: Set tracker = Infra_Error.Track("Export")
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim ctx As Infra_ActionContext
    Dim request As Infra_ExportRequest
    Dim stateInfo As Collection
    Set stateInfo = Infra_Error.NewStateInfo()
    
    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Then GoTo CleanExit
    Infra_Error.AddState stateInfo, "Worksheet", ctx.WorksheetRef.Name
    
    ' IMPROVED WORKFLOW: Use UI Factory for decentralized request building.
    Set request = Infra_UIFactory.ShowExportDialog(ctx)
    If request Is Nothing Then GoTo CleanExit
    
    ExecuteExport request

CleanExit:
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    HandleError "Export", Err, stateInfo
    Resume CleanExit
End Sub

' BuildExportRequest and related helpers removed in favor of Infra_UIFactory.ShowExportDialog

Private Sub ExecuteExport(ByVal request As Infra_ExportRequest)
    PushContext "ExecuteExport"
    On Error GoTo ErrHandler
    
    If request Is Nothing Then GoTo CleanExit
    If request.SourceRange Is Nothing Then GoTo CleanExit
    
    Application.ScreenUpdating = False
    
    If request.ExportAsPng Then
        Application.StatusBar = "Beaver: Exporting image - please wait..."
        ExportAsPngImpl request.SourceRange, request.ScaleFactor
    Else
        Application.StatusBar = "Beaver: Exporting PDF - please wait..."
        ExportAsPdfImpl request.SourceRange
    End If

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "ExecuteExport", Err
End Sub

' Copies the range as a picture, pastes into a temporary chart on a temp sheet,
' scales the chart by scaleFactor, and exports a PNG to the Desktop.
Private Sub ExportAsPngImpl(ByVal selRng As Range, ByVal scaleFactor As Long)
    PushContext "ExportAsPngImpl"
    On Error GoTo ErrHandler
    
    Dim srcWS As Worksheet
    Dim tmpWS As Worksheet
    Dim chtObj As ChartObject
    Dim filePath As String
    Dim userWB As Workbook

    If selRng Is Nothing Then GoTo CleanExit

    Set srcWS = selRng.Worksheet
    filePath = Infra_AppState.GetDesktopPath() & "\RangeImage_" & srcWS.Name & _
               "_Scale" & scaleFactor & "x_" & Format(Now, "yyyymmdd_hhnnss") & ".png"

    ' Copy range as picture
    On Error Resume Next
    selRng.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
    If Err.Number <> 0 Then
        MsgBox "Failed to copy range as picture.", vbCritical, Infra_Config.ADDIN_NAME
        On Error GoTo ErrHandler
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler

    ' Temp sheet on the user's workbook
    Set userWB = selRng.Worksheet.Parent
    Set tmpWS = userWB.Worksheets.Add(After:=userWB.Worksheets(userWB.Worksheets.Count))
    tmpWS.Name = "tmpExport_" & Format(Now, "hhmmss")

    ' Temp chart
    Set chtObj = tmpWS.ChartObjects.Add(10, 10, Application.CentimetersToPoints(2), Application.CentimetersToPoints(2))

    ' Paste and resize
    On Error Resume Next
    chtObj.Activate
    chtObj.Chart.Paste
    If Err.Number <> 0 Then
        MsgBox "Failed to paste picture into chart.", vbCritical, Infra_Config.ADDIN_NAME
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
        MsgBox "Export failed. Try a lower scale factor or check Desktop permissions.", _
               vbCritical, Infra_Config.ADDIN_NAME & " â€” Export Error"
    End If

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
    PopContext
    Exit Sub

ErrHandler:
    HandleError "ExportAsPngImpl", Err
End Sub

' Exports the given range as a PDF to the Desktop.
' Hardcoded to A4 landscape, fit-all-columns-to-one-page-wide, narrow margins.
Private Sub ExportAsPdfImpl(ByVal exportRng As Range)
    PushContext "ExportAsPdfImpl"
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim filePath As String
    Dim oldPrintArea As String
    Dim oldOrientation As XlPageOrientation
    Dim oldPaperSize As XlPaperSize
    Dim oldLeftMargin As Double
    Dim oldRightMargin As Double
    Dim oldTopMargin As Double
    Dim oldBottomMargin As Double
    Dim oldHeaderMargin As Double
    Dim oldFooterMargin As Double
    Dim oldZoom As Variant
    Dim oldFitToPagesWide As Variant
    Dim oldFitToPagesTall As Variant
    Dim oldCenterHorizontally As Boolean
    Dim oldCenterVertically As Boolean

    Set ws = exportRng.Worksheet
    With ws.PageSetup
        oldPrintArea = .PrintArea
        oldOrientation = .Orientation
        oldPaperSize = .PaperSize
        oldLeftMargin = .LeftMargin
        oldRightMargin = .RightMargin
        oldTopMargin = .TopMargin
        oldBottomMargin = .BottomMargin
        oldHeaderMargin = .HeaderMargin
        oldFooterMargin = .FooterMargin
        oldZoom = .Zoom
        oldFitToPagesWide = .FitToPagesWide
        oldFitToPagesTall = .FitToPagesTall
        oldCenterHorizontally = .CenterHorizontally
        oldCenterVertically = .CenterVertically
    End With

    filePath = Infra_AppState.GetDesktopPath() & "\RangePDF_" & ws.Name & "_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"

    With ws.PageSetup
        .PrintArea = exportRng.Address
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

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

CleanExit:
    On Error Resume Next
    If Not ws Is Nothing Then
        With ws.PageSetup
            .PrintArea = oldPrintArea
            .Orientation = oldOrientation
            .PaperSize = oldPaperSize
            .LeftMargin = oldLeftMargin
            .RightMargin = oldRightMargin
            .TopMargin = oldTopMargin
            .BottomMargin = oldBottomMargin
            .HeaderMargin = oldHeaderMargin
            .FooterMargin = oldFooterMargin
            .Zoom = oldZoom
            .FitToPagesWide = oldFitToPagesWide
            .FitToPagesTall = oldFitToPagesTall
            .CenterHorizontally = oldCenterHorizontally
            .CenterVertically = oldCenterVertically
        End With
    End If
    On Error GoTo 0
    PopContext
    Exit Sub

ErrHandler:
    HandleError "ExportAsPdfImpl", Err
End Sub
