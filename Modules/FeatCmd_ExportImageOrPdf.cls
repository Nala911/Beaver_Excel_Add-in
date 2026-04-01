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

Private Function ICommand_CanExecute() As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("Export.CanExecute")
    On Error GoTo ErrHandler
    Dim ctx As Infra_ActionContext
    Set ctx = Infra_AppState.CaptureActionContext()
    ICommand_CanExecute = Not ctx Is Nothing
    If ICommand_CanExecute Then
        ICommand_CanExecute = Not ctx.WorksheetRef Is Nothing
    End If
    Exit Function
ErrHandler:
    Infra_Error.HandleError "Export.CanExecute", Err
    ICommand_CanExecute = False
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("Export.Execute")
    On Error GoTo ErrHandler

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
    Infra_Error.HandleError "Export.Execute", Err, stateInfo
    Resume CleanExit
End Sub

Private Sub ExecuteExport(ByVal request As Infra_ExportRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteExport")
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
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ExecuteExport", Err
    Resume CleanExit
End Sub

Private Sub ExportAsPngImpl(ByVal selRng As Range, ByVal scaleFactor As Long)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExportAsPngImpl")
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
               vbCritical, Infra_Config.ADDIN_NAME & " - Export Error"
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
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ExportAsPngImpl", Err
    Resume PngCleanup
End Sub

Private Sub ExportAsPdfImpl(ByVal exportRng As Range)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExportAsPdfImpl")
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
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ExportAsPdfImpl", Err
    Resume CleanExit
End Sub
