VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_PageSetupBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_PageSetupBackup
' @Category: Infrastructure
' @Description: Backs up and restores worksheet PageSetup settings and visibility.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Public SheetRef As Worksheet
Public OriginalPrintArea As String
Public OriginalOrientation As Long
Public OriginalPaperSize As Long
Public OriginalLeftMargin As Double
Public OriginalRightMargin As Double
Public OriginalTopMargin As Double
Public OriginalBottomMargin As Double
Public OriginalHeaderMargin As Double
Public OriginalFooterMargin As Double
Public OriginalZoom As Variant
Public OriginalFitToPagesWide As Variant
Public OriginalFitToPagesTall As Variant
Public OriginalCenterHorizontally As Boolean
Public OriginalCenterVertically As Boolean
Public OriginalVisible As Long

Public Sub Backup(ByVal ws As Worksheet)
    Set SheetRef = ws
    OriginalVisible = ws.Visible
    
    With ws.PageSetup
        OriginalPrintArea = .PrintArea
        OriginalOrientation = .Orientation
        OriginalPaperSize = .PaperSize
        OriginalLeftMargin = .LeftMargin
        OriginalRightMargin = .RightMargin
        OriginalTopMargin = .TopMargin
        OriginalBottomMargin = .BottomMargin
        OriginalHeaderMargin = .HeaderMargin
        OriginalFooterMargin = .FooterMargin
        OriginalZoom = .Zoom
        OriginalFitToPagesWide = .FitToPagesWide
        OriginalFitToPagesTall = .FitToPagesTall
        OriginalCenterHorizontally = .CenterHorizontally
        OriginalCenterVertically = .CenterVertically
    End With
End Sub

Public Sub Restore()
    If SheetRef Is Nothing Then Exit Sub
    
    ' Restore visibility
    SheetRef.Visible = OriginalVisible
    
    ' Restore PageSetup
    With SheetRef.PageSetup
        .PrintArea = OriginalPrintArea
        .Orientation = OriginalOrientation
        .PaperSize = OriginalPaperSize
        .LeftMargin = OriginalLeftMargin
        .RightMargin = OriginalRightMargin
        .TopMargin = OriginalTopMargin
        .BottomMargin = OriginalBottomMargin
        .HeaderMargin = OriginalHeaderMargin
        .FooterMargin = OriginalFooterMargin
        .Zoom = OriginalZoom
        .FitToPagesWide = OriginalFitToPagesWide
        .FitToPagesTall = OriginalFitToPagesTall
        .CenterHorizontally = OriginalCenterHorizontally
        .CenterVertically = OriginalCenterVertically
    End With
End Sub
