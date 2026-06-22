VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_HelpCenter 
   Caption         =   "Beaver Help Center"
   ClientHeight    =   5480
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   6800
   OleObjectBlob   =   "UI_HelpCenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UI_HelpCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: UI_HelpCenter
' @Category: UI
' @Description: Custom scrollable Help Center dialog displaying Beaver features, hotkeys, and diagnostics.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Hotkeys, Lib_UdfRegistry, Infra_Config, Infra_Error

Public Sub DisplayHelp()
    Dim tracker As Object: Set tracker = Infra_Error.Track("DisplayHelp")
    On Error GoTo ErrHandler
    
    If Application.Visible And Not Lib_Tests.IsRunning Then
        Me.Show vbModal
    Else
        Debug.Print "  [SKIP] UI_HelpCenter.DisplayHelp vbModal display bypassed in headless/background environment or unit testing"
    End If
    
CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "DisplayHelp", Err
    Resume CleanExit
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    
    ' Hide standard designer controls from the copied option picker template
    Dim ctrl As Object
    For Each ctrl In Me.Controls
        If ctrl.Name <> "btnOK" Then
            ctrl.Visible = False
        End If
    Next ctrl
    
    ' Set up UserForm dimensions
    Me.Width = 460
    Me.Height = 490
    
    ' 1. Create Title Banner
    Dim pnlBanner As Object
    Set pnlBanner = Me.Controls.Add("Forms.Frame.1", "pnlBanner")
    pnlBanner.Left = 0
    pnlBanner.Top = 0
    pnlBanner.Width = Me.InsideWidth
    pnlBanner.Height = 50
    pnlBanner.BackColor = RGB(45, 48, 58)
    pnlBanner.BorderStyle = 0
    pnlBanner.SpecialEffect = 0
    
    Dim lblTitle As Object
    Set lblTitle = pnlBanner.Controls.Add("Forms.Label.1", "lblTitle")
    lblTitle.Caption = "BEAVER HELP CENTER"
    lblTitle.Left = 15
    lblTitle.Top = 8
    lblTitle.Width = pnlBanner.Width - 30
    lblTitle.Height = 18
    lblTitle.ForeColor = RGB(255, 255, 255)
    lblTitle.Font.Name = "Segoe UI"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 13
    
    Dim lblSubtitle As Object
    Set lblSubtitle = pnlBanner.Controls.Add("Forms.Label.1", "lblSubtitle")
    lblSubtitle.Caption = "Quick reference guide for keyboard shortcuts, functions, and support."
    lblSubtitle.Left = 15
    lblSubtitle.Top = 27
    lblSubtitle.Width = pnlBanner.Width - 30
    lblSubtitle.Height = 15
    lblSubtitle.ForeColor = RGB(190, 195, 205)
    lblSubtitle.Font.Name = "Segoe UI"
    lblSubtitle.Font.Size = 9
    
    ' 2. Create Scrollable Frame
    Dim fraScroll As Object
    Set fraScroll = Me.Controls.Add("Forms.Frame.1", "fraScroll")
    fraScroll.Left = 10
    fraScroll.Top = 60
    fraScroll.Width = Me.InsideWidth - 20
    fraScroll.Height = 345
    fraScroll.BackColor = RGB(250, 250, 252)
    fraScroll.BorderStyle = 0
    fraScroll.SpecialEffect = 0
    fraScroll.ScrollBars = 2 ' fmScrollBarsVertical
    fraScroll.KeepScrollBarsVisible = 2 ' fmScrollBarsVertical
    
    ' Populate scrollable area
    Dim currentTop As Double
    currentTop = 10
    
    ' SECTION 1: Keyboard Shortcuts
    AddSectionHeader "Keyboard Shortcuts", fraScroll, currentTop
    
    Dim defs As Variant
    defs = Infra_Hotkeys.HotkeyDefinitions()
    If Not IsEmpty(defs) Then
        Dim i As Long
        For i = LBound(defs, 1) To UBound(defs, 1)
            If defs(i, 1) <> "" And defs(i, 3) <> "" Then
                AddShortcutRow CStr(defs(i, 1)), CStr(defs(i, 3)), fraScroll, currentTop
            End If
        Next i
    Else
        Dim lblNoKeys As Object
        Set lblNoKeys = fraScroll.Controls.Add("Forms.Label.1")
        lblNoKeys.Caption = "No keyboard shortcuts defined."
        lblNoKeys.Left = 15
        lblNoKeys.Top = currentTop
        lblNoKeys.Width = fraScroll.Width - 30
        lblNoKeys.Height = 15
        lblNoKeys.Font.Name = "Segoe UI"
        lblNoKeys.Font.Italic = True
        lblNoKeys.ForeColor = RGB(120, 125, 135)
        currentTop = currentTop + 20
    End If
    
    currentTop = currentTop + 10
    
    ' SECTION 2: User Defined Functions (UDFs)
    AddSectionHeader "User Defined Functions (UDFs)", fraScroll, currentTop
    
    Dim udfs As Collection
    Set udfs = Lib_UdfRegistry.GetAllUdfs()
    If Not udfs Is Nothing And udfs.Count > 0 Then
        Dim udfMeta As Object
        For Each udfMeta In udfs
            AddUdfRow udfMeta, fraScroll, currentTop
        Next udfMeta
    Else
        Dim lblNoUdfs As Object
        Set lblNoUdfs = fraScroll.Controls.Add("Forms.Label.1")
        lblNoUdfs.Caption = "No user defined functions registered."
        lblNoUdfs.Left = 15
        lblNoUdfs.Top = currentTop
        lblNoUdfs.Width = fraScroll.Width - 30
        lblNoUdfs.Height = 15
        lblNoUdfs.Font.Name = "Segoe UI"
        lblNoUdfs.Font.Italic = True
        lblNoUdfs.ForeColor = RGB(120, 125, 135)
        currentTop = currentTop + 20
    End If
    
    currentTop = currentTop + 10
    
    ' SECTION 3: Ribbon Features
    AddSectionHeader "Ribbon Features", fraScroll, currentTop
    
    Dim features As Collection
    On Error Resume Next
    Set features = Lib_HelpManifest.GetFeatureHelp()
    On Error GoTo ErrHandler
    
    If Not features Is Nothing And features.Count > 0 Then
        Dim featDict As Object
        For Each featDict In features
            Dim fLabel As String: fLabel = ""
            Dim fScreentip As String: fScreentip = ""
            Dim fSupertip As String: fSupertip = ""
            
            On Error Resume Next
            fLabel = featDict("Label")
            fScreentip = featDict("Screentip")
            fSupertip = featDict("Supertip")
            On Error GoTo ErrHandler
            
            If fLabel <> "" Then
                AddFeatureRow fLabel, fScreentip, fSupertip, fraScroll, currentTop
            End If
        Next featDict
    Else
        Dim lblNoFeatures As Object
        Set lblNoFeatures = fraScroll.Controls.Add("Forms.Label.1")
        lblNoFeatures.Caption = "No ribbon features registered."
        lblNoFeatures.Left = 15
        lblNoFeatures.Top = currentTop
        lblNoFeatures.Width = fraScroll.Width - 30
        lblNoFeatures.Height = 15
        lblNoFeatures.Font.Name = "Segoe UI"
        lblNoFeatures.Font.Italic = True
        lblNoFeatures.ForeColor = RGB(120, 125, 135)
        currentTop = currentTop + 20
    End If
    
    ' Set the scrolling height to fit all content + bottom padding
    fraScroll.ScrollHeight = currentTop + 10
    
    ' 3. Configure OK / Close Button
    btnOK.Caption = "Close"
    btnOK.Width = 80
    btnOK.Height = 24
    btnOK.Left = Me.InsideWidth - 90
    btnOK.Top = 415
    btnOK.Font.Name = "Segoe UI"
    btnOK.Font.Size = 10
    btnOK.Default = True
    btnOK.Cancel = True
    btnOK.Visible = True

CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error initializing Help Center: " & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub AddSectionHeader(ByVal title As String, ByVal parent As Object, ByRef currentTop As Double)
    ' Title label
    Dim lbl As Object
    Set lbl = parent.Controls.Add("Forms.Label.1")
    lbl.Caption = title
    lbl.Left = 10
    lbl.Top = currentTop
    lbl.Width = parent.Width - 40
    lbl.Height = 18
    lbl.Font.Name = "Segoe UI"
    lbl.Font.Bold = True
    lbl.Font.Size = 11
    lbl.ForeColor = RGB(10, 37, 64)
    
    currentTop = currentTop + 18
    
    ' Line separator
    Dim line As Object
    Set line = parent.Controls.Add("Forms.Label.1")
    line.Caption = ""
    line.Left = 10
    line.Top = currentTop
    line.Width = parent.Width - 40
    line.Height = 1
    line.BackColor = RGB(215, 220, 228)
    
    currentTop = currentTop + 8
End Sub

Private Sub AddShortcutRow(ByVal keyPattern As String, ByVal description As String, ByVal parent As Object, ByRef currentTop As Double)
    Dim keyText As String
    keyText = Infra_Hotkeys.TranslateHotkey(keyPattern)
    
    ' Keyboard badge label
    Dim badge As Object
    Set badge = parent.Controls.Add("Forms.Label.1")
    badge.Caption = keyText
    badge.Left = 15
    badge.Top = currentTop
    badge.Width = 120
    badge.Height = 16
    badge.BackColor = RGB(242, 244, 247)
    badge.ForeColor = RGB(50, 55, 65)
    badge.BorderColor = RGB(205, 210, 218)
    badge.BorderStyle = 1
    badge.TextAlign = 2
    badge.Font.Name = "Segoe UI"
    badge.Font.Bold = True
    badge.Font.Size = 8.5
    
    ' Description label
    Dim lblDesc As Object
    Set lblDesc = parent.Controls.Add("Forms.Label.1")
    lblDesc.Caption = description
    lblDesc.Left = 145
    lblDesc.Top = currentTop + 1
    lblDesc.Width = parent.Width - 170
    lblDesc.Height = 15
    lblDesc.Font.Name = "Segoe UI"
    lblDesc.Font.Size = 9.5
    lblDesc.ForeColor = RGB(70, 75, 85)
    
    currentTop = currentTop + 22
End Sub

Private Sub AddUdfRow(ByVal udfMeta As Object, ByVal parent As Object, ByRef currentTop As Double)
    ' Syntax
    Dim lblSyntax As Object
    Set lblSyntax = parent.Controls.Add("Forms.Label.1")
    lblSyntax.Caption = udfMeta("Syntax")
    lblSyntax.Left = 15
    lblSyntax.Top = currentTop
    lblSyntax.Width = parent.Width - 45
    lblSyntax.Height = 15
    lblSyntax.Font.Name = "Consolas"
    lblSyntax.Font.Bold = True
    lblSyntax.Font.Size = 9
    lblSyntax.ForeColor = RGB(0, 102, 204)
    
    currentTop = currentTop + 16
    
    ' Description
    Dim lblDesc As Object
    Set lblDesc = parent.Controls.Add("Forms.Label.1")
    lblDesc.Caption = udfMeta("Description")
    lblDesc.Left = 25
    lblDesc.Top = currentTop
    lblDesc.Width = parent.Width - 55
    lblDesc.Height = 26
    lblDesc.WordWrap = True
    lblDesc.Font.Name = "Segoe UI"
    lblDesc.Font.Size = 9.5
    lblDesc.ForeColor = RGB(70, 75, 85)
    
    currentTop = currentTop + 28
    
    ' Arguments
    Dim argDesc As Variant
    argDesc = udfMeta("ArgumentDescriptions")
    
    If IsArray(argDesc) Then
        Dim i As Long
        For i = LBound(argDesc) To UBound(argDesc)
            Dim lblBullet As Object
            Set lblBullet = parent.Controls.Add("Forms.Label.1")
            lblBullet.Caption = ChrW(8226)
            lblBullet.Left = 32
            lblBullet.Top = currentTop
            lblBullet.Width = 10
            lblBullet.Height = 12
            lblBullet.Font.Name = "Segoe UI"
            lblBullet.Font.Bold = True
            lblBullet.Font.Size = 9
            lblBullet.ForeColor = RGB(120, 125, 135)
            
            Dim lblArg As Object
            Set lblArg = parent.Controls.Add("Forms.Label.1")
            lblArg.Caption = CStr(argDesc(i))
            lblArg.Left = 45
            lblArg.Top = currentTop
            lblArg.Width = parent.Width - 75
            lblArg.WordWrap = True
            lblArg.Font.Name = "Segoe UI"
            lblArg.Font.Size = 9
            lblArg.ForeColor = RGB(85, 90, 100)
            
            Dim approxLines As Long
            approxLines = (Len(lblArg.Caption) \ 60) + 1
            lblArg.Height = approxLines * 12.5
            
            currentTop = currentTop + lblArg.Height + 2
        Next i
    End If
    
    currentTop = currentTop + 6
End Sub

Private Sub AddFeatureRow(ByVal label As String, ByVal screentip As String, ByVal supertip As String, ByVal parent As Object, ByRef currentTop As Double)
    ' Feature Label
    Dim lblLabel As Object
    Set lblLabel = parent.Controls.Add("Forms.Label.1")
    lblLabel.Caption = label
    lblLabel.Left = 15
    lblLabel.Top = currentTop
    lblLabel.Width = 120
    lblLabel.Height = 15
    lblLabel.Font.Name = "Segoe UI"
    lblLabel.Font.Bold = True
    lblLabel.Font.Size = 9.5
    lblLabel.ForeColor = RGB(10, 37, 64)
    
    ' Screentip (Brief description)
    Dim lblScreentip As Object
    Set lblScreentip = parent.Controls.Add("Forms.Label.1")
    lblScreentip.Caption = screentip
    lblScreentip.Left = 145
    lblScreentip.Top = currentTop
    lblScreentip.Width = parent.Width - 170
    lblScreentip.Height = 15
    lblScreentip.Font.Name = "Segoe UI"
    lblScreentip.Font.Size = 9.5
    lblScreentip.ForeColor = RGB(50, 55, 65)
    
    currentTop = currentTop + 16
    
    ' Supertip (Extended details) if present
    If supertip <> "" Then
        Dim lblSupertip As Object
        Set lblSupertip = parent.Controls.Add("Forms.Label.1")
        lblSupertip.Caption = supertip
        lblSupertip.Left = 145
        lblSupertip.Top = currentTop
        lblSupertip.Width = parent.Width - 170
        lblSupertip.WordWrap = True
        lblSupertip.Font.Name = "Segoe UI"
        lblSupertip.Font.Size = 8.5
        lblSupertip.Font.Italic = True
        lblSupertip.ForeColor = RGB(120, 125, 135)
        
        Dim approxLines As Long
        approxLines = (Len(supertip) \ 60) + 1
        lblSupertip.Height = approxLines * 12
        currentTop = currentTop + lblSupertip.Height + 4
    Else
        currentTop = currentTop + 4
    End If
End Sub


