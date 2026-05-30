VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_HelpCenter 
   Caption         =   "Help Center"
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
' @Description: Displays hotkeys and UDFs in the Beaver Help Center dialog.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Hotkeys



Private Sub UserForm_Initialize()
    LoadHotkeysList
    ConfigureHotkeysLayout
End Sub



Private Sub LoadHotkeysList()
    Dim defs As Variant
    Dim i As Long
    
    ' Fetch hotkeys from Infra_Hotkeys
    defs = Infra_Hotkeys.HotkeyDefinitions()
    
    With Me.lstHotkeys
        .Clear
        .ColumnCount = 2
        
        Dim r As Long
        r = 0
        
        ' --- 1. Keyboard Shortcuts ---
        .AddItem " Keyboard Shortcuts "
        .List(r, 1) = ""
        r = r + 1
        
        If Not IsEmpty(defs) Then
            For i = LBound(defs, 1) To UBound(defs, 1)
                If defs(i, 1) <> "" And defs(i, 3) <> "" Then
                    .AddItem "  " & Infra_Hotkeys.TranslateHotkey(CStr(defs(i, 1)))
                    .List(r, 1) = defs(i, 3)
                    r = r + 1
                End If
            Next i
        Else
            .AddItem "  No hotkeys defined."
            .List(r, 1) = ""
            r = r + 1
        End If
        
        .AddItem ""
        .List(r, 1) = ""
        r = r + 1
        
        ' --- 2. User Defined Functions ---
        .AddItem " User Defined Functions "
        .List(r, 1) = ""
        r = r + 1
        
        .AddItem "  XFilter(Range_A, Range_B, code)"
        .List(r, 1) = "Advanced set operations (1=Intersect, 2=Diff)"
        r = r + 1
        

        
    End With
End Sub



Private Sub ConfigureHotkeysLayout()
    Const HOTKEYS_FORM_WIDTH As Double = 420
    Const HOTKEYS_MIN_LIST_HEIGHT As Double = 120
    Const HOTKEYS_MAX_LIST_HEIGHT As Double = 350
    Const HOTKEY_ROW_HEIGHT As Double = 18
    Const FORM_BOTTOM_PADDING As Double = 20
    Const BUTTON_GAP As Double = 10
    Const SIDE_MARGIN As Double = 18

    Dim listHeight As Double

    Me.Caption = "Help Center"
    btnOK.Caption = "Close"

    Me.Width = HOTKEYS_FORM_WIDTH

    lstHotkeys.Left = SIDE_MARGIN
    lstHotkeys.Top = 18
    lstHotkeys.Width = Me.InsideWidth - (SIDE_MARGIN * 2)
    lstHotkeys.ColumnWidths = "150 pt;250 pt"

    listHeight = (lstHotkeys.ListCount * HOTKEY_ROW_HEIGHT) + 8
    If listHeight < HOTKEYS_MIN_LIST_HEIGHT Then listHeight = HOTKEYS_MIN_LIST_HEIGHT
    If listHeight > HOTKEYS_MAX_LIST_HEIGHT Then listHeight = HOTKEYS_MAX_LIST_HEIGHT
    lstHotkeys.Height = listHeight

    btnOK.Top = lstHotkeys.Top + lstHotkeys.Height + BUTTON_GAP
    btnOK.Left = Me.InsideWidth - btnOK.Width - SIDE_MARGIN
    SetFormInsideHeight btnOK.Top + btnOK.Height + FORM_BOTTOM_PADDING
End Sub



Private Sub SetFormInsideHeight(ByVal desiredInsideHeight As Double, Optional ByVal minimumOverallHeight As Double = 0)
    Dim frameHeight As Double
    Dim targetHeight As Double

    frameHeight = Me.Height - Me.InsideHeight
    If frameHeight < 0 Then frameHeight = 0

    targetHeight = desiredInsideHeight + frameHeight
    If targetHeight < minimumOverallHeight Then targetHeight = minimumOverallHeight

    Me.Height = targetHeight
End Sub



Private Sub btnOK_Click()
    Unload Me
End Sub


