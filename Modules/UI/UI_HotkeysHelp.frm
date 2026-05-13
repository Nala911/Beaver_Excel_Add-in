VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_HotkeysHelp 
   Caption         =   "Keyboard Shortcuts"
   ClientHeight    =   5480
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   6800
   OleObjectBlob   =   "UI_HotkeysHelp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UI_HotkeysHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' @Module: UI_HotkeysHelp
' @Category: UI
' @Description: Displays the configured hotkeys in a standard UserForm.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Hotkeys

Private Sub UserForm_Initialize()
    Dim defs As Variant
    Dim i As Long
    
    ' Fetch hotkeys from Infra_Hotkeys
    defs = Infra_Hotkeys.HotkeyDefinitions()
    
    With Me.lstHotkeys
        .Clear
        .ColumnCount = 2
        
        If IsEmpty(defs) Then
            .AddItem "No hotkeys defined."
            .List(0, 1) = ""
            Exit Sub
        End If
        
        Dim r As Long
        r = 0
        For i = LBound(defs, 1) To UBound(defs, 1)
            If defs(i, 1) <> "" And defs(i, 3) <> "" Then
                .AddItem Infra_Hotkeys.TranslateHotkey(CStr(defs(i, 1)))
                .List(r, 1) = defs(i, 3)
                r = r + 1
            End If
        Next i
    End With
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub
