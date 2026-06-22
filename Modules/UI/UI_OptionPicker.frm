VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_OptionPicker 
   Caption         =   "Choose Option"
   ClientHeight    =   5480
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   6800
   OleObjectBlob   =   "UI_OptionPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UI_OptionPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' @Module: UI_OptionPicker
' @Category: UI
' @Description: Dedicated option-picker form reused by interactive command dialogs.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error


Private mConfirmed As Boolean
Private mSelectedValue As String
Private mIgnoreClick As Boolean
Private mIsMultiSelectCheckList As Boolean

Public Property Get IsIgnoringClick() As Boolean
    IsIgnoringClick = mIgnoreClick
End Property

Private Function CheckedPrefix() As String
    CheckedPrefix = ChrW$(9745) & "  "
End Function

Private Function UncheckedPrefix() As String
    UncheckedPrefix = ChrW$(9744) & "  "
End Function

Private Function PrefixLen() As Long
    PrefixLen = 3
End Function

Private Sub UserForm_Initialize()
    mConfirmed = False
    mSelectedValue = vbNullString
    mIgnoreClick = False
    mIsMultiSelectCheckList = False
    
    ' Hide controls so they do not show in the option picker window by default
    On Error Resume Next
    lblPrompt.Visible = False
    btnOK.Visible = False
    btnCancel.Visible = False
    On Error GoTo 0
End Sub

Public Sub ConfigureOptionPicker(ByVal dialogTitle As String, ByVal promptText As String, ByVal defaultChoice As String, ByVal options As Variant)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConfigureOptionPicker")
    On Error GoTo ErrHandler

    mConfirmed = False
    mSelectedValue = vbNullString
    mIgnoreClick = False
    mIsMultiSelectCheckList = False

    Me.Caption = dialogTitle
    
    ' Reset list style to default single-select
    Me.lstHotkeys.MultiSelect = 0 ' fmMultiSelectSingle
    Me.lstHotkeys.ListStyle = 0 ' fmListStylePlain
    
    On Error Resume Next
    Me.lstHotkeys.Font.Name = "Segoe UI"
    Me.lstHotkeys.Font.Size = 10
    On Error GoTo ErrHandler
    
    mIgnoreClick = True
    LoadOptionList defaultChoice, options
    mIgnoreClick = False
    
    ResizeOptionPickerLayout

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "ConfigureOptionPicker", Err
    Resume CleanExit
End Sub

Public Sub ConfigureMultiOptionPicker(ByVal dialogTitle As String, ByVal promptText As String, ByVal options As Variant, ByVal defaultChecked As Variant)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConfigureMultiOptionPicker")
    On Error GoTo ErrHandler

    mConfirmed = False
    mSelectedValue = vbNullString
    mIgnoreClick = False
    mIsMultiSelectCheckList = True

    Me.Caption = dialogTitle
    
    ' Set up list box for single-select so keyboard highlight is visible, plain style
    Me.lstHotkeys.MultiSelect = 0 ' fmMultiSelectSingle
    Me.lstHotkeys.ListStyle = 0 ' fmListStylePlain
    
    On Error Resume Next
    Me.lstHotkeys.Font.Name = "Segoe UI"
    Me.lstHotkeys.Font.Size = 10
    On Error GoTo ErrHandler
    
    LoadMultiOptionList options, defaultChecked
    ResizeOptionPickerLayout

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "ConfigureMultiOptionPicker", Err
    Resume CleanExit
End Sub

Private Sub LoadMultiOptionList(ByVal options As Variant, ByVal defaultChecked As Variant)
    Dim i As Long
    Dim candidateValue As String
    Dim isChecked As Boolean
    Dim regKey As String
    Dim savedSetting As String

    With Me.lstHotkeys
        .Clear
        .ColumnCount = 1

        If IsArray(options) Then
            For i = LBound(options) To UBound(options)
                candidateValue = Trim$(CStr(options(i)))
                If candidateValue <> vbNullString Then
                    If IsHeaderOption(candidateValue) Then
                        .AddItem candidateValue
                    Else
                        ' Load setting from registry if present
                        regKey = Me.Caption & "_" & candidateValue
                        regKey = Replace(regKey, " ", "_")
                        
                        savedSetting = Infra_Interaction.GetUserPreference("Preferences", regKey, vbNullString)
                        
                        If savedSetting <> vbNullString Then
                            isChecked = CBool(savedSetting)
                        Else
                            isChecked = True
                            If IsArray(defaultChecked) Then
                                If i <= UBound(defaultChecked) And i >= LBound(defaultChecked) Then
                                    isChecked = CBool(defaultChecked(i))
                                End If
                            End If
                        End If
                        
                        If isChecked Then
                            .AddItem CheckedPrefix & candidateValue
                        Else
                            .AddItem UncheckedPrefix & candidateValue
                        End If
                    End If
                End If
            Next i
        End If
    End With
End Sub

Public Property Get WasConfirmed() As Boolean
    WasConfirmed = mConfirmed
End Property

Public Property Get SelectedValue() As String
    SelectedValue = mSelectedValue
End Property

Public Property Get SelectedIndices() As Variant
    Dim result() As Long
    Dim count As Long
    Dim i As Long
    
    count = 0
    For i = 0 To lstHotkeys.ListCount - 1
        Dim isChecked As Boolean
        If mIsMultiSelectCheckList Then
            isChecked = (Left$(lstHotkeys.List(i), PrefixLen) = CheckedPrefix)
        Else
            isChecked = lstHotkeys.Selected(i)
        End If
        
        If isChecked Then
            ReDim Preserve result(0 To count)
            result(count) = i
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        SelectedIndices = Array()
    Else
        SelectedIndices = result
    End If
End Property

Private Sub LoadOptionList(ByVal defaultChoice As String, ByVal options As Variant)
    Dim i As Long
    Dim defaultIndex As Long
    Dim candidateValue As String

    defaultIndex = -1

    With Me.lstHotkeys
        .Clear
        .ColumnCount = 1

        If IsArray(options) Then
            For i = LBound(options) To UBound(options)
                candidateValue = Trim$(CStr(options(i)))
                If candidateValue <> vbNullString Then
                    .AddItem candidateValue
                    If defaultIndex = -1 Then defaultIndex = .ListCount - 1
                    If StrComp(candidateValue, defaultChoice, vbTextCompare) = 0 Then
                        defaultIndex = .ListCount - 1
                    End If
                End If
            Next i
        End If

        If .ListCount > 0 Then
            .ListIndex = defaultIndex
        End If
    End With
End Sub

Private Sub ResizeOptionPickerLayout()
    Const MARGIN_LEFT As Double = 12
    Const MARGIN_RIGHT As Double = 12
    Const MARGIN_TOP As Double = 12
    Const MARGIN_BOTTOM As Double = 12
    Const MIN_LISTBOX_WIDTH As Double = 220
    Const MAX_LISTBOX_WIDTH As Double = 420

    ' Calculate ListBox and form width based on maximum option length
    Dim maxOptionChars As Long
    Dim i As Long
    Dim optionText As String
    maxOptionChars = 0
    For i = 0 To lstHotkeys.ListCount - 1
        optionText = lstHotkeys.List(i)
        If Len(optionText) > maxOptionChars Then
            maxOptionChars = Len(optionText)
        End If
    Next i

    Dim listPadding As Double
    If mIsMultiSelectCheckList Then
        listPadding = 38
    ElseIf lstHotkeys.MultiSelect = 0 Then
        listPadding = 24
    Else
        listPadding = 38
    End If

    ' If the list is long, add width for vertical scrollbar
    Dim isLongList As Boolean
    Dim rowHeight As Double
    Dim maxListHeight As Double
    Dim minListHeight As Double
    
    If lstHotkeys.MultiSelect = 0 And Not mIsMultiSelectCheckList Then
        rowHeight = lstHotkeys.Font.Size + 5
        minListHeight = rowHeight * 2
        maxListHeight = rowHeight * 8
    Else
        rowHeight = lstHotkeys.Font.Size + 8
        minListHeight = rowHeight * 2
        maxListHeight = rowHeight * 12
    End If

    Dim listHeight As Double
    listHeight = (lstHotkeys.ListCount * rowHeight) + 4
    
    If listHeight > maxListHeight Then
        isLongList = True
        listPadding = listPadding + 16
    End If

    Dim targetWidth As Double
    targetWidth = (maxOptionChars * 6.2) + listPadding
    
    Dim listBoxWidth As Double
    listBoxWidth = Application.Max(MIN_LISTBOX_WIDTH, targetWidth)
    If listBoxWidth > MAX_LISTBOX_WIDTH Then listBoxWidth = MAX_LISTBOX_WIDTH

    ' Lay out ListBox
    lstHotkeys.Left = MARGIN_LEFT
    lstHotkeys.Top = MARGIN_TOP
    lstHotkeys.Width = listBoxWidth
    
    lstHotkeys.IntegralHeight = True
    
    If listHeight < minListHeight Then listHeight = minListHeight
    If listHeight > maxListHeight Then listHeight = maxListHeight
    lstHotkeys.Height = listHeight

    ' Set Form Inside dimensions
    Dim desiredInsideWidth As Double
    desiredInsideWidth = listBoxWidth + MARGIN_LEFT + MARGIN_RIGHT
    
    Dim desiredInsideHeight As Double
    
    If mIsMultiSelectCheckList Then
        ' Programmatically show and position btnOK and btnCancel
        btnOK.Visible = True
        btnCancel.Visible = True
        
        Dim btnWidth As Double
        Dim btnHeight As Double
        btnWidth = 60
        btnHeight = 20
        
        btnCancel.Width = btnWidth
        btnCancel.Height = btnHeight
        btnCancel.Top = lstHotkeys.Top + lstHotkeys.Height + 8
        btnCancel.Left = (lstHotkeys.Left + lstHotkeys.Width) - btnWidth
        
        btnOK.Width = btnWidth
        btnOK.Height = btnHeight
        btnOK.Top = btnCancel.Top
        btnOK.Left = btnCancel.Left - btnWidth - 6
        
        desiredInsideHeight = btnCancel.Top + btnCancel.Height + MARGIN_BOTTOM
    Else
        btnOK.Visible = False
        btnCancel.Visible = False
        desiredInsideHeight = lstHotkeys.Top + lstHotkeys.Height + MARGIN_BOTTOM
    End If

    ' Set Form dimensions with safe fallback frame offsets to prevent clipping/cut-off
    Dim frameWidth As Double
    Dim frameHeight As Double
    
    frameWidth = Me.Width - Me.InsideWidth
    frameHeight = Me.Height - Me.InsideHeight
    
    If frameWidth <= 0 Or frameWidth > 40 Then frameWidth = 10
    If frameHeight <= 0 Or frameHeight > 60 Then frameHeight = 34
    
    Me.Width = desiredInsideWidth + frameWidth
    Me.Height = desiredInsideHeight + frameHeight
End Sub

Private Sub UserForm_Activate()
    Dim tracker As Object: Set tracker = Infra_Error.Track("UserForm_Activate")
    On Error GoTo ErrHandler

    ResizeOptionPickerLayout
    
    ' Select the first non-header item by default if checklist
    If mIsMultiSelectCheckList Then
        Dim i As Long
        For i = 0 To lstHotkeys.ListCount - 1
            If Not IsHeader(lstHotkeys.List(i)) Then
                lstHotkeys.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    On Error Resume Next
    lstHotkeys.SetFocus
    On Error GoTo 0

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "UserForm_Activate", Err
    Resume CleanExit
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelSelection
    End If
End Sub

Private Sub ConfirmSelection()
    Dim i As Long
    Dim hasSelection As Boolean
    Dim regKey As String
    Dim itemText As String

    If mIsMultiSelectCheckList Then
        hasSelection = False
        For i = 0 To lstHotkeys.ListCount - 1
            If Not IsHeader(lstHotkeys.List(i)) Then
                If Left$(lstHotkeys.List(i), PrefixLen) = CheckedPrefix Then
                    hasSelection = True
                    Exit For
                End If
            End If
        Next i
        
        If Not hasSelection Then
            MsgBox "Please select at least one option.", vbExclamation, Me.Caption
            Exit Sub
        End If
        
        ' Save multi-select preferences to Registry
        For i = 0 To lstHotkeys.ListCount - 1
            itemText = lstHotkeys.List(i)
            If Not IsHeader(itemText) Then
                Dim cleanText As String
                cleanText = Mid$(itemText, PrefixLen + 1)
                regKey = Me.Caption & "_" & cleanText
                regKey = Replace(regKey, " ", "_")
                
                Dim isChecked As Boolean
                isChecked = (Left$(itemText, PrefixLen) = CheckedPrefix)
                Infra_Interaction.SaveUserPreference "Preferences", regKey, CStr(isChecked)
            End If
        Next i
        
        mConfirmed = True
        Me.Hide
    Else
        ' Standard single-select (or fallback simple multi-select if called)
        If Me.lstHotkeys.MultiSelect = 0 Then
            If lstHotkeys.ListIndex < 0 Then Exit Sub
            mSelectedValue = CStr(lstHotkeys.List(lstHotkeys.ListIndex))
            mConfirmed = True
            Me.Hide
        Else
            hasSelection = False
            For i = 0 To lstHotkeys.ListCount - 1
                If lstHotkeys.Selected(i) Then
                    hasSelection = True
                    Exit For
                End If
            Next i
            
            If Not hasSelection Then
                MsgBox "Please select at least one option.", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            For i = 0 To lstHotkeys.ListCount - 1
                itemText = Trim$(CStr(lstHotkeys.List(i)))
                regKey = Me.Caption & "_" & itemText
                regKey = Replace(regKey, " ", "_")
                Infra_Interaction.SaveUserPreference "Preferences", regKey, CStr(lstHotkeys.Selected(i))
            Next i
            
            mConfirmed = True
            Me.Hide
        End If
    End If
End Sub

Private Sub CancelSelection()
    mConfirmed = False
    mSelectedValue = vbNullString
    Me.Hide
End Sub

Private Sub lstHotkeys_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    If mIsMultiSelectCheckList Then
        ToggleCurrentItem
    Else
        ConfirmSelection
    End If
End Sub

Private Sub lstHotkeys_Click()
    If mIgnoreClick Then Exit Sub
    If Not mIsMultiSelectCheckList And Me.lstHotkeys.MultiSelect = 0 Then
        ConfirmSelection
    End If
End Sub

Private Sub lstHotkeys_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If mIgnoreClick Then Exit Sub
    If mIsMultiSelectCheckList Then
        ToggleCurrentItem
    End If
End Sub

Public Sub HandleKeyDown(ByVal KeyVal As Long, ByVal Shift As Integer, Optional ByRef KeyCodeObj As Object = Nothing)
    Dim tracker As Object: Set tracker = Infra_Error.Track("HandleKeyDown")
    On Error GoTo ErrHandler

    If KeyVal = 13 Then ' vbKeyReturn (Enter)
        ConfirmSelection
        If Not KeyCodeObj Is Nothing Then KeyCodeObj.Value = 0
    ElseIf KeyVal = 27 Then ' vbKeyEscape (Escape)
        CancelSelection
        If Not KeyCodeObj Is Nothing Then KeyCodeObj.Value = 0
    Else
        mIgnoreClick = True
    End If

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "HandleKeyDown", Err
    Resume CleanExit
End Sub

Public Sub HandleKeyUp(ByVal KeyVal As Long, ByVal Shift As Integer)
    Dim tracker As Object: Set tracker = Infra_Error.Track("HandleKeyUp")
    On Error GoTo ErrHandler

    mIgnoreClick = False

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "HandleKeyUp", Err
    Resume CleanExit
End Sub

Private Sub lstHotkeys_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If mIsMultiSelectCheckList Then
        If KeyCode = 38 Then ' Up arrow
            MoveSelectionUp KeyCode
            Exit Sub
        ElseIf KeyCode = 40 Then ' Down arrow
            MoveSelectionDown KeyCode
            Exit Sub
        ElseIf KeyCode = 32 Then ' Spacebar
            ToggleCurrentItem
            KeyCode = 0
            Exit Sub
        End If
    End If
    HandleKeyDown KeyCode.Value, Shift, KeyCode
End Sub

Private Sub lstHotkeys_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleKeyUp KeyCode.Value, Shift
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' vbKeyReturn (Enter)
        ConfirmSelection
        KeyCode = 0
    ElseIf KeyCode = 27 Then ' vbKeyEscape (Escape)
        CancelSelection
        KeyCode = 0
    End If
End Sub

Private Sub btnOK_Click()
    ConfirmSelection
End Sub

Private Sub btnCancel_Click()
    CancelSelection
End Sub

' --- Helper Methods ---

Private Function IsHeader(ByVal txt As String) As Boolean
    IsHeader = (Left$(txt, PrefixLen) <> CheckedPrefix And Left$(txt, PrefixLen) <> UncheckedPrefix)
End Function

Private Function IsHeaderOption(ByVal txt As String) As Boolean
    Dim t As String: t = Trim$(txt)
    If Len(t) > 0 Then
        IsHeaderOption = (Left$(t, 1) = ChrW$(9670))
    Else
        IsHeaderOption = False
    End If
End Function

Private Sub ToggleCurrentItem()
    Dim idx As Long
    idx = lstHotkeys.ListIndex
    If idx < 0 Then Exit Sub
    
    Dim itemText As String
    itemText = lstHotkeys.List(idx)
    
    If IsHeader(itemText) Then Exit Sub
    
    mIgnoreClick = True
    
    Dim cleanText As String
    If Left$(itemText, PrefixLen) = CheckedPrefix Then
        cleanText = Mid$(itemText, PrefixLen + 1)
        lstHotkeys.List(idx) = UncheckedPrefix & cleanText
    ElseIf Left$(itemText, PrefixLen) = UncheckedPrefix Then
        cleanText = Mid$(itemText, PrefixLen + 1)
        lstHotkeys.List(idx) = CheckedPrefix & cleanText
        EnforceMutualExclusivity idx, cleanText
    End If
    
    mIgnoreClick = False
End Sub

Private Sub EnforceMutualExclusivity(ByVal currentIdx As Long, ByVal cleanText As String)
    Dim trimmedText As String
    trimmedText = Trim$(cleanText)
    If Left$(trimmedText, 12) = "Line breaks:" Then
        Dim i As Long
        Dim otherText As String
        For i = 0 To lstHotkeys.ListCount - 1
            If i <> currentIdx Then
                otherText = lstHotkeys.List(i)
                If Not IsHeader(otherText) Then
                    Dim otherClean As String
                    otherClean = Trim$(Mid$(otherText, PrefixLen + 1))
                    If Left$(otherClean, 12) = "Line breaks:" Then
                        lstHotkeys.List(i) = UncheckedPrefix & Mid$(otherText, PrefixLen + 1)
                    End If
                End If
            End If
        Next i
    End If
End Sub

Private Sub MoveSelectionUp(ByRef KeyCode As MSForms.ReturnInteger)
    Dim currIdx As Long
    currIdx = lstHotkeys.ListIndex
    If currIdx <= 0 Then Exit Sub
    
    Dim targetIdx As Long
    targetIdx = currIdx - 1
    
    Do While targetIdx >= 0
        If Not IsHeader(lstHotkeys.List(targetIdx)) Then
            lstHotkeys.ListIndex = targetIdx
            KeyCode = 0
            Exit Sub
        End If
        targetIdx = targetIdx - 1
    Loop
    
    KeyCode = 0
End Sub

Private Sub MoveSelectionDown(ByRef KeyCode As MSForms.ReturnInteger)
    Dim currIdx As Long
    currIdx = lstHotkeys.ListIndex
    If currIdx < 0 Or currIdx >= lstHotkeys.ListCount - 1 Then Exit Sub
    
    Dim targetIdx As Long
    targetIdx = currIdx + 1
    
    Do While targetIdx < lstHotkeys.ListCount
        If Not IsHeader(lstHotkeys.List(targetIdx)) Then
            lstHotkeys.ListIndex = targetIdx
            KeyCode = 0
            Exit Sub
        End If
        targetIdx = targetIdx + 1
    Loop
    
    KeyCode = 0
End Sub
