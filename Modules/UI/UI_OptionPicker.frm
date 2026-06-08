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
Private mPromptText As String

Private Sub UserForm_Initialize()
    mConfirmed = False
    mSelectedValue = vbNullString
End Sub

Public Sub ConfigureOptionPicker(ByVal dialogTitle As String, ByVal promptText As String, ByVal defaultChoice As String, ByVal options As Variant)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConfigureOptionPicker")
    On Error GoTo ErrHandler

    mConfirmed = False
    mSelectedValue = vbNullString
    mPromptText = promptText

    Me.Caption = dialogTitle
    btnOK.Caption = "Select"
    
    ' Reset list style to default single-select
    Me.lstHotkeys.MultiSelect = 0 ' fmMultiSelectSingle
    Me.lstHotkeys.ListStyle = 0 ' fmListStylePlain
    
    LoadOptionList promptText, defaultChoice, options
    ResizeOptionPickerLayout promptText

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
    mPromptText = promptText

    Me.Caption = dialogTitle
    btnOK.Caption = "Select"
    
    ' Configure list box for multi-select checkboxes
    Me.lstHotkeys.MultiSelect = 1 ' fmMultiSimple
    Me.lstHotkeys.ListStyle = 1 ' fmListStyleOption
    
    LoadMultiOptionList promptText, options, defaultChecked
    ResizeOptionPickerLayout promptText

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ConfigureMultiOptionPicker", Err
    Resume CleanExit
End Sub

Private Sub LoadMultiOptionList(ByVal promptText As String, ByVal options As Variant, ByVal defaultChecked As Variant)
    Dim i As Long
    Dim candidateValue As String
    Dim isChecked As Boolean
    Dim regKey As String
    Dim savedSetting As String

    With Me.lstHotkeys
        .Clear
        .ColumnCount = 1

        ConfigurePromptLabel promptText

        If IsArray(options) Then
            For i = LBound(options) To UBound(options)
                candidateValue = Trim$(CStr(options(i)))
                If candidateValue <> vbNullString Then
                    .AddItem candidateValue
                    
                    ' Load setting from registry if present
                    regKey = Me.Caption & "_" & candidateValue
                    regKey = Replace(regKey, " ", "_")
                    
                    On Error Resume Next
                    savedSetting = GetSetting(appname:="BeaverAddin", section:="Preferences", key:=regKey, Default:=vbNullString)
                    On Error GoTo 0
                    
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
                    
                    .Selected(.ListCount - 1) = isChecked
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
        If lstHotkeys.Selected(i) Then
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



Private Sub LoadOptionList(ByVal promptText As String, ByVal defaultChoice As String, ByVal options As Variant)
    Dim i As Long
    Dim defaultIndex As Long
    Dim candidateValue As String

    defaultIndex = -1

    With Me.lstHotkeys
        .Clear
        .ColumnCount = 1

        ConfigurePromptLabel promptText

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

Private Sub ResizeOptionPickerLayout(ByVal promptText As String)
    Const MIN_FORM_HEIGHT As Double = 150
    Const FORM_BOTTOM_PADDING As Double = 18
    Const CONTROL_GAP As Double = 6
    Const BUTTON_GAP As Double = 8
    Const LABEL_LINE_HEIGHT As Double = 13
    Const LABEL_MIN_HEIGHT As Double = 18
    Const MIN_FORM_WIDTH As Double = 320
    Const MAX_FORM_WIDTH As Double = 380
    Const CHARS_PER_LINE As Long = 54

    Dim promptLabel As Object
    Dim promptLines As Long
    Dim estimatedLabelHeight As Double
    Dim listHeight As Double
    Dim targetWidth As Double
    Dim rowHeight As Double
    Dim maxListHeight As Double
    Dim minListHeight As Double

    Set promptLabel = GetPromptLabel()

    targetWidth = Me.Width
    If targetWidth < MIN_FORM_WIDTH Then targetWidth = MIN_FORM_WIDTH
    If targetWidth > MAX_FORM_WIDTH Then targetWidth = MAX_FORM_WIDTH
    Me.Width = targetWidth

    If Not promptLabel Is Nothing Then
        promptLabel.Left = lstHotkeys.Left
        promptLabel.Width = lstHotkeys.Width
        promptLabel.WordWrap = True

        promptLines = EstimatePromptLineCount(promptText, CHARS_PER_LINE)
        estimatedLabelHeight = LABEL_LINE_HEIGHT * promptLines
        If estimatedLabelHeight < LABEL_MIN_HEIGHT Then estimatedLabelHeight = LABEL_MIN_HEIGHT
        promptLabel.Height = estimatedLabelHeight

        lstHotkeys.Top = promptLabel.Top + promptLabel.Height + CONTROL_GAP
    End If

    If Me.lstHotkeys.MultiSelect = 0 Then
        rowHeight = 15
        minListHeight = 30
        maxListHeight = 120
    Else
        rowHeight = 18
        minListHeight = 36
        maxListHeight = 180
    End If

    listHeight = (lstHotkeys.ListCount * rowHeight) + 4
    If listHeight < minListHeight Then listHeight = minListHeight
    If listHeight > maxListHeight Then listHeight = maxListHeight
    lstHotkeys.Height = listHeight

    btnOK.Top = lstHotkeys.Top + lstHotkeys.Height + BUTTON_GAP
    SetFormInsideHeight btnOK.Top + btnOK.Height + FORM_BOTTOM_PADDING, MIN_FORM_HEIGHT
End Sub



Private Sub ConfigurePromptLabel(ByVal promptText As String)
    Dim ctrl As Object
    Dim fallbackText As String

    fallbackText = Trim$(Replace(promptText, vbCrLf, " "))
    If fallbackText = vbNullString Then fallbackText = "Choose an option"

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            ctrl.Caption = promptText
            ctrl.Visible = True
            Exit Sub
        End If
    Next ctrl

    Me.Caption = Me.Caption & " - " & fallbackText
End Sub

Private Function GetPromptLabel() As Object
    Dim ctrl As Object

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            Set GetPromptLabel = ctrl
            Exit Function
        End If
    Next ctrl
End Function

Private Function EstimatePromptLineCount(ByVal promptText As String, ByVal approxCharsPerLine As Long) As Long
    Dim segments() As String
    Dim i As Long
    Dim segmentLength As Long
    Dim totalLines As Long

    If approxCharsPerLine < 1 Then approxCharsPerLine = 1

    segments = Split(promptText, vbCrLf)
    totalLines = 0

    For i = LBound(segments) To UBound(segments)
        segmentLength = Len(Trim$(segments(i)))
        If segmentLength = 0 Then
            totalLines = totalLines + 1
        Else
            totalLines = totalLines + ((segmentLength - 1) \ approxCharsPerLine) + 1
        End If
    Next i

    If totalLines < 1 Then totalLines = 1
    EstimatePromptLineCount = totalLines
End Function

Private Sub SetFormInsideHeight(ByVal desiredInsideHeight As Double, Optional ByVal minimumOverallHeight As Double = 0)
    Dim frameHeight As Double
    Dim targetHeight As Double

    frameHeight = Me.Height - Me.InsideHeight
    ' Fall back to standard title bar + border height of 28 points if calculation returns 0 or negative (due to form not shown yet) or unreasonably large
    If frameHeight <= 0 Or frameHeight > 40 Then
        frameHeight = 28
    End If

    targetHeight = desiredInsideHeight + frameHeight
    If targetHeight < minimumOverallHeight Then targetHeight = minimumOverallHeight

    Me.Height = targetHeight
End Sub

Private Sub UserForm_Activate()
    Dim tracker As Object: Set tracker = Infra_Error.Track("UserForm_Activate")
    On Error GoTo ErrHandler

    ResizeOptionPickerLayout mPromptText

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "UserForm_Activate", Err
    Resume CleanExit
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = True
        mConfirmed = False
        mSelectedValue = vbNullString
        Me.Hide
    End If
End Sub

Private Sub btnOK_Click()
    Dim i As Long
    Dim hasSelection As Boolean
    Dim regKey As String
    Dim itemText As String

    If Me.lstHotkeys.MultiSelect = 0 Then
        If lstHotkeys.ListIndex < 0 Then Exit Sub
        mSelectedValue = CStr(lstHotkeys.List(lstHotkeys.ListIndex))
        mConfirmed = True
    Else
        hasSelection = False
        For i = 0 To lstHotkeys.ListCount - 1
            If lstHotkeys.Selected(i) Then
                hasSelection = True
                Exit For
            End If
        Next i
        
        If Not hasSelection Then
            MsgBox "Please select at least one cleaning option.", vbExclamation, Me.Caption
            Exit Sub
        End If
        
        ' Save multi-select preferences to Registry
        On Error Resume Next
        For i = 0 To lstHotkeys.ListCount - 1
            itemText = Trim$(CStr(lstHotkeys.List(i)))
            regKey = Me.Caption & "_" & itemText
            regKey = Replace(regKey, " ", "_")
            SaveSetting appname:="BeaverAddin", section:="Preferences", key:=regKey, setting:=CStr(lstHotkeys.Selected(i))
        Next i
        On Error GoTo 0
        
        mConfirmed = True
    End If
    Me.Hide
End Sub

Private Sub lstHotkeys_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Cancel = True
    btnOK_Click
End Sub
