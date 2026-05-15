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

Private Sub UserForm_Initialize()
    mConfirmed = False
    mSelectedValue = vbNullString
End Sub

Public Sub ConfigureOptionPicker(ByVal dialogTitle As String, ByVal promptText As String, ByVal defaultChoice As String, ByVal options As Variant)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ConfigureOptionPicker")
    On Error GoTo ErrHandler

    mConfirmed = False
    mSelectedValue = vbNullString

    Me.Caption = dialogTitle
    btnOK.Caption = "Select"
    LoadOptionList promptText, defaultChoice, options
    ResizeOptionPickerLayout promptText

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ConfigureOptionPicker", Err
    Resume CleanExit
End Sub

Public Property Get WasConfirmed() As Boolean
    WasConfirmed = mConfirmed
End Property

Public Property Get SelectedValue() As String
    SelectedValue = mSelectedValue
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
    Const MIN_LIST_HEIGHT As Double = 36
    Const MAX_LIST_HEIGHT As Double = 84
    Const LIST_ROW_HEIGHT As Double = 18
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

    listHeight = (lstHotkeys.ListCount * LIST_ROW_HEIGHT) + 6
    If listHeight < MIN_LIST_HEIGHT Then listHeight = MIN_LIST_HEIGHT
    If listHeight > MAX_LIST_HEIGHT Then listHeight = MAX_LIST_HEIGHT
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
    If frameHeight < 0 Then frameHeight = 0

    targetHeight = desiredInsideHeight + frameHeight
    If targetHeight < minimumOverallHeight Then targetHeight = minimumOverallHeight

    Me.Height = targetHeight
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
    If lstHotkeys.ListIndex < 0 Then Exit Sub

    mSelectedValue = CStr(lstHotkeys.List(lstHotkeys.ListIndex))
    mConfirmed = True
    Me.Hide
End Sub

Private Sub lstHotkeys_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Cancel = True
    btnOK_Click
End Sub
