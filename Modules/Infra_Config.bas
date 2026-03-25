Attribute VB_Name = "Infra_Config"
Option Explicit

' @Module: Infra_Config
' @Category: Infrastructure
' @Description: Central registry for all shared constants, now loaded from config.json.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Lib_JsonConverter, Infra_ConfigModel, Infra_HotkeyDefinition, Infra_Error, Infra_Diagnostics

Private pConfigModel As Infra_ConfigModel
Private pRawConfig As Object
Private pIsLoading As Boolean

' --- Public Entry Point ---

' Clears the cached configuration, forcing a reload on the next access.
Public Sub ResetConfig()
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResetConfig")
    On Error GoTo ErrHandler
    
    Set pConfigModel = Nothing
    Set pRawConfig = Nothing

CleanExit:
    Exit Sub
ErrHandler:
    HandleError "ResetConfig", Err
    Resume CleanExit
End Sub

' Returns the singleton instance of the typed configuration model.
Public Property Get Model() As Infra_ConfigModel
    If pConfigModel Is Nothing Then LoadConfig
    Set Model = pConfigModel
End Property

' Shorthand access for Identity (Frequently used)
Public Property Get ADDIN_NAME() As String
    ADDIN_NAME = Model.Name
End Property

Public Property Get ADDIN_VERSION() As String
    ADDIN_VERSION = Model.Version
End Property

Public Property Get DEFAULT_DATE_FORMAT() As String
    DEFAULT_DATE_FORMAT = Model.DefaultDateFormat
End Property

Public Property Get DISPLAY_DATE_FORMAT() As String
    DISPLAY_DATE_FORMAT = Model.DisplayDateFormat
End Property

Public Property Get DEFAULT_NUMBER_FORMAT() As String
    DEFAULT_NUMBER_FORMAT = Model.DefaultNumberFormat
End Property

Public Property Get HEADER_COLOR() As Long
    HEADER_COLOR = Model.HeaderColor
End Property

Public Property Get HEADER_FONT_SIZE() As Long
    HEADER_FONT_SIZE = Model.HeaderFontSize
End Property

Public Property Get DEFAULT_FONT_NAME() As String
    DEFAULT_FONT_NAME = Model.DefaultFontName
End Property

Public Property Get DEFAULT_FONT_SIZE() As Long
    DEFAULT_FONT_SIZE = Model.DefaultFontSize
End Property

Public Property Get MAX_EXPORT_SCALE() As Long
    MAX_EXPORT_SCALE = Model.MaxExportScale
End Property

Public Property Get DEFAULT_EXPORT_SCALE() As Long
    DEFAULT_EXPORT_SCALE = Model.DefaultExportScale
End Property

Public Property Get MAX_COLUMN_WIDTH() As Long
    MAX_COLUMN_WIDTH = Model.MaxColumnWidth
End Property

Public Property Get COLUMN_WIDTH_THRESHOLD() As Long
    COLUMN_WIDTH_THRESHOLD = Model.ColumnWidthThreshold
End Property

' Returns the icon imageMso for a given control ID.
Public Property Get GetIcon(ByVal controlId As String) As String
    If pConfigModel Is Nothing Then LoadConfig
    On Error Resume Next
    GetIcon = pConfigModel.Icons(controlId)
    If Err.Number <> 0 Or GetIcon = "" Then GetIcon = "Help" ' Fallback
    On Error GoTo 0
End Property

' Helper for Infra_Hotkeys to get the Hotkeys collection
Public Property Get Hotkeys() As Collection
    If pConfigModel Is Nothing Then LoadConfig
    Set Hotkeys = pConfigModel.Hotkeys
End Property

Public Property Get RELEASE_TIER() As String
    If pConfigModel Is Nothing Then LoadConfig
    On Error Resume Next
    RELEASE_TIER = CStr(pConfigModel.FeatureFlags("ReleaseTier"))
    If RELEASE_TIER = "" Then RELEASE_TIER = "stable"
    On Error GoTo 0
End Property

Public Property Get INCLUDE_DEV_FEATURES() As Boolean
    If pConfigModel Is Nothing Then LoadConfig
    On Error Resume Next
    INCLUDE_DEV_FEATURES = CBool(pConfigModel.FeatureFlags("IncludeDevFeatures"))
    On Error GoTo 0
End Property

' --- Internal Logic ---

Private Sub LoadConfig()
    If pIsLoading Then Exit Sub
    Dim tracker As Object: Set tracker = Infra_Error.Track("LoadConfig")
    On Error GoTo ErrHandler
    
    pIsLoading = True
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pConfigModel = New Infra_ConfigModel
    ApplyDefaultConfigModel pConfigModel
    
    Dim configPath As String
    configPath = fso.BuildPath(ThisWorkbook.Path, "config.json")
    
    If Not fso.FileExists(configPath) Then
        Debug.Print "BEAVER [WARNING]: config.json not found at " & configPath
        Set pRawConfig = CreateObject("Scripting.Dictionary")
    Else
        Dim jsonString As String
        Dim ts As Object
        Set ts = fso.OpenTextFile(configPath, 1)
        jsonString = ts.ReadAll
        ts.Close
        Set pRawConfig = Lib_JsonConverter.ParseJson(jsonString)
    End If
    
    ' --- Populate Typed Model ---
    With pConfigModel
        ' Identity
        .Name = GetValidatedValue("AddinIdentity", "Name", .Name, "String")
        .Version = GetValidatedValue("AddinIdentity", "Version", .Version, "String")
        
        ' UI Constants
        .DefaultFontName = GetValidatedValue("UIConstants", "DefaultFontName", .DefaultFontName, "String")
        .DefaultFontSize = GetValidatedValue("UIConstants", "DefaultFontSize", .DefaultFontSize, "Long", 6, 72)
        .HeaderFontSize = GetValidatedValue("UIConstants", "HeaderFontSize", .HeaderFontSize, "Long", 6, 72)
        .DefaultNumberFormat = GetValidatedValue("UIConstants", "DefaultNumberFormat", .DefaultNumberFormat, "String")
        .DefaultDateFormat = GetValidatedValue("UIConstants", "DefaultDateFormat", .DefaultDateFormat, "String")
        .DisplayDateFormat = GetValidatedValue("UIConstants", "DisplayDateFormat", .DisplayDateFormat, "String")
        .ColumnWidthThreshold = GetValidatedValue("UIConstants", "ColumnWidthThreshold", .ColumnWidthThreshold, "Long", 5, 255)
        .MaxColumnWidth = GetValidatedValue("UIConstants", "MaxColumnWidth", .MaxColumnWidth, "Long", 5, 255)
        .HeaderColor = ParseHexColor(CStr(GetValue("UIConstants", "HeaderColor", "#AEAAAA")))
        .DefaultExportScale = GetValidatedValue("UIConstants", "DefaultExportScale", .DefaultExportScale, "Long", 1, 10)
        .MaxExportScale = GetValidatedValue("UIConstants", "MaxExportScale", .MaxExportScale, "Long", 1, 20)
        
        ' Icons
        On Error Resume Next
        Set .Icons = pRawConfig("Icons")
        If .Icons Is Nothing Then Set .Icons = CreateObject("Scripting.Dictionary")
        Set .FeatureFlags = GetObjectDictionary("FeatureFlags")
        Set .Hotkeys = CreateHotkeyDefinitions()
        On Error GoTo ErrHandler
    End With

CleanExit:
    pIsLoading = False
    Exit Sub
ErrHandler:
    Infra_Diagnostics.LogWarning "LoadConfig", "Falling back to defaults after config load error."
    If pRawConfig Is Nothing Then Set pRawConfig = CreateObject("Scripting.Dictionary")
    If pConfigModel Is Nothing Then Set pConfigModel = New Infra_ConfigModel
    ApplyDefaultConfigModel pConfigModel
    Resume CleanExit
End Sub

Private Sub ApplyDefaultConfigModel(ByVal model As Infra_ConfigModel)
    If model Is Nothing Then Exit Sub

    model.Name = "Beaver Add-in"
    model.Version = "2.0.0"
    model.DefaultFontName = "Calibri"
    model.DefaultFontSize = 10
    model.HeaderFontSize = 11
    model.DefaultNumberFormat = "#,##0"
    model.DefaultDateFormat = "dd-mmm-yyyy"
    model.DisplayDateFormat = "dd/mm/yyyy"
    model.ColumnWidthThreshold = 40
    model.MaxColumnWidth = 25
    model.HeaderColor = RGB(174, 170, 170)
    model.DefaultExportScale = 3
    model.MaxExportScale = 10
    Set model.Icons = CreateObject("Scripting.Dictionary")
    Set model.FeatureFlags = CreateObject("Scripting.Dictionary")
    Set model.Hotkeys = New Collection
End Sub

Private Function GetObjectDictionary(ByVal Category As String) As Object
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetObjectDictionary")
    On Error GoTo ErrHandler

    On Error Resume Next
    Set GetObjectDictionary = pRawConfig(Category)
    On Error GoTo ErrHandler

    If GetObjectDictionary Is Nothing Then
        Set GetObjectDictionary = CreateObject("Scripting.Dictionary")
    End If

CleanExit:
    Exit Function
ErrHandler:
    Set GetObjectDictionary = CreateObject("Scripting.Dictionary")
    Resume CleanExit
End Function

Private Function CreateHotkeyDefinitions() As Collection
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateHotkeyDefinitions")
    On Error GoTo ErrHandler

    Dim definitions As New Collection
    Dim hotkeyItems As Object
    Dim item As Variant
    Dim hotkey As Infra_HotkeyDefinition

    On Error Resume Next
    Set hotkeyItems = pRawConfig("Hotkeys")
    On Error GoTo ErrHandler

    If hotkeyItems Is Nothing Then
        Set CreateHotkeyDefinitions = definitions
        GoTo CleanExit
    End If

    For Each item In hotkeyItems
        Set hotkey = New Infra_HotkeyDefinition
        hotkey.KeyPattern = GetObjectValue(item, "Key")
        hotkey.MacroName = GetObjectValue(item, "Macro")
        hotkey.Description = GetObjectValue(item, "Description")
        hotkey.ReleaseTier = GetObjectValue(item, "ReleaseTier")
        If hotkey.ReleaseTier = "" Then hotkey.ReleaseTier = "stable"
        definitions.Add hotkey
    Next item

    Set CreateHotkeyDefinitions = definitions

CleanExit:
    Exit Function
ErrHandler:
    Set CreateHotkeyDefinitions = definitions
    Resume CleanExit
End Function

Private Function GetObjectValue(ByVal source As Object, ByVal key As String) As String
    On Error Resume Next
    GetObjectValue = CStr(source(key))
    On Error GoTo 0
End Function

Private Function GetValue(Category As String, Key As String, DefaultValue As Variant) As Variant
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetValue")
    On Error GoTo ErrHandler
    
    On Error Resume Next
    Dim val As Variant
    val = pRawConfig(Category)(Key)
    If Err.Number <> 0 Or IsEmpty(val) Or IsNull(val) Then
        GetValue = DefaultValue
    Else
        GetValue = val
    End If
    On Error GoTo ErrHandler

CleanExit:
    Exit Function
ErrHandler:
    HandleError "GetValue", Err
    Resume CleanExit
End Function

Private Function GetValidatedValue(Category As String, Key As String, DefaultValue As Variant, ExpectedType As String, Optional MinValue As Variant, Optional MaxValue As Variant) As Variant
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetValidatedValue")
    On Error GoTo ErrHandler
    
    Dim val As Variant
    val = GetValue(Category, Key, DefaultValue)
    
    ' Check Type
    Select Case ExpectedType
        Case "Long", "Integer", "Double"
            If Not IsNumeric(val) Then GoTo ValidationFailed
            val = CDbl(val)
        Case "String"
            If Not VarType(val) = vbString Then val = CStr(val)
    End Select
    
    ' Check Range (Numerical only)
    If Not IsMissing(MinValue) Then
        If val < MinValue Then val = MinValue
    End If
    If Not IsMissing(MaxValue) Then
        If val > MaxValue Then val = MaxValue
    End If
    
    GetValidatedValue = val
    GoTo CleanExit

ValidationFailed:
    Debug.Print "BEAVER [CONFIG]: Validation failed for " & Category & "." & Key & ". Using default: " & DefaultValue
    GetValidatedValue = DefaultValue

CleanExit:
    Exit Function
ErrHandler:
    HandleError "GetValidatedValue", Err
    Resume CleanExit
End Function

Private Function ParseHexColor(ByVal hexStr As String) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ParseHexColor")
    On Error GoTo ErrHandler
    
    Dim r As Long, g As Long, b As Long
    
    ' Default fallback (Grey)
    ParseHexColor = RGB(174, 170, 170)
    
    hexStr = Replace(hexStr, "#", "")
    If Len(hexStr) <> 6 Then GoTo CleanExit
    
    r = CLng("&H" & Mid(hexStr, 1, 2))
    g = CLng("&H" & Mid(hexStr, 3, 2))
    b = CLng("&H" & Mid(hexStr, 5, 2))
    
    ParseHexColor = RGB(r, g, b)

CleanExit:
    Exit Function
ErrHandler:
    HandleError "ParseHexColor", Err
    Resume CleanExit
End Function
