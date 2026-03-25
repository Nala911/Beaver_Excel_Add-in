VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_ConfigModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_ConfigModel
' @Category: Infrastructure
' @Description: Typed data model for the application configuration.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

' --- Private Backing Variables ---
Private pDefaultFontName As String
Private pDefaultFontSize As Long
Private pHeaderFontSize As Long
Private pDefaultNumberFormat As String
Private pDefaultDateFormat As String
Private pDisplayDateFormat As String
Private pColumnWidthThreshold As Long
Private pMaxColumnWidth As Long
Private pHeaderColor As Long
Private pDefaultExportScale As Long
Private pMaxExportScale As Long
Private pName As String
Private pVersion As String
Private pIcons As Object
Private pFeatureFlags As Object
Private pHotkeys As Collection

' --- UI Constants (Read-Only) ---

Public Property Get DefaultFontName() As String
    DefaultFontName = pDefaultFontName
End Property
Public Property Let DefaultFontName(ByVal v As String)
    pDefaultFontName = v ' @Internal
End Property

Public Property Get DefaultFontSize() As Long
    DefaultFontSize = pDefaultFontSize
End Property
Public Property Let DefaultFontSize(ByVal v As Long)
    pDefaultFontSize = v ' @Internal
End Property

Public Property Get HeaderFontSize() As Long
    HeaderFontSize = pHeaderFontSize
End Property
Public Property Let HeaderFontSize(ByVal v As Long)
    pHeaderFontSize = v ' @Internal
End Property

Public Property Get DefaultNumberFormat() As String
    DefaultNumberFormat = pDefaultNumberFormat
End Property
Public Property Let DefaultNumberFormat(ByVal v As String)
    pDefaultNumberFormat = v ' @Internal
End Property

Public Property Get DefaultDateFormat() As String
    DefaultDateFormat = pDefaultDateFormat
End Property
Public Property Let DefaultDateFormat(ByVal v As String)
    pDefaultDateFormat = v ' @Internal
End Property

Public Property Get DisplayDateFormat() As String
    DisplayDateFormat = pDisplayDateFormat
End Property
Public Property Let DisplayDateFormat(ByVal v As String)
    pDisplayDateFormat = v ' @Internal
End Property

Public Property Get ColumnWidthThreshold() As Long
    ColumnWidthThreshold = pColumnWidthThreshold
End Property
Public Property Let ColumnWidthThreshold(ByVal v As Long)
    pColumnWidthThreshold = v ' @Internal
End Property

Public Property Get MaxColumnWidth() As Long
    MaxColumnWidth = pMaxColumnWidth
End Property
Public Property Let MaxColumnWidth(ByVal v As Long)
    pMaxColumnWidth = v ' @Internal
End Property

Public Property Get HeaderColor() As Long
    HeaderColor = pHeaderColor
End Property
Public Property Let HeaderColor(ByVal v As Long)
    pHeaderColor = v ' @Internal
End Property

Public Property Get DefaultExportScale() As Long
    DefaultExportScale = pDefaultExportScale
End Property
Public Property Let DefaultExportScale(ByVal v As Long)
    pDefaultExportScale = v ' @Internal
End Property

Public Property Get MaxExportScale() As Long
    MaxExportScale = pMaxExportScale
End Property
Public Property Let MaxExportScale(ByVal v As Long)
    pMaxExportScale = v ' @Internal
End Property

' --- Addin Identity (Read-Only) ---

Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(ByVal v As String)
    pName = v ' @Internal
End Property

Public Property Get Version() As String
    Version = pVersion
End Property
Public Property Let Version(ByVal v As String)
    pVersion = v ' @Internal
End Property

' --- Icons ---

Public Property Get Icons() As Object
    Set Icons = pIcons
End Property
Public Property Set Icons(ByVal v As Object)
    Set pIcons = v ' @Internal
End Property

' --- Feature Flags ---

Public Property Get FeatureFlags() As Object
    Set FeatureFlags = pFeatureFlags
End Property
Public Property Set FeatureFlags(ByVal v As Object)
    Set pFeatureFlags = v ' @Internal
End Property

' --- Hotkeys ---

Public Property Get Hotkeys() As Collection
    Set Hotkeys = pHotkeys
End Property
Public Property Set Hotkeys(ByVal v As Collection)
    Set pHotkeys = v ' @Internal
End Property
