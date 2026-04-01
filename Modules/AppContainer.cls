VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: AppContainer
' @Category: Infrastructure
' @Description: Central Dependency Injection container / Service Locator. Uses PredeclaredId = True for global singleton access.

Public Config As Object
Public ErrorTracker As Object
Public State As Object

Private Sub Class_Initialize()
    ' Wait until explicit Init() is called so we can inject mocks if needed
End Sub

Public Sub Init(ByVal cfg As Object, ByVal errTrk As Object, ByVal stateStore As Object)
    Set Config = cfg
    Set ErrorTracker = errTrk
    Set State = stateStore
End Sub

Public Function ResolveCommand(ByVal commandName As String) As ICommand
    Select Case UCase$(Trim$(commandName))
        Case "APPLYCUSTOMNUMBERFORMAT"
            Set ResolveCommand = New FeatCmd_ApplyCustomNumberFormat
        Case "BACKSPACE"
            Set ResolveCommand = New FeatCmd_Backspace
        Case "BREAKEXTERNALLINKS"
            Set ResolveCommand = New FeatCmd_BreakExternalLinks
        Case "CLEANDATA"
            Set ResolveCommand = New FeatCmd_CleanData
        Case "CREATESHEET"
            Set ResolveCommand = New FeatCmd_CreateSheet
        Case "DATECONVERSION"
            Set ResolveCommand = New FeatCmd_DateConversion
        Case "DELETE"
            Set ResolveCommand = New FeatCmd_Delete
        Case "DUPLICATE"
            Set ResolveCommand = New FeatCmd_Duplicate
        Case "EXPORTIMAGEORPDF"
            Set ResolveCommand = New FeatCmd_ExportImageOrPdf
        Case "FILLDOWN"
            Set ResolveCommand = New FeatCmd_FillDown
        Case "FILTERBYCELL"
            Set ResolveCommand = New FeatCmd_FilterByCell
        Case "FORMATRANGE"
            Set ResolveCommand = New FeatCmd_FormatRange
        Case "MAKEPERMANENT"
            Set ResolveCommand = New FeatCmd_MakePermanent
        Case "MERGEFORMULAS"
            Set ResolveCommand = New FeatCmd_MergeFormulas
        Case "PASTEFORMAT"
            Set ResolveCommand = New FeatCmd_PasteFormat
        Case "STATICSHEETWORKBOOK"
            Set ResolveCommand = New FeatCmd_StaticSheetWorkbook
        Case "TOGGLEFULLSCREEN"
            Set ResolveCommand = New FeatCmd_ToggleFullScreen
        Case "WRAPSELECTEDRANGE"
            Set ResolveCommand = New FeatCmd_WrapSelectedRange
    End Select
End Function

Public Function CreateCommandContext(Optional ByVal sourceName As String = vbNullString, Optional ByVal triggerKind As String = vbNullString) As ICommandContext
    Dim ctx As New BasicCommandContext
    ctx.SourceName = sourceName
    ctx.TriggerKind = triggerKind
    Set CreateCommandContext = ctx
End Function

Public Sub ExecuteCommand(ByVal commandName As String, Optional ByVal sourceName As String = vbNullString, Optional ByVal triggerKind As String = vbNullString)
    Dim invoker As CommandInvoker
    Dim cmd As ICommand
    Dim ctx As ICommandContext

    Set cmd = ResolveCommand(commandName)
    If cmd Is Nothing Then Err.Raise 5, "AppContainer.ExecuteCommand", "Unknown command: " & commandName

    Set invoker = New CommandInvoker
    Set ctx = CreateCommandContext(sourceName, triggerKind)
    invoker.ExecuteCommand cmd, ctx
End Sub
