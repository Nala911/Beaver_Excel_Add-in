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
' @ManagedBy: BeaverAddin Agent
' @Dependencies: IConfig, Infra_Error, StateStore, BasicCommandContext, Infra_CommandCatalog, Infra_Bootstrap

Private pConfig As IConfig
Private pErrorTracker As Infra_Error
Private pState As StateStore
Private pIsInitialized As Boolean

Private Sub Class_Initialize()
    ' Wait until explicit Init() is called so we can inject mocks if needed
End Sub

Public Sub Initialize(ByVal cfg As IConfig, ByVal errTrk As Infra_Error, ByVal stateStore As StateStore)
    Set pConfig = cfg
    Set pErrorTracker = errTrk
    Set pState = stateStore
    pIsInitialized = True
    Infra_CommandCatalog.ResetCatalog
End Sub

Public Sub EnsureInitialized()
    If pIsInitialized Then Exit Sub
    Initialize Infra_Config, Infra_Error, StateStore
End Sub

Public Property Get Config() As IConfig
    EnsureInitialized
    Set Config = pConfig
End Property

Public Property Get ErrorTracker() As Infra_Error
    EnsureInitialized
    Set ErrorTracker = pErrorTracker
End Property

Public Property Get State() As StateStore
    EnsureInitialized
    Set State = pState
End Property

Public Function ResolveCommand(ByVal commandName As String) As ICommand
    EnsureInitialized
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
        Case "DASHBOARD"
            Set ResolveCommand = New FeatCmd_Dashboard
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
        Case "SHOWHOTKEYSHELP"
            Set ResolveCommand = New FeatCmd_ShowHotkeysHelp
        Case "STATICSHEETWORKBOOK"
            Set ResolveCommand = New FeatCmd_StaticSheetWorkbook
        Case "TOGGLEFULLSCREEN"
            Set ResolveCommand = New FeatCmd_ToggleFullScreen
        Case "WRAPSELECTEDRANGE"
            Set ResolveCommand = New FeatCmd_WrapSelectedRange
    End Select
End Function

Public Function CreateCommandContext(ByVal commandName As String, Optional ByVal entryMacro As String = vbNullString, Optional ByVal sourceName As String = vbNullString, Optional ByVal triggerKind As String = vbNullString) As ICommandContext
    EnsureInitialized
    Dim ctx As New BasicCommandContext
    ctx.CommandName = commandName
    ctx.EntryMacro = entryMacro
    ctx.SourceName = sourceName
    ctx.TriggerKind = triggerKind
    Set ctx.ActionContext = State.CaptureActionContext()
    Set CreateCommandContext = ctx
End Function

Public Sub ExecuteCommand(ByVal commandName As String, Optional ByVal sourceName As String = vbNullString, Optional ByVal triggerKind As String = vbNullString)
    Dim invoker As CommandInvoker
    Dim cmd As ICommand
    Dim ctx As ICommandContext

    Set cmd = ResolveCommand(commandName)
    If cmd Is Nothing Then Err.Raise 5, "AppContainer.ExecuteCommand", "Unknown command: " & commandName

    Set invoker = New CommandInvoker
    Set ctx = CreateCommandContext(commandName, vbNullString, sourceName, triggerKind)
    invoker.ExecuteCommand cmd, ctx
End Sub

Public Sub ExecuteEntryPoint(ByVal entryMacro As String, Optional ByVal sourceName As String = vbNullString, Optional ByVal triggerKind As String = vbNullString)
    Dim commandName As String
    Dim invoker As CommandInvoker
    Dim cmd As ICommand
    Dim ctx As ICommandContext

    EnsureInitialized

    commandName = Infra_CommandCatalog.ResolveCommandName(entryMacro)
    If commandName = vbNullString Then
        Err.Raise 5, "AppContainer.ExecuteEntryPoint", "Unknown entry point: " & entryMacro
    End If

    Set cmd = ResolveCommand(commandName)
    If cmd Is Nothing Then
        Err.Raise 5, "AppContainer.ExecuteEntryPoint", "Unknown command for entry point: " & entryMacro
    End If

    Set invoker = New CommandInvoker
    Set ctx = CreateCommandContext(commandName, entryMacro, sourceName, triggerKind)
    invoker.ExecuteCommand cmd, ctx
End Sub
