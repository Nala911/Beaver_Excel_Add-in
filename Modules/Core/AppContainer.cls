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
' @Category: Core
' @Description: Central Dependency Injection container / Service Locator. Uses PredeclaredId = True for global singleton access.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: IConfig, Infra_Error, ExcelContextProvider, BasicCommandContext, Infra_CommandRegistry, Infra_Bootstrap

Private pConfig As IConfig
Private pErrorTracker As Infra_Error
Private pContextProvider As ExcelContextProvider
Private pIsInitialized As Boolean

Private Sub Class_Initialize()
    ' Wait until explicit Init() is called so we can inject mocks if needed.
    ' This lazy initialization pattern ensures dependencies are loaded correctly.
End Sub

Public Sub Initialize(ByVal cfg As IConfig, ByVal errTrk As Infra_Error, ByVal contextProvider As ExcelContextProvider)
    Set pConfig = cfg
    Set pErrorTracker = errTrk
    Set pContextProvider = contextProvider
    pIsInitialized = True
End Sub

Public Sub EnsureInitialized()
    If pIsInitialized Then Exit Sub
    Initialize Infra_Config, Infra_Error, ExcelContextProvider
End Sub

Public Property Get Config() As IConfig
    EnsureInitialized
    Set Config = pConfig
End Property

Public Property Get ErrorTracker() As Infra_Error
    EnsureInitialized
    Set ErrorTracker = pErrorTracker
End Property

Public Property Get ContextProvider() As ExcelContextProvider
    EnsureInitialized
    Set ContextProvider = pContextProvider
End Property

Public Function ResolveCommand(ByVal commandName As String) As ICommand
    EnsureInitialized
    Set ResolveCommand = Infra_CommandRegistry.CreateCommand(commandName)
End Function

Public Function CreateCommandContext(ByVal commandName As String, Optional ByVal entryMacro As String = vbNullString, Optional ByVal sourceName As String = vbNullString, Optional ByVal triggerKind As String = vbNullString) As ICommandContext
    EnsureInitialized
    Dim ctx As New BasicCommandContext
    ctx.CommandName = commandName
    ctx.EntryMacro = entryMacro
    ctx.SourceName = sourceName
    ctx.TriggerKind = triggerKind
    Set ctx.ActionContext = ContextProvider.CaptureActionContext()
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

    commandName = Infra_CommandRegistry.ResolveCommandName(entryMacro)
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
