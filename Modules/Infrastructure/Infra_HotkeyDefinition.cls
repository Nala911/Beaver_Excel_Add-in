VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_HotkeyDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_HotkeyDefinition
' @Category: Infrastructure
' @Description: Typed hotkey definition loaded from config.json.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Public KeyPattern As String
Public MacroName As String
Public Description As String
Public ReleaseTier As String
