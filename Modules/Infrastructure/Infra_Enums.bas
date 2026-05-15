Attribute VB_Name = "Infra_Enums"
Option Explicit

' @Module: Infra_Enums
' @Category: Infrastructure
' @Description: Centralized shared enumerations for scopes, modes, and command options.
' @ManagedBy: BeaverAddin Agent

' Scope for cleaning data operations.
Public Enum CleanDataScope
    CleanDataScopeSelection = 1
    CleanDataScopeActiveSheet = 2
    CleanDataScopeWorkbook = 3
End Enum

' Scope for formula-to-value conversion.
Public Enum StaticConversionScope
    StaticConversionScopeActiveSheet = 1
    StaticConversionScopeWorkbook = 2
End Enum

' Modes for formula wrapping.
Public Enum WrapMode
    WrapModeCell = 1
    WrapModeTyped = 2
End Enum

' Scope for breaking external links.
Public Enum BreakLinksScope
    BreakLinksScopeActiveSheet = 1
    BreakLinksScopeWorkbook = 2
End Enum

' Placement options when creating a new worksheet.
Public Enum SheetInsertPosition
    SheetInsertPositionBeforeCurrent = 1
    SheetInsertPositionAfterCurrent = 2
End Enum
