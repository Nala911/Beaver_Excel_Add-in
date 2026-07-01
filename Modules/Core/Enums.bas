Attribute VB_Name = "Enums"
Option Explicit

' @Module: Enums
' @Category: Core
' @Description: Centralized shared enumerations for scopes, modes, and command options.
' @ManagedBy: BeaverAddin Agent

' Unified target/execution scope for operations.
Public Enum TargetScope
    TargetScopeSelection = 1
    TargetScopeActiveSheet = 2
    TargetScopeWorkbook = 3
End Enum

' Modes for formula wrapping.
Public Enum WrapMode
    WrapModeCell = 1
    WrapModeTyped = 2
End Enum

' Placement options when creating a new worksheet.
Public Enum SheetInsertPosition
    SheetInsertPositionBeforeCurrent = 1
    SheetInsertPositionAfterCurrent = 2
End Enum

' Capture modes for Undo.
Public Enum UndoCaptureMode
    UndoCaptureFull = 0
    UndoCaptureFormulaOnly = 1
    UndoCaptureFormatOnly = 2
End Enum
