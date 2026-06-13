VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_CleanDataRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_CleanDataRequest
' @Category: Core
' @Description: Typed request describing the target scope, action context, and data cleaning options.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_ActionContext, TargetScope

Public Context As Infra_ActionContext
Public Scope As TargetScope

Public CleanTrimSpaces As Boolean
Public CleanNonPrintables As Boolean
Public CleanInvisibleChars As Boolean
Public CleanReplaceLineBreaksWithSpace As Boolean
Public CleanRemoveLineBreaks As Boolean
Public CleanStandardizeLineBreaks As Boolean
Public CleanConvertNumbers As Boolean
Public CleanBrokenNames As Boolean
Public CleanSpecialSymbols As Boolean
Public CleanStandardizeDashes As Boolean
Public CleanRemoveAccents As Boolean

Private Sub Class_Initialize()
    CleanTrimSpaces = True
    CleanNonPrintables = True
    CleanInvisibleChars = True
    CleanReplaceLineBreaksWithSpace = False
    CleanRemoveLineBreaks = False
    CleanStandardizeLineBreaks = False
    CleanConvertNumbers = False
    CleanBrokenNames = True
    CleanSpecialSymbols = False
    CleanStandardizeDashes = False
    CleanRemoveAccents = False
End Sub
