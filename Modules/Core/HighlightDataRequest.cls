VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HighlightDataRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRequest

' @Module: HighlightDataRequest
' @Category: Core
' @Description: Typed request describing the target scope, action context, and data highlighting options.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ActionContext, TargetScope

Public Context As ActionContext
Public Scope As TargetScope

Public HighlightInconsistentFormulas As Boolean
Public HighlightDuplicates As Boolean
Public HighlightErrors As Boolean
Public HighlightHardcodedValues As Boolean
Public HighlightDataValidations As Boolean
Public HighlightConditionalFormatting As Boolean

Private Sub Class_Initialize()
    HighlightInconsistentFormulas = True
    HighlightDuplicates = True
    HighlightErrors = False
    HighlightHardcodedValues = False
    HighlightDataValidations = False
    HighlightConditionalFormatting = False
End Sub

Private Property Get IRequest_Context() As ActionContext
    Set IRequest_Context = Me.Context
End Property

Private Property Get IRequest_Scope() As TargetScope
    IRequest_Scope = Me.Scope
End Property
