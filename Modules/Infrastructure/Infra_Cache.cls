VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Infra_Cache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @Module: Infra_Cache
' @Category: Infrastructure
' @Description: Transient shared in-memory cache for commands and evaluations.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Private pCache As Object

Private Sub Class_Initialize()
    Set pCache = CreateObject("Scripting.Dictionary")
End Sub

Public Sub SetVal(ByVal Key As String, ByVal Value As Variant)
    If pCache Is Nothing Then Set pCache = CreateObject("Scripting.Dictionary")
    
    If IsObject(Value) Then
        Set pCache(Key) = Value
    Else
        pCache(Key) = Value
    End If
End Sub

Public Function GetVal(ByVal Key As String) As Variant
    If pCache Is Nothing Then
        GetVal = Empty
        Exit Function
    End If
    
    If Not pCache.Exists(Key) Then
        GetVal = Empty
        Exit Function
    End If
    
    If IsObject(pCache(Key)) Then
        Set GetVal = pCache(Key)
    Else
        GetVal = pCache(Key)
    End If
End Function

Public Function HasVal(ByVal Key As String) As Boolean
    If pCache Is Nothing Then
        HasVal = False
        Exit Function
    End If
    HasVal = pCache.Exists(Key)
End Function

Public Sub RemoveVal(ByVal Key As String)
    If pCache Is Nothing Then Exit Sub
    If pCache.Exists(Key) Then pCache.Remove Key
End Sub

Public Sub Clear()
    If Not pCache Is Nothing Then pCache.RemoveAll
End Sub
