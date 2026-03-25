Attribute VB_Name = "Infra_Diagnostics"
Option Explicit

' @Module: Infra_Diagnostics
' @Category: Infrastructure
' @Description: Structured diagnostics logging for operations, warnings, and failures.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: None

Private Const LOG_FILE_NAME As String = "BeaverAddin.log"

Public Function NextOperationId(ByVal procedureName As String) As String
    On Error Resume Next
    Randomize
    NextOperationId = Format$(Now, "yyyymmddhhnnss") & "_" & procedureName & "_" & Format$(CLng(Rnd() * 100000), "00000")
    On Error GoTo 0
End Function

Public Sub LogOperationStart(ByVal operationId As String, ByVal procedureName As String)
    LogEvent "operation_start", procedureName, operationId, ""
End Sub

Public Sub LogOperationFinish(ByVal operationId As String, ByVal procedureName As String, ByVal elapsedSeconds As Double)
    LogEvent "operation_finish", procedureName, operationId, "elapsed_seconds=" & Format$(elapsedSeconds, "0.000")
End Sub

Public Sub LogWarning(ByVal procedureName As String, ByVal detail As String)
    LogEvent "warning", procedureName, "", detail
End Sub

Public Sub LogError(ByVal procedureName As String, ByVal detail As String)
    LogEvent "error", procedureName, "", detail
End Sub

Public Sub LogEvent(ByVal eventName As String, ByVal procedureName As String, ByVal operationId As String, ByVal detail As String)
    On Error Resume Next

    Dim lineText As String
    lineText = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
               "event=" & eventName & " | " & _
               "procedure=" & procedureName

    If operationId <> "" Then
        lineText = lineText & " | op=" & operationId
    End If
    If detail <> "" Then
        lineText = lineText & " | " & detail
    End If

    Debug.Print lineText
    AppendLineToLog lineText
    On Error GoTo 0
End Sub

Private Sub AppendLineToLog(ByVal lineText As String)
    On Error Resume Next

    Dim fso As Object
    Dim logPath As String
    Dim stream As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = Environ$("TEMP") & "\" & LOG_FILE_NAME
    Set stream = fso.OpenTextFile(logPath, 8, True)
    stream.WriteLine lineText
    stream.Close

    Set stream = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Sub
