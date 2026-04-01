Attribute VB_Name = "Infra_AppState"
Option Explicit

' @Module: Infra_AppState
' @Category: Infrastructure
' @Description: Shared helpers: consistent error dialogs, selection guards, and Desktop path retrieval.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Config, Infra_ActionContext, Infra_Error

' Returns True if the active cell on the active sheet can be modified.
' Checks for sheet protection and the cell's locked status.
Public Function CanModifyActiveCell() As Boolean
    CanModifyActiveCell = AppContainer.State.CanModifyActiveCell
End Function

' Returns True if the current selection is a Range.
' Use as a guard at the top of any macro that requires a range selection.
Public Function IsRangeSelected() As Boolean
    IsRangeSelected = AppContainer.State.IsRangeSelected
End Function

' Captures the current workbook, worksheet, selection, and active-cell state
' into a typed object for downstream feature logic.
Public Function CaptureActionContext() As Infra_ActionContext
    Set CaptureActionContext = AppContainer.State.CaptureActionContext()
End Function

' Returns the path to the current user's Desktop folder.
' Detects OneDrive-synced Desktops for improved reliability.
Public Function GetDesktopPath() As String
    Infra_Error.PushContext "GetDesktopPath"
    On Error GoTo ErrHandler
    
    Dim shell As Object
    Dim path As String
    
    On Error Resume Next
    Set shell = CreateObject("WScript.Shell")
    ' 1. Try registry for the actual user shell folder (most reliable for OneDrive)
    path = shell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Desktop")
    
    ' Expand environment variables (e.g., %USERPROFILE%)
    If path <> "" Then path = shell.ExpandEnvironmentStrings(path)
    
    ' 2. Fallback to WScript.Shell SpecialFolders
    If path = "" Then path = shell.SpecialFolders("Desktop")
    Set shell = Nothing
    On Error GoTo ErrHandler
    
    ' 3. Manual Fallbacks
    If path = "" Then
        ' Check OneDrive environment variable
        Dim oneDrivePath As String
        oneDrivePath = Environ("OneDrive")
        If oneDrivePath <> "" Then
            path = oneDrivePath & "\Desktop"
        Else
            path = Environ("USERPROFILE") & "\Desktop"
        End If
    End If
    
    GetDesktopPath = path

CleanExit:
    Infra_Error.PopContext
    Exit Function

ErrHandler:
    Infra_Error.HandleError "GetDesktopPath", Err
End Function

