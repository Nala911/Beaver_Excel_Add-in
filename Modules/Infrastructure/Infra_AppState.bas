Attribute VB_Name = "Infra_AppState"
Option Explicit

Public ActiveGuard As Infra_AppStateGuard

' @Module: Infra_AppState
' @Category: Infrastructure
' @Description: Shared helpers: consistent error dialogs, selection guards, Desktop path retrieval, and file system utilities.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Config, Infra_ActionContext, Infra_Error

' Returns True if the active cell on the active sheet can be modified.
' Checks for sheet protection and the cell's locked status.
Public Function CanModifyActiveCell() As Boolean
    CanModifyActiveCell = AppContainer.ContextProvider.CanModifyActiveCell
End Function

' Returns True if the current selection is a Range.
' Use as a guard at the top of any macro that requires a range selection.
Public Function IsRangeSelected() As Boolean
    IsRangeSelected = AppContainer.ContextProvider.IsRangeSelected
End Function

' Captures the current workbook, worksheet, selection, and active-cell state
' into a typed object for downstream feature logic.
Public Function CaptureActionContext() As Infra_ActionContext
    Set CaptureActionContext = AppContainer.ContextProvider.CaptureActionContext()
End Function

Public Function CanModifyContext(ByVal ctx As Infra_ActionContext) As Boolean
    CanModifyContext = AppContainer.ContextProvider.CanModifyContext(ctx)
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

' Returns True if the specified file exists.
Public Function FileExists(ByVal filePath As String) As Boolean
    Infra_Error.PushContext "FileExists"
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Set fso = Nothing

CleanExit:
    Infra_Error.PopContext
    Exit Function
ErrHandler:
    Infra_Error.HandleError "FileExists", Err
    Resume CleanExit
End Function

' Combines a folder path and a file name with a backslash if necessary.
Public Function CombinePath(ByVal folderPath As String, ByVal fileName As String) As String
    Infra_Error.PushContext "CombinePath"
    On Error GoTo ErrHandler

    If Right$(folderPath, 1) = "\" Then
        CombinePath = folderPath & fileName
    Else
        CombinePath = folderPath & "\" & fileName
    End If

CleanExit:
    Infra_Error.PopContext
    Exit Function
ErrHandler:
    Infra_Error.HandleError "CombinePath", Err
    Resume CleanExit
End Function

' Sanitizes a file name stem by replacing invalid Windows characters with underscores.
Public Function SanitizeFileNameStem(ByVal fileName As String) As String
    Infra_Error.PushContext "SanitizeFileNameStem"
    On Error GoTo ErrHandler

    Dim invalidChars As Variant
    Dim item As Variant

    fileName = Trim$(fileName)
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each item In invalidChars
        fileName = Replace(fileName, CStr(item), "_")
    Next item

    Do While InStr(fileName, "__") > 0
        fileName = Replace(fileName, "__", "_")
    Loop

    fileName = Trim$(fileName)
    If Right$(fileName, 1) = "." Then fileName = Left$(fileName, Len(fileName) - 1)
    SanitizeFileNameStem = fileName

CleanExit:
    Infra_Error.PopContext
    Exit Function
ErrHandler:
    Infra_Error.HandleError "SanitizeFileNameStem", Err
    Resume CleanExit
End Function

' Ensures the file path ends with the specified extension (without dot).
Public Function EnsureExtension(ByVal selectedPath As String, ByVal extensionWithoutDot As String) As String
    Infra_Error.PushContext "EnsureExtension"
    On Error GoTo ErrHandler

    Dim expectedExtension As String

    expectedExtension = "." & LCase$(extensionWithoutDot)
    If LCase$(Right$(selectedPath, Len(expectedExtension))) <> expectedExtension Then
        EnsureExtension = selectedPath & expectedExtension
    Else
        EnsureExtension = selectedPath
    End If

CleanExit:
    Infra_Error.PopContext
    Exit Function
ErrHandler:
    Infra_Error.HandleError "EnsureExtension", Err
    Resume CleanExit
End Function

' Checks if a file name is valid under Windows naming conventions.
Public Function IsValidWindowsFileName(ByVal fileName As String) As Boolean
    Infra_Error.PushContext "IsValidWindowsFileName"
    On Error GoTo ErrHandler

    Dim invalidChars As Variant
    Dim item As Variant

    fileName = Trim$(fileName)
    If fileName = vbNullString Then GoTo CleanExit
    If Right$(fileName, 1) = "." Then GoTo CleanExit

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each item In invalidChars
        If InStr(1, fileName, CStr(item), vbBinaryCompare) > 0 Then GoTo CleanExit
    Next item

    IsValidWindowsFileName = True

CleanExit:
    Infra_Error.PopContext
    Exit Function
ErrHandler:
    Infra_Error.HandleError "IsValidWindowsFileName", Err
    Resume CleanExit
End Function
