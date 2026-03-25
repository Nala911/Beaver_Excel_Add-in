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
    PushContext "CanModifyActiveCell"
    On Error GoTo ErrHandler
    
    ' 1. Guard against non-worksheet active sheets
    If TypeName(ActiveSheet) <> "Worksheet" Then
        CanModifyActiveCell = False
        GoTo CleanExit
    End If

    ' 2. If sheet is not protected, everything is modifiable
    If Not ActiveSheet.ProtectContents Then
        CanModifyActiveCell = True
        GoTo CleanExit
    End If

    ' 3. If sheet is protected, check the active cell's locked status
    ' (ActiveCell could be Nothing in rare edge cases, so check existence)
    If ActiveCell Is Nothing Then
        CanModifyActiveCell = False
        GoTo CleanExit
    End If

    CanModifyActiveCell = Not ActiveCell.Locked

CleanExit:
    PopContext
    Exit Function

ErrHandler:
    HandleError "CanModifyActiveCell", Err
End Function

' Returns True if the current selection is a Range.
' Use as a guard at the top of any macro that requires a range selection.
Public Function IsRangeSelected() As Boolean
    PushContext "IsRangeSelected"
    On Error GoTo ErrHandler
    
    IsRangeSelected = (TypeName(Selection) = "Range")

CleanExit:
    PopContext
    Exit Function

ErrHandler:
    HandleError "IsRangeSelected", Err
End Function

' Captures the current workbook, worksheet, selection, and active-cell state
' into a typed object for downstream feature logic.
Public Function CaptureActionContext() As Infra_ActionContext
    PushContext "CaptureActionContext"
    On Error GoTo ErrHandler
    
    Dim ctx As Infra_ActionContext
    Set ctx = New Infra_ActionContext
    
    On Error Resume Next
    Set ctx.WorkbookRef = ActiveWorkbook
    If TypeName(ActiveSheet) = "Worksheet" Then Set ctx.WorksheetRef = ActiveSheet
    If TypeName(Selection) = "Range" Then
        ctx.HasRangeSelection = True
        Set ctx.SelectionRange = Selection
    End If
    If TypeName(ActiveCell) = "Range" Then Set ctx.ActiveCellRef = ActiveCell
    On Error GoTo ErrHandler
    
    Set CaptureActionContext = ctx

CleanExit:
    PopContext
    Exit Function

ErrHandler:
    HandleError "CaptureActionContext", Err
End Function

' Returns the path to the current user's Desktop folder.
' Detects OneDrive-synced Desktops for improved reliability.
Public Function GetDesktopPath() As String
    PushContext "GetDesktopPath"
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
    PopContext
    Exit Function

ErrHandler:
    HandleError "GetDesktopPath", Err
End Function
