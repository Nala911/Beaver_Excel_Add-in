Attribute VB_Name = "Infra_CommandRegistry"
Option Explicit

' @Module: Infra_CommandRegistry
' @Category: Infrastructure
' @Description: Generated command registry mapping entry macros and command names to command classes.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand

Public Function ResolveCommandName(ByVal entryMacro As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("ResolveCommandName")
    On Error GoTo ErrHandler

    Select Case UCase$(Trim$(entryMacro))
        Case "UI_RIBBON.RIBBON_ONWRAP"
            ResolveCommandName = "Wrap"
        Case "UI_RIBBON.RIBBON_ONSTATICSHEETWORKBOOK"
            ResolveCommandName = "StaticSheetWorkbook"
        Case "UI_RIBBON.RIBBON_ONCLEANDATA"
            ResolveCommandName = "CleanData"
        Case "UI_RIBBON.RIBBON_ONBREAKEXTERNALLINKS"
            ResolveCommandName = "BreakExternalLinks"
        Case "UI_RIBBON.RIBBON_ONCONVERTTEXTTOPROPERDATE"
            ResolveCommandName = "DateConversion"
        Case "UI_RIBBON.RIBBON_ONDUPLICATE"
            ResolveCommandName = "Duplicate"
        Case "UI_RIBBON.RIBBON_ONEXPORT"
            ResolveCommandName = "ExportImageOrPdf"
        Case "UI_RIBBON.RIBBON_ONDASHBOARD"
            ResolveCommandName = "Dashboard"
        Case "UI_RIBBON.RIBBON_ONTOGGLEFULLSCREEN"
            ResolveCommandName = "ToggleFullScreen"
        Case "INFRA_HOTKEYS.SHOWHOTKEYSHELP"
            ResolveCommandName = "ShowHotkeysHelp"
        Case "UI_HOTKEYS.HOTKEY_APPLYCUSTOMNUMBERFORMAT"
            ResolveCommandName = "ApplyCustomNumberFormat"
        Case "UI_HOTKEYS.HOTKEY_MAKEPERMANENT"
            ResolveCommandName = "MakePermanent"
        Case "UI_HOTKEYS.HOTKEY_CREATENAMEDSHEET"
            ResolveCommandName = "CreateSheet"
        Case "UI_HOTKEYS.HOTKEY_FILLDOWN"
            ResolveCommandName = "FillDown"
        Case "UI_HOTKEYS.HOTKEY_FILTERBYSELECTEDCELL"
            ResolveCommandName = "FilterByCell"
        Case "UI_HOTKEYS.HOTKEY_PASTEFORMAT"
            ResolveCommandName = "PasteFormat"
        Case "UI_HOTKEYS.HOTKEY_FORMATSELECTEDRANGE"
            ResolveCommandName = "FormatRange"
        Case "UI_HOTKEYS.HOTKEY_BACKSPACE"
            ResolveCommandName = "Backspace"
        Case "UI_HOTKEYS.HOTKEY_DELETE"
            ResolveCommandName = "Delete"
        Case "UI_RIBBON.RIBBON_ONSHOWHOTKEYSHELP"
            ResolveCommandName = "ShowHotkeysHelp"
    End Select

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ResolveCommandName", Err
    Resume CleanExit
End Function

Public Function CreateCommand(ByVal commandName As String) As ICommand
    Dim tracker As Object: Set tracker = Infra_Error.Track("CreateCommand")
    On Error GoTo ErrHandler

    Select Case UCase$(Trim$(commandName))
        Case "WRAP"
            Set CreateCommand = New FeatCmd_Wrap
        Case "STATICSHEETWORKBOOK"
            Set CreateCommand = New FeatCmd_StaticSheetWorkbook
        Case "CLEANDATA"
            Set CreateCommand = New FeatCmd_CleanData
        Case "BREAKEXTERNALLINKS"
            Set CreateCommand = New FeatCmd_BreakExternalLinks
        Case "DATECONVERSION"
            Set CreateCommand = New FeatCmd_DateConversion
        Case "DUPLICATE"
            Set CreateCommand = New FeatCmd_Duplicate
        Case "EXPORTIMAGEORPDF"
            Set CreateCommand = New FeatCmd_ExportImageOrPdf
        Case "DASHBOARD"
            Set CreateCommand = New FeatCmd_Dashboard
        Case "TOGGLEFULLSCREEN"
            Set CreateCommand = New FeatCmd_ToggleFullScreen
        Case "SHOWHOTKEYSHELP"
            Set CreateCommand = New FeatCmd_ShowHotkeysHelp
        Case "APPLYCUSTOMNUMBERFORMAT"
            Set CreateCommand = New FeatCmd_ApplyCustomNumberFormat
        Case "MAKEPERMANENT"
            Set CreateCommand = New FeatCmd_MakePermanent
        Case "CREATESHEET"
            Set CreateCommand = New FeatCmd_CreateSheet
        Case "FILLDOWN"
            Set CreateCommand = New FeatCmd_FillDown
        Case "FILTERBYCELL"
            Set CreateCommand = New FeatCmd_FilterByCell
        Case "PASTEFORMAT"
            Set CreateCommand = New FeatCmd_PasteFormat
        Case "FORMATRANGE"
            Set CreateCommand = New FeatCmd_FormatRange
        Case "BACKSPACE"
            Set CreateCommand = New FeatCmd_Backspace
        Case "DELETE"
            Set CreateCommand = New FeatCmd_Delete
    End Select

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "CreateCommand", Err
    Resume CleanExit
End Function