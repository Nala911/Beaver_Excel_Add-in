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
        Case "BTNWRAP"
            ResolveCommandName = "Wrap"
        Case "UI_RIBBON.RIBBON_ONSTATICSHEETWORKBOOK"
            ResolveCommandName = "StaticSheetWorkbook"
        Case "BTNSTATICSHEETWORKBOOK"
            ResolveCommandName = "StaticSheetWorkbook"
        Case "UI_RIBBON.RIBBON_ONCLEANDATA"
            ResolveCommandName = "CleanData"
        Case "BTNCLEANDATA"
            ResolveCommandName = "CleanData"
        Case "BTNMODIFYDATA"
            ResolveCommandName = "ModifyData"
        Case "UI_RIBBON.RIBBON_ONDATEFIXER"
            ResolveCommandName = "DateFixer"
        Case "BTNDATEFIXER"
            ResolveCommandName = "DateFixer"
        Case "UI_RIBBON.RIBBON_ONCASEFIXER"
            ResolveCommandName = "CaseFixer"
        Case "BTNCASEFIXER"
            ResolveCommandName = "CaseFixer"
        Case "BTNHIGHLIGHTDATA"
            ResolveCommandName = "HighlightData"
        Case "UI_RIBBON.RIBBON_ONHIGHLIGHTINCONSISTENTFORMULAS"
            ResolveCommandName = "HighlightInconsistentFormulas"
        Case "BTNHIGHLIGHTINCONSISTENTFORMULAS"
            ResolveCommandName = "HighlightInconsistentFormulas"
        Case "UI_RIBBON.RIBBON_ONHIGHLIGHTDUPLICATES"
            ResolveCommandName = "HighlightDuplicates"
        Case "BTNHIGHLIGHTDUPLICATES"
            ResolveCommandName = "HighlightDuplicates"
        Case "UI_RIBBON.RIBBON_ONHIGHLIGHTERRORS"
            ResolveCommandName = "HighlightErrors"
        Case "BTNHIGHLIGHTERRORS"
            ResolveCommandName = "HighlightErrors"
        Case "UI_RIBBON.RIBBON_ONHIGHLIGHTHARDCODEDVALUES"
            ResolveCommandName = "HighlightHardcodedValues"
        Case "BTNHIGHLIGHTHARDCODEDVALUES"
            ResolveCommandName = "HighlightHardcodedValues"
        Case "UI_RIBBON.RIBBON_ONHIGHLIGHTDATAVALIDATIONS"
            ResolveCommandName = "HighlightDataValidations"
        Case "BTNHIGHLIGHTDATAVALIDATIONS"
            ResolveCommandName = "HighlightDataValidations"
        Case "UI_RIBBON.RIBBON_ONHIGHLIGHTCONDITIONALFORMATTING"
            ResolveCommandName = "HighlightConditionalFormatting"
        Case "BTNHIGHLIGHTCONDITIONALFORMATTING"
            ResolveCommandName = "HighlightConditionalFormatting"
        Case "UI_RIBBON.RIBBON_ONHIGHLIGHTNAMEDRANGES"
            ResolveCommandName = "HighlightNamedRanges"
        Case "BTNHIGHLIGHTNAMEDRANGES"
            ResolveCommandName = "HighlightNamedRanges"
        Case "UI_RIBBON.RIBBON_ONHIGHLIGHTFINANCIALMODELLING"
            ResolveCommandName = "HighlightFinancialModelling"
        Case "BTNHIGHLIGHTFINANCIALMODELLING"
            ResolveCommandName = "HighlightFinancialModelling"
        Case "UI_RIBBON.RIBBON_ONCLEARHIGHLIGHTS"
            ResolveCommandName = "ClearHighlights"
        Case "BTNCLEARHIGHLIGHTS"
            ResolveCommandName = "ClearHighlights"
        Case "UI_RIBBON.RIBBON_ONBREAKEXTERNALLINKS"
            ResolveCommandName = "BreakExternalLinks"
        Case "BTNBREAKLINKS"
            ResolveCommandName = "BreakExternalLinks"
        Case "UI_RIBBON.RIBBON_ONDUPLICATE"
            ResolveCommandName = "Duplicate"
        Case "BTNDUPLICATE"
            ResolveCommandName = "Duplicate"
        Case "BTNEXPORT"
            ResolveCommandName = "ExportImageOrPdf"
        Case "UI_RIBBON.RIBBON_ONEXPORTPNG"
            ResolveCommandName = "ExportPng"
        Case "BTNEXPORTPNG"
            ResolveCommandName = "ExportPng"
        Case "UI_RIBBON.RIBBON_ONEXPORTPDF"
            ResolveCommandName = "ExportPdf"
        Case "BTNEXPORTPDF"
            ResolveCommandName = "ExportPdf"
        Case "UI_RIBBON.RIBBON_ONSHOWHELPCENTER"
            ResolveCommandName = "ShowHelpCenter"
        Case "BTNHELPCENTER"
            ResolveCommandName = "ShowHelpCenter"
        Case "UI_RIBBON.RIBBON_ONHELLOWORLD"
            ResolveCommandName = "HelloWorld"
        Case "BTNHELLOWORLD"
            ResolveCommandName = "HelloWorld"
        Case "UI_RIBBON.RIBBON_ONTABLEOFCONTENTS"
            ResolveCommandName = "TableOfContents"
        Case "BTNTABLEOFCONTENTS"
            ResolveCommandName = "TableOfContents"
        Case "UI_RIBBON.RIBBON_ONUNMERGEFILL"
            ResolveCommandName = "UnmergeFill"
        Case "BTNUNMERGEFILL"
            ResolveCommandName = "UnmergeFill"
        Case "UI_RIBBON.RIBBON_ONFORCENUMBER"
            ResolveCommandName = "ForceNumber"
        Case "BTNFORCENUMBER"
            ResolveCommandName = "ForceNumber"
        Case "UI_RIBBON.RIBBON_ONCREATENAMEDRANGES"
            ResolveCommandName = "CreateNamedRanges"
        Case "BTNCREATENAMEDRANGES"
            ResolveCommandName = "CreateNamedRanges"
        Case "UI_HOTKEYS.HOTKEY_APPLYDEFAULTFORMAT"
            ResolveCommandName = "ApplyDefaultFormat"
        Case "UI_HOTKEYS.HOTKEY_APPLYCUSTOMNUMBERFORMAT"
            ResolveCommandName = "ApplyCustomNumberFormat"
        Case "UI_HOTKEYS.HOTKEY_MAKEPERMANENT"
            ResolveCommandName = "MakePermanent"
        Case "UI_HOTKEYS.HOTKEY_CREATENAMEDSHEET"
            ResolveCommandName = "CreateSheet"
        Case "UI_HOTKEYS.HOTKEY_FILLDOWN"
            ResolveCommandName = "FillDown"
        Case "UI_HOTKEYS.HOTKEY_FILLRIGHT"
            ResolveCommandName = "FillRight"
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
        Case "MODIFYDATA"
            Set CreateCommand = New FeatCmd_ModifyData
        Case "DATEFIXER"
            Set CreateCommand = New FeatCmd_ModifyData
        Case "CASEFIXER"
            Set CreateCommand = New FeatCmd_ModifyData
        Case "HIGHLIGHTDATA"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "HIGHLIGHTINCONSISTENTFORMULAS"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "HIGHLIGHTDUPLICATES"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "HIGHLIGHTERRORS"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "HIGHLIGHTHARDCODEDVALUES"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "HIGHLIGHTDATAVALIDATIONS"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "HIGHLIGHTCONDITIONALFORMATTING"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "HIGHLIGHTNAMEDRANGES"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "HIGHLIGHTFINANCIALMODELLING"
            Set CreateCommand = New FeatCmd_FinancialModelling
        Case "CLEARHIGHLIGHTS"
            Set CreateCommand = New FeatCmd_HighlightData
        Case "BREAKEXTERNALLINKS"
            Set CreateCommand = New FeatCmd_BreakExternalLinks
        Case "DUPLICATE"
            Set CreateCommand = New FeatCmd_Duplicate
        Case "EXPORTIMAGEORPDF"
            Set CreateCommand = New FeatCmd_ExportImageOrPdf
        Case "EXPORTPNG"
            Set CreateCommand = New FeatCmd_ExportImageOrPdf
        Case "EXPORTPDF"
            Set CreateCommand = New FeatCmd_ExportImageOrPdf
        Case "SHOWHELPCENTER"
            Set CreateCommand = New FeatCmd_ShowHelpCenter
        Case "HELLOWORLD"
            Set CreateCommand = New FeatCmd_HelloWorld
        Case "TABLEOFCONTENTS"
            Set CreateCommand = New FeatCmd_TableOfContents
        Case "UNMERGEFILL"
            Set CreateCommand = New FeatCmd_UnmergeFill
        Case "FORCENUMBER"
            Set CreateCommand = New FeatCmd_ForceNumber
        Case "CREATENAMEDRANGES"
            Set CreateCommand = New FeatCmd_CreateNamedRanges
        Case "APPLYDEFAULTFORMAT"
            Set CreateCommand = New FeatCmd_ApplyDefaultFormat
        Case "APPLYCUSTOMNUMBERFORMAT"
            Set CreateCommand = New FeatCmd_ApplyCustomNumberFormat
        Case "MAKEPERMANENT"
            Set CreateCommand = New FeatCmd_MakePermanent
        Case "CREATESHEET"
            Set CreateCommand = New FeatCmd_CreateSheet
        Case "FILLDOWN"
            Set CreateCommand = New FeatCmd_FillDown
        Case "FILLRIGHT"
            Set CreateCommand = New FeatCmd_FillRight
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