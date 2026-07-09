Attribute VB_Name = "Lib_HelpManifest"
Option Explicit
Option Private Module

' @Module: Lib_HelpManifest
' @Category: Library
' @Description: Generated help manifest containing ribbon feature descriptions for dynamic help display.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

Public Function GetFeatureHelp() As Collection
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetFeatureHelp")
    On Error GoTo ErrHandler
    
    Dim col As New Collection
    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Wrap"
    dict.Add "Screentip", "Wrap Selection"
    dict.Add "Supertip", "Choose whether to wrap the selection by reusing a wrapper formula from another cell or by typing a formula pattern with [value]."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Static Sheet/WB"
    dict.Add "Screentip", "Convert Formulas to Static Values"
    dict.Add "Supertip", "Permanently replaces all formulas with their current values across the active sheet or the entire workbook. Handles spill ranges correctly."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Clean Data"
    dict.Add "Screentip", "Clean Data"
    dict.Add "Supertip", "Trims spaces and removes non-printable characters for the selected scope (Range, Sheet, or Workbook)."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Modify Data"
    dict.Add "Screentip", "Modify Data"
    dict.Add "Supertip", "Modifies text and dates in the selection (casing or date standardization)."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Date Fixer"
    dict.Add "Screentip", "Date Fixer"
    dict.Add "Supertip", "Standardizes dates in the selection."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Case Fixer"
    dict.Add "Screentip", "Case Fixer"
    dict.Add "Supertip", "Standardizes casing of text in the selection."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Highlight Data"
    dict.Add "Screentip", "Highlight Data"
    dict.Add "Supertip", "Highlights key data patterns and duplicates."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Inconsistent Formulas"
    dict.Add "Screentip", "Inconsistent Formulas"
    dict.Add "Supertip", "Highlights inconsistent formulas in yellow."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Duplicates"
    dict.Add "Screentip", "Duplicates"
    dict.Add "Supertip", "Highlights duplicate values in soft red."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Errors"
    dict.Add "Screentip", "Errors"
    dict.Add "Supertip", "Highlights cells with errors in orange."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Hardcoded Values"
    dict.Add "Screentip", "Hardcoded Values"
    dict.Add "Supertip", "Highlights formulas containing hardcoded values in lavender."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Data Validations"
    dict.Add "Screentip", "Data Validations"
    dict.Add "Supertip", "Highlights cells with data validation rules in soft green."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Conditional Formatting"
    dict.Add "Screentip", "Conditional Formatting"
    dict.Add "Supertip", "Highlights cells with conditional formatting in soft blue."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Financial Modelling"
    dict.Add "Screentip", "Financial Modelling"
    dict.Add "Supertip", "Formats cells using the standard financial model coloring scheme."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Clear Highlights"
    dict.Add "Screentip", "Clear Highlights"
    dict.Add "Supertip", "Clears all cells in the workbook matching the highlight color."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Break Links"
    dict.Add "Screentip", "Break External Links"
    dict.Add "Supertip", "Removes all external workbook references and keeps current values."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Duplicate Workbook"
    dict.Add "Screentip", "Duplicate Active Workbook"
    dict.Add "Supertip", "Duplicates the active workbook to your Desktop, opens it, and closes the original so you can work safely on the copy."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Export"
    dict.Add "Screentip", "Export"
    dict.Add "Supertip", "Exports the selected range as a PNG image or a PDF document."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Export as PNG"
    dict.Add "Screentip", "Export as PNG"
    dict.Add "Supertip", "Exports the selected range as a PNG image."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Export as PDF"
    dict.Add "Screentip", "Export as PDF"
    dict.Add "Supertip", "Exports the selected range as a PDF document."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Help Center"
    dict.Add "Screentip", "Show Help Center"
    dict.Add "Supertip", "Opens the Beaver Help Center, listing all registered hotkeys, user-defined functions, and their descriptions."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Hello World"
    dict.Add "Screentip", "Say Hello World"
    dict.Add "Supertip", "Puts 'Hello world' into the active cell (metadata test)."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Report Generator"
    dict.Add "Screentip", "Report Generator"
    dict.Add "Supertip", "Generate structured Table of Contents."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Table of Contents"
    dict.Add "Screentip", "Table of Contents"
    dict.Add "Supertip", "Creates a hyperlinked index sheet at the beginning of the workbook."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Unmerge FillDown"
    dict.Add "Screentip", "Unmerge FillDown"
    dict.Add "Supertip", "Unmerges the selected cells and fills the parent value to all unmerged cells."
    col.Add dict

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Label", "Convert Text to Number"
    dict.Add "Screentip", "Convert Text to Number"
    dict.Add "Supertip", "Forces text-formatted numbers in the selection to become actual numeric values."
    col.Add dict

    Set GetFeatureHelp = col

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "GetFeatureHelp", Err
    Resume CleanExit
End Function