Attribute VB_Name = "Feat_MakeItStatic"
Option Explicit

' @Module: Feat_MakeItStatic
' @Category: Feature
' @Description: Compatibility wrappers that forward legacy entry points into the command pipeline.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: AppContainer, Infra_Error

' Prompts for scope (workbook vs sheet), then converts all formulas to values.
Public Sub StaticSheetWorkbook()
    Dim tracker As Object: Set tracker = Infra_Error.Track("StaticSheetWorkbook")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "StaticSheetWorkbook", "Feat_MakeItStatic.StaticSheetWorkbook", "Compatibility"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "StaticSheetWorkbook", Err
    Resume CleanExit
End Sub

' Converts the selected range's formulas to static values while preserving
' formatting. For small selections this is faster than sheet-level scanning.
Public Sub MakePermanent()
    Dim tracker As Object: Set tracker = Infra_Error.Track("MakePermanent")
    On Error GoTo ErrHandler

    AppContainer.ExecuteCommand "MakePermanent", "Feat_MakeItStatic.MakePermanent", "Compatibility"

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "MakePermanent", Err
    Resume CleanExit
End Sub
