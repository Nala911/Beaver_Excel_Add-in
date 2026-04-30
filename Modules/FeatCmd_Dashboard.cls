VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: FeatCmd_Dashboard
' @Category: Feature
' @Description: Generates a minimal financial dashboard.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ICommandContext, AppContainer, Infra_Error, Infra_AppStateGuard, Infra_Undo

Implements ICommand

Private Sub ICommand_Execute(ByVal Context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("FeatCmd_Dashboard.Execute")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Dim targetWorkbook As Workbook
    Dim ws As Worksheet

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(Context)
    If ctx Is Nothing Then GoTo CleanExit
    Set targetWorkbook = ctx.WorkbookRef
    If targetWorkbook Is Nothing Then GoTo CleanExit

    Set ws = targetWorkbook.Worksheets.Add
    ws.Name = "Dashboard_" & Format(Now, "hhmmss")
    
    ' Minimum effort financial dashboard
    ws.Range("B2").Value = "Minimal Financial Dashboard"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 14
    
    ws.Range("B4").Value = "Metric"
    ws.Range("C4").Value = "Value"
    
    ws.Range("B5").Value = "Revenue"
    ws.Range("C5").Value = 100000
    
    ws.Range("B6").Value = "Expenses"
    ws.Range("C6").Value = 45000
    
    ws.Range("B7").Value = "Net Profit"
    ws.Range("C7").Formula2 = "=C5-C6"
    
    ' Simple formatting
    ws.Range("B4:C4").Interior.ThemeColor = xlThemeColorAccent1
    ws.Range("B4:C4").Font.ThemeColor = xlThemeColorDark1
    ws.Range("B4:C4").Font.Bold = True
    
    ws.Range("C5:C7").NumberFormat = "$#,##0"
    ws.Columns("B:C").AutoFit
    
CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "FeatCmd_Dashboard.Execute", Err
    Resume CleanExit
End Sub

Private Function ICommand_CanExecute(ByVal context As ICommandContext) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("FeatCmd_Dashboard.CanExecute")
    On Error GoTo ErrHandler

    Dim ctx As Infra_ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    ICommand_CanExecute = Not ctx Is Nothing
    If ICommand_CanExecute Then ICommand_CanExecute = Not ctx.WorkbookRef Is Nothing

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "FeatCmd_Dashboard.CanExecute", Err
    ICommand_CanExecute = False
    Resume CleanExit
End Function
