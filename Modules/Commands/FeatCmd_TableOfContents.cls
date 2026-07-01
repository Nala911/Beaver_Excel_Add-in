VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_TableOfContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @Module: FeatCmd_TableOfContents
' @Category: Feature
' @Description: Generates a hyperlinked Table of Contents sheet in the active workbook.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ICommandContext, Infra_Error, Infra_CommandSupport, Infra_Config, Infra_Interaction

Implements ICommand

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("TableOfContents.Validate")
    On Error GoTo ErrHandler

    Dim result As CommandValidationResult
    Set result = Infra_CommandSupport.ValidateHasWorkbook(context)
    If Not result.IsExecutable Then
        Set ICommand_Validate = result
        Exit Function
    End If

    Set result = Infra_CommandSupport.ValidateActiveWorkbookNotAddin(context, "Table of Contents")
    If Not result.IsExecutable Then
        Set ICommand_Validate = result
        Exit Function
    End If

    Set result = Infra_CommandSupport.ValidateWorkbookNotProtected(context)
    If Not result.IsExecutable Then
        ' Override error message slightly if needed to match original behavior, or use standard
        result.Message = "The workbook structure is protected. Cannot add or overwrite the Table of Contents sheet."
        Set ICommand_Validate = result
        Exit Function
    End If

    Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "TableOfContents.Validate", Err
    Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("Failed to validate Table of Contents command.")
    Resume CleanExit
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("TableOfContents.Execute")
    On Error GoTo ErrHandler

    Dim ctx As ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorkbookRef Is Nothing Then GoTo CleanExit

    Dim wb As Workbook: Set wb = ctx.WorkbookRef

    ' Check if "Table of Contents" already exists
    Dim wsTOC As Worksheet
    On Error Resume Next
    Set wsTOC = wb.Worksheets("Table of Contents")
    On Error GoTo ErrHandler

    If Not wsTOC Is Nothing Then
        If Not Infra_Interaction.Confirm( _
            "A 'Table of Contents' worksheet already exists. Overwriting it will clear its current content and clear Excel's Undo stack. Do you want to proceed?", _
            "Table of Contents Generator") Then
            GoTo CleanExit
        End If

        ' Delete the existing sheet
        Application.DisplayAlerts = False
        wsTOC.Delete
        Application.DisplayAlerts = True
        Set wsTOC = Nothing
    End If

    ' Create new sheet at the beginning
    Set wsTOC = wb.Worksheets.Add(Before:=wb.Sheets(1))
    wsTOC.Name = "Table of Contents"

    ' Enable gridlines (requires sheet to be active in the window)
    wsTOC.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo ErrHandler

    ' Standard styling
    Dim fontName As String: fontName = Infra_Config.DEFAULT_FONT_NAME
    Dim defaultSize As Long: defaultSize = Infra_Config.DEFAULT_FONT_SIZE
    Dim headerSize As Long: headerSize = Infra_Config.HEADER_FONT_SIZE
    Dim headerColor As Long: headerColor = Infra_Config.HEADER_COLOR

    ' Calculate sheet statistics in target scope
    Dim visibleCount As Long
    Dim hiddenCount As Long
    Dim veryHiddenCount As Long
    Dim totalSheetCount As Long
    Dim ws As Worksheet
    Dim wsItem As Variant
    
    For Each wsItem In wb.Worksheets
        Set ws = wsItem
        If ws.Name <> "Table of Contents" And _
           ws.Name <> "_BeaverUndo" And _
           ws.Name <> "_BeaverTests" And _
           Left$(ws.Name, 10) <> "Test_Temp_" And _
           Left$(ws.Name, 1) <> "_" Then
            totalSheetCount = totalSheetCount + 1
            Select Case ws.Visible
                Case xlSheetVisible
                    visibleCount = visibleCount + 1
                Case xlSheetHidden
                    hiddenCount = hiddenCount + 1
                Case xlSheetVeryHidden
                    veryHiddenCount = veryHiddenCount + 1
            End Select
        End If
    Next wsItem

    ' Format Title Block & Workbook Summary Card (Rows 2 & 3)
    wsTOC.Range("B2:F3").Interior.Color = RGB(248, 250, 252) ' Light slate background
    wsTOC.Range("B2:F3").Borders.LineStyle = xlContinuous
    wsTOC.Range("B2:F3").Borders.Color = RGB(226, 232, 240) ' Soft border
    
    wsTOC.Cells(2, 2).Value = "Workbook Table of Contents"
    With wsTOC.Cells(2, 2).Font
        .Name = fontName
        .Size = 14
        .Bold = True
        .Color = RGB(30, 41, 59)
    End With

    wsTOC.Cells(3, 2).Value = "Generated on " & Format$(Now, "yyyy-mm-dd hh:nn") & "  |  Sheets: " & visibleCount & " Visible, " & hiddenCount & " Hidden, " & veryHiddenCount & " Very Hidden (Total: " & totalSheetCount & ")"
    With wsTOC.Cells(3, 2).Font
        .Name = fontName
        .Size = 9
        .Italic = True
        .Color = RGB(100, 116, 139)
    End With

    ' Format Headers
    Dim headers As Variant
    headers = Array("#", "Sheet Name", "Visibility", "Populated Cells", "Description / Notes")

    Dim colIndex As Long
    For colIndex = 0 To UBound(headers)
        wsTOC.Cells(5, 2 + colIndex).Value = headers(colIndex)
    Next colIndex

    Dim headerRange As Range
    Set headerRange = wsTOC.Range(wsTOC.Cells(5, 2), wsTOC.Cells(5, 2 + UBound(headers)))
    With headerRange
        .Font.Name = fontName
        .Font.Size = headerSize
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) ' White text on dark header color
        .Interior.Color = headerColor
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(226, 232, 240)
    End With

    ' Write Sheet Info
    Dim rowIndex As Long: rowIndex = 6
    Dim count As Long: count = 1

    For Each wsItem In wb.Worksheets
        Set ws = wsItem
        If ws.Name <> "Table of Contents" And _
           ws.Name <> "_BeaverUndo" And _
           ws.Name <> "_BeaverTests" And _
           Left$(ws.Name, 10) <> "Test_Temp_" And _
           Left$(ws.Name, 1) <> "_" Then
           
            ' #
            wsTOC.Cells(rowIndex, 2).Value = count
            wsTOC.Cells(rowIndex, 2).Font.Name = fontName
            wsTOC.Cells(rowIndex, 2).Font.Size = defaultSize
            wsTOC.Cells(rowIndex, 2).Font.Color = RGB(100, 116, 139)

            ' Sheet Name (Hyperlink if visible)
            wsTOC.Cells(rowIndex, 3).Font.Name = fontName
            wsTOC.Cells(rowIndex, 3).Font.Size = defaultSize
            If ws.Visible = xlSheetVisible Then
                wsTOC.Hyperlinks.Add Anchor:=wsTOC.Cells(rowIndex, 3), _
                                      Address:="", _
                                      SubAddress:="'" & Replace(ws.Name, "'", "''") & "'!A1", _
                                      TextToDisplay:=ws.Name
            Else
                wsTOC.Cells(rowIndex, 3).Value = ws.Name
                wsTOC.Cells(rowIndex, 3).Font.Italic = True
                wsTOC.Cells(rowIndex, 3).Font.Color = RGB(148, 163, 184)
            End If

            ' Visibility (Clean styled indicators)
            Dim visibilityStr As String
            Dim visColor As Long
            Select Case ws.Visible
                Case xlSheetVisible
                    visibilityStr = "Visible"
                    visColor = RGB(22, 163, 74) ' Calm Green
                Case xlSheetHidden
                    visibilityStr = "Hidden"
                    visColor = RGB(217, 119, 6) ' Amber
                Case xlSheetVeryHidden
                    visibilityStr = "Very Hidden"
                    visColor = RGB(220, 38, 38) ' Crimson
                Case Else
                    visibilityStr = "Unknown"
                    visColor = RGB(100, 116, 139)
            End Select
            
            wsTOC.Cells(rowIndex, 4).Value = visibilityStr
            wsTOC.Cells(rowIndex, 4).Font.Name = fontName
            wsTOC.Cells(rowIndex, 4).Font.Size = defaultSize
            wsTOC.Cells(rowIndex, 4).Font.Color = visColor
            wsTOC.Cells(rowIndex, 4).Font.Bold = (ws.Visible <> xlSheetVisible)

            ' Populated Cells
            Dim cellCount As Long
            cellCount = 0
            On Error Resume Next
            cellCount = Application.WorksheetFunction.CountA(ws.UsedRange)
            On Error GoTo ErrHandler
            wsTOC.Cells(rowIndex, 5).Value = cellCount
            wsTOC.Cells(rowIndex, 5).Font.Name = fontName
            wsTOC.Cells(rowIndex, 5).Font.Size = defaultSize
            wsTOC.Cells(rowIndex, 5).NumberFormat = Infra_Config.DEFAULT_NUMBER_FORMAT
            wsTOC.Cells(rowIndex, 5).Font.Color = RGB(71, 85, 105)

            ' Notes (Blank)
            wsTOC.Cells(rowIndex, 6).Font.Name = fontName
            wsTOC.Cells(rowIndex, 6).Font.Size = defaultSize

            ' Add borders to this row
            With wsTOC.Range(wsTOC.Cells(rowIndex, 2), wsTOC.Cells(rowIndex, 6)).Borders
                .LineStyle = xlContinuous
                .Color = RGB(226, 232, 240)
            End With

            rowIndex = rowIndex + 1
            count = count + 1
        End If
    Next wsItem

    ' Autofit columns with safety cap
    Dim targetCol As Range
    For colIndex = 2 To 6
        Set targetCol = wsTOC.Columns(colIndex)
        targetCol.EntireColumn.AutoFit
        If targetCol.ColumnWidth > Infra_Config.COLUMN_WIDTH_THRESHOLD Then
            targetCol.ColumnWidth = Infra_Config.MAX_COLUMN_WIDTH
        End If
    Next colIndex

    ' Set selection on first data cell
    wsTOC.Cells(6, 3).Select

    Infra_Interaction.ShowInfo "Table of Contents generated successfully!", "Table of Contents Generator"

CleanExit:
    Exit Sub
ErrHandler:
    Infra_Error.HandleError "TableOfContents.Execute", Err
    Resume CleanExit
End Sub
