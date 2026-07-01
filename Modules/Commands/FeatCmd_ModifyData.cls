VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_ModifyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand
Implements IBatchProcessor
Implements ICellTransformer

' @Module: FeatCmd_ModifyData
' @Category: Feature
' @Description: Modification tools for selected ranges (Casing: UPPER/lower/Proper, Date Standardization using custom patterns).
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ActionContext, ModifyDataRequest, Infra_Error, Infra_Undo, Infra_Progress, Infra_Interaction, Infra_Config

Private m_CurrentRequest As ModifyDataRequest

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Dim tracker As Object: Set tracker = Infra_Error.Track("ModifyData.Validate")
    On Error GoTo ErrHandler

    Dim result As CommandValidationResult
    Set result = Infra_CommandSupport.ValidateHasWorksheet(context)
    If Not result.IsExecutable Then
        Set ICommand_Validate = result
        Exit Function
    End If

    Set result = Infra_CommandSupport.ValidateWorksheetNotProtected(context, "Cannot modify data on a protected worksheet. Please unprotect the sheet and try again.")
    If Not result.IsExecutable Then
        Set ICommand_Validate = result
        Exit Function
    End If

    Set ICommand_Validate = Infra_CommandSupport.ValidationSuccess()

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "ModifyData.Validate", Err
    Set ICommand_Validate = Infra_CommandSupport.ValidationFailure("The data cannot be modified right now.")
    Resume CleanExit
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ModifyData.Execute")
    On Error GoTo ErrHandler
    
    Dim ctx As ActionContext
    Dim request As ModifyDataRequest
    Dim stateInfo As Collection

    Set stateInfo = Infra_Error.NewStateInfo()

    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Or ctx.WorkbookRef Is Nothing Then GoTo CleanExit
    
    Infra_Error.AddState stateInfo, "Workbook", ctx.WorkbookRef.Name
    Infra_Error.AddState stateInfo, "Worksheet", ctx.WorksheetRef.Name
    If ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing Then
        Infra_Error.AddState stateInfo, "Selection", ctx.SelectionRange.Address(False, False)
    End If

    Set request = UI_Factory.ShowModifyDataDialog(ctx, context.CommandName)
    If request Is Nothing Then GoTo CleanExit
    
    ExecuteModifyData request

CleanExit:
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "ModifyData.Execute", Err, stateInfo
    Resume CleanExit
End Sub

Private Sub ExecuteModifyData(ByVal request As ModifyDataRequest)
    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteModifyData")
    On Error GoTo ErrHandler
    
    If request Is Nothing Then GoTo CleanExit
    If request.Context Is Nothing Then GoTo CleanExit
    
    Set m_CurrentRequest = request
    
    Dim totalModified As Long
    totalModified = Infra_BatchProcessor.Execute(request.Context, TargetScopeSelection, Me, "Modify Data", False, UndoCaptureFormulaOnly)
    
    Dim msg As String
    msg = "Done! Modified " & totalModified & " cells in the selected range."
    Infra_Interaction.ShowInfo msg, "Modify Data"

CleanExit:
    Set m_CurrentRequest = Nothing
    Exit Sub

ErrHandler:
    Set m_CurrentRequest = Nothing
    Infra_Error.HandleError "ExecuteModifyData", Err
    Resume CleanExit
End Sub

Private Function IBatchProcessor_ProcessRange(ByVal targetRange As Range) As Long
    IBatchProcessor_ProcessRange = ProcessRangeForModify(targetRange, m_CurrentRequest)
End Function

Private Function IBatchProcessor_ProcessWorksheet(ByVal ws As Worksheet) As Long
    IBatchProcessor_ProcessWorksheet = 0
End Function

Private Function ICellTransformer_TransformCell(ByVal rawValue As Variant, ByVal currentFormat As String, ByRef outValue As Variant, ByRef outFormat As String) As Boolean
    If m_CurrentRequest Is Nothing Then
        ICellTransformer_TransformCell = False
        Exit Function
    End If
    ICellTransformer_TransformCell = ModifyTextValue(rawValue, currentFormat, m_CurrentRequest, outValue, outFormat)
End Function

Private Function ProcessRangeForModify(ByRef targetRange As Range, ByVal request As ModifyDataRequest) As Long
    If targetRange Is Nothing Then Exit Function
    
    ' Temporarily set current request to support direct/test calls
    Dim prevRequest As ModifyDataRequest
    Set prevRequest = m_CurrentRequest
    Set m_CurrentRequest = request
    
    ProcessRangeForModify = Infra_CommandSupport.ProcessRangeUnified(targetRange, Me, xlTextValues Or xlNumbers)
    
    ' Restore previous request
    Set m_CurrentRequest = prevRequest
End Function

Private Function ModifyTextValue( _
    ByVal rawValue As Variant, _
    ByVal currentFormat As String, _
    ByVal request As ModifyDataRequest, _
    ByRef outValue As Variant, _
    ByRef outFormat As String) As Boolean
    
    If IsError(rawValue) Then
        ModifyTextValue = False
        Exit Function
    End If
    
    Dim textValue As String
    Dim success As Boolean
    
    If VarType(rawValue) = vbDate Then
        If request.Operation = "Date Standardization" Then
            ' If the cell already contains a date, we try two approaches to reconstruct
            ' the original date string components for re-parsing:
            ' 1. Reconstruct using the cell's current format (if it contains date formatting characters)
            Dim isDateFormatted As Boolean
            isDateFormatted = False
            If currentFormat <> vbNullString And LCase$(currentFormat) <> "general" Then
                If InStr(1, currentFormat, "d", vbTextCompare) > 0 Or _
                   InStr(1, currentFormat, "m", vbTextCompare) > 0 Or _
                   InStr(1, currentFormat, "y", vbTextCompare) > 0 Then
                    isDateFormatted = True
                End If
            End If
            
            success = False
            If isDateFormatted Then
                On Error Resume Next
                textValue = Format$(rawValue, currentFormat)
                If Err.Number = 0 And Trim$(textValue) <> vbNullString Then
                    Dim tempDate As Date
                    If ParseCustomDate(textValue, request.DatePattern, tempDate) Then
                        If tempDate <> rawValue Then
                            outValue = tempDate
                            outFormat = Infra_Config.Model.DisplayDateFormat
                            ModifyTextValue = True
                            success = True
                        End If
                    End If
                End If
                Err.Clear
                On Error GoTo 0
            End If
            
            ' 2. Reconstruct using the system's local date format order
            If Not success Then
                Dim sysPattern As String
                sysPattern = Infra_ValueConversion.GetSystemDateFormatPattern()
                On Error Resume Next
                textValue = Format$(rawValue, sysPattern)
                If Err.Number = 0 And Trim$(textValue) <> vbNullString Then
                    Dim tempDate2 As Date
                    If ParseCustomDate(textValue, request.DatePattern, tempDate2) Then
                        If tempDate2 <> rawValue Then
                            outValue = tempDate2
                            outFormat = Infra_Config.Model.DisplayDateFormat
                            ModifyTextValue = True
                            success = True
                        End If
                    End If
                End If
                Err.Clear
                On Error GoTo 0
            End If
            
            If success Then
                Exit Function
            End If
            
            ' If no re-parsing succeeded in changing the date value, do not modify
            ModifyTextValue = False
            Exit Function
        Else
            ' Non-date operation on a Date value (e.g., Casing)
            If currentFormat <> vbNullString And LCase$(currentFormat) <> "general" Then
                textValue = Format$(rawValue, currentFormat)
            Else
                textValue = CStr(rawValue)
            End If
        End If
    Else
        textValue = CStr(rawValue)
    End If
    
    If Trim$(textValue) = vbNullString Then
        ModifyTextValue = False
        Exit Function
    End If
    
    Select Case request.Operation
        Case "Date Standardization"
            Dim d As Date
            If ParseCustomDate(textValue, request.DatePattern, d) Then
                outValue = d
                outFormat = Infra_Config.Model.DisplayDateFormat
                ModifyTextValue = True
                Exit Function
            End If
            
        Case "Case: UPPERCASE"
            Dim upperVal As String
            upperVal = UCase$(textValue)
            If upperVal <> textValue Then
                outValue = upperVal
                outFormat = currentFormat
                ModifyTextValue = True
                Exit Function
            End If
            
        Case "Case: lowercase"
            Dim lowerVal As String
            lowerVal = LCase$(textValue)
            If lowerVal <> textValue Then
                outValue = lowerVal
                outFormat = currentFormat
                ModifyTextValue = True
                Exit Function
            End If
            
        Case "Case: Proper Case"
            Dim properVal As String
            properVal = StrConv(textValue, vbProperCase)
            If properVal <> textValue Then
                outValue = properVal
                outFormat = currentFormat
                ModifyTextValue = True
                Exit Function
            End If
    End Select
    
    ModifyTextValue = False
End Function

' --- Date Parsing Helpers ---

Private Function NormalizeDelimiters(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "-", "/")
    result = Replace(result, ".", "/")
    result = Replace(result, ",", "/")
    result = Replace(result, "\", "/")
    result = Replace(result, " ", "/")
    NormalizeDelimiters = result
End Function

Private Function ParseDay(ByVal part As String) As Integer
    If IsNumeric(part) Then
        ParseDay = CInt(part)
    Else
        ParseDay = 0
    End If
End Function

Private Function ParseMonth(ByVal part As String) As Integer
    If IsNumeric(part) Then
        ParseMonth = CInt(part)
    Else
        Select Case LCase$(Left$(Trim$(part), 3))
            Case "jan": ParseMonth = 1
            Case "feb": ParseMonth = 2
            Case "mar": ParseMonth = 3
            Case "apr": ParseMonth = 4
            Case "may": ParseMonth = 5
            Case "jun": ParseMonth = 6
            Case "jul": ParseMonth = 7
            Case "aug": ParseMonth = 8
            Case "sep": ParseMonth = 9
            Case "oct": ParseMonth = 10
            Case "nov": ParseMonth = 11
            Case "dec": ParseMonth = 12
            Case Else: ParseMonth = 0
        End Select
    End If
End Function

Private Function ParseYear(ByVal part As String) As Integer
    If IsNumeric(part) Then
        Dim yr As Integer
        yr = CInt(part)
        If yr < 100 Then
            If yr <= 30 Then
                ParseYear = 2000 + yr
            Else
                ParseYear = 1900 + yr
            End If
        Else
            ParseYear = yr
        End If
    Else
        ParseYear = 0
    End If
End Function

Private Function ParseNonDelimitedDate(ByVal text As String, ByVal pattern As String, ByRef outDate As Date) As Boolean
    Dim cleanText As String, cleanPattern As String
    cleanText = Trim$(text)
    cleanPattern = Trim$(pattern)
    
    If Len(cleanText) <> Len(cleanPattern) Then
        ParseNonDelimitedDate = False
        Exit Function
    End If
    
    Dim i As Long
    Dim dyStr As String, mthStr As String, yrStr As String
    Dim char As String
    
    For i = 1 To Len(cleanPattern)
        char = LCase$(Mid$(cleanPattern, i, 1))
        If char = "d" Then
            dyStr = dyStr & Mid$(cleanText, i, 1)
        ElseIf char = "m" Then
            mthStr = mthStr & Mid$(cleanText, i, 1)
        ElseIf char = "y" Then
            yrStr = yrStr & Mid$(cleanText, i, 1)
        End If
    Next i
    
    If Len(dyStr) = 0 Or Len(mthStr) = 0 Or Len(yrStr) = 0 Then
        ParseNonDelimitedDate = False
        Exit Function
    End If
    
    Dim dy As Integer, mth As Integer, yr As Integer
    dy = ParseDay(dyStr)
    mth = ParseMonth(mthStr)
    yr = ParseYear(yrStr)
    
    If dy = 0 Or mth = 0 Or yr = 0 Then
        ParseNonDelimitedDate = False
        Exit Function
    End If
    
    On Error Resume Next
    Dim tempDate As Date
    tempDate = DateSerial(yr, mth, dy)
    If Err.Number = 0 Then
        If Year(tempDate) = yr And Month(tempDate) = mth And Day(tempDate) = dy Then
            outDate = tempDate
            ParseNonDelimitedDate = True
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0
    
    ParseNonDelimitedDate = False
End Function

Private Function ParseCustomDate(ByVal text As String, ByVal pattern As String, ByRef outDate As Date) As Boolean
    Dim normPattern As String
    normPattern = NormalizeDelimiters(pattern)
    
    Dim patternHasDelimiters As Boolean
    patternHasDelimiters = (InStr(1, normPattern, "/") > 0)
    
    If patternHasDelimiters Then
        Dim normText As String
        normText = NormalizeDelimiters(text)
        
        Dim textParts() As String
        Dim patternParts() As String
        textParts = Split(normText, "/")
        patternParts = Split(normPattern, "/")
        
        If UBound(textParts) <> UBound(patternParts) Then
            ParseCustomDate = False
            Exit Function
        End If
        
        Dim i As Long
        Dim dy As Integer, mth As Integer, yr As Integer
        dy = 0: mth = 0: yr = 0
        
        For i = LBound(patternParts) To UBound(patternParts)
            Dim pPart As String
            pPart = LCase$(Trim$(patternParts(i)))
            
            If InStr(1, pPart, "d") > 0 Then
                dy = ParseDay(textParts(i))
            ElseIf InStr(1, pPart, "m") > 0 Then
                mth = ParseMonth(textParts(i))
            ElseIf InStr(1, pPart, "y") > 0 Then
                yr = ParseYear(textParts(i))
            End If
        Next i
        
        If dy = 0 Or mth = 0 Or yr = 0 Then
            ParseCustomDate = False
            Exit Function
        End If
        
        On Error Resume Next
        Dim tempDate As Date
        tempDate = DateSerial(yr, mth, dy)
        If Err.Number = 0 Then
            If Year(tempDate) = yr And Month(tempDate) = mth And Day(tempDate) = dy Then
                outDate = tempDate
                ParseCustomDate = True
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Else
        ' No delimiters in pattern. Remove delimiters from text and parse.
        Dim cleanText As String
        cleanText = Replace(text, "/", "")
        cleanText = Replace(cleanText, "-", "")
        cleanText = Replace(cleanText, ".", "")
        cleanText = Replace(cleanText, ",", "")
        cleanText = Replace(cleanText, " ", "")
        
        ParseCustomDate = ParseNonDelimitedDate(cleanText, pattern, outDate)
        Exit Function
    End If
    
    ParseCustomDate = False
End Function



' --- Headless testing support ---

Public Function ModifyRangeWithOptionsDirect(ByVal targetRange As Range, ByVal request As ModifyDataRequest) As Long
    Dim tracker As Object: Set tracker = Infra_Error.Track("ModifyRangeWithOptionsDirect")
    On Error GoTo ErrHandler
    
    ModifyRangeWithOptionsDirect = ProcessRangeForModify(targetRange, request)

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "ModifyRangeWithOptionsDirect", Err
    Resume CleanExit
End Function
