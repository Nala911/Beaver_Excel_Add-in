VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeatCmd_CreateNamedRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

' @Module: FeatCmd_CreateNamedRanges
' @Category: Feature
' @Description: Creates Excel named ranges in bulk from non-adjacent value and label ranges. Automatically sanitizes labels.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: ICommand, ActionContext, Infra_Error, Infra_Undo, Infra_Interaction, Infra_CommandSupport, Infra_AppStateGuard

Private Property Get ICommand_ExecutionPolicy() As CommandExecutionPolicy
    Set ICommand_ExecutionPolicy = Infra_CommandSupport.PolicyInteractiveUI()
End Property

Private Function ICommand_Validate(ByVal context As ICommandContext) As CommandValidationResult
    Set ICommand_Validate = Infra_CommandSupport.ValidateHasWorksheet(context)
End Function

Private Sub ICommand_Execute(ByVal context As ICommandContext)
    Dim tracker As Object: Set tracker = Infra_Error.Track("FeatCmd_CreateNamedRanges.Execute")
    Dim stateInfo As Collection: Set stateInfo = Infra_Error.NewStateInfo()
    On Error GoTo ErrHandler
    
    Dim ctx As ActionContext
    Set ctx = Infra_CommandSupport.ActionContextFromCommandContext(context)
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Or ctx.WorkbookRef Is Nothing Then GoTo CleanExit
    
    Dim targetWb As Workbook: Set targetWb = ctx.WorkbookRef
    Dim targetWs As Worksheet: Set targetWs = ctx.WorksheetRef
    
    Infra_Error.AddState stateInfo, "Workbook", targetWb.Name
    Infra_Error.AddState stateInfo, "Worksheet", targetWs.Name
    
    ' Determine Value Range
    Dim valueRange As Range
    Dim skipValuePrompt As Boolean: skipValuePrompt = False
    
    If ctx.HasRangeSelection And Not ctx.SelectionRange Is Nothing Then
        If ctx.SelectionRange.Cells.CountLarge > 1 Then
            Set valueRange = ctx.SelectionRange
            skipValuePrompt = True
        End If
    End If
    
    If Not skipValuePrompt Then
        ' Prompt 1: Select Value Range
        If Not Infra_Interaction.PromptRange("Select the Value Range (cells containing numbers/formulas to name):", "Select Value Range", valueRange) Then
            GoTo CleanExit
        End If
    End If
    
    ' Prompt 2: Select Label Range
    Dim labelRange As Range
    Dim expectedCount As Long: expectedCount = valueRange.Cells.CountLarge
    Dim labelPrompt As String
    labelPrompt = "Select the Label Range containing exactly " & expectedCount & " cell"
    If expectedCount <> 1 Then labelPrompt = labelPrompt & "s"
    labelPrompt = labelPrompt & ":"
    
    If Not Infra_Interaction.PromptRange(labelPrompt, "Select Label Range", labelRange) Then
        GoTo CleanExit
    End If
    
    ' Validation: Cells count must match
    If valueRange.Cells.CountLarge <> labelRange.Cells.CountLarge Then
        Infra_Interaction.ShowWarning "The Value Range has " & valueRange.Cells.CountLarge & " cells, but the Label Range has " & labelRange.Cells.CountLarge & " cells." & vbCrLf & vbCrLf & "Both ranges must contain the same number of cells.", "Bulk Named Ranges"
        GoTo CleanExit
    End If
    
    ' Default scope is always Workbook, no need to prompt
    Dim scopeChoice As String: scopeChoice = "Workbook"
    
    ' Run direct execution
    Dim createdCount As Long: createdCount = 0
    Dim createdNamesList As String: createdNamesList = ""
    Dim success As Boolean
    
    ' Set up AppStateGuard for speed during name creation
    Dim appGuard As Infra_AppStateGuard
    Set appGuard = New Infra_AppStateGuard
    
    ' Set overwriteAllParam (6th argument) to True so existing names are overwritten directly without asking
    success = ExecuteBulkNamedRangesDirect(valueRange, labelRange, scopeChoice, targetWb, targetWs, True, False, createdCount, createdNamesList)
    
    ' Release guard to show final messages/confirmations
    Set appGuard = Nothing
    
    If Not success Then GoTo CleanExit
    
    ' Register Undo if any names were created
    If createdNamesList <> "" Then
        Infra_Undo.SaveCreatedNamesState targetWb, createdNamesList
        Infra_Undo.RegisterPendingUndo
    End If
    
    ' Show Completion Msg
    Dim successMsg As String
    successMsg = "Done! Successfully created " & createdCount & " named ranges (" & scopeChoice & " scope)."
    If createdCount > 0 Then
        successMsg = successMsg & vbCrLf & vbCrLf & "You can undo this action by pressing Ctrl+Z."
    End If
    
    Infra_Interaction.ShowInfo successMsg, "Bulk Named Ranges"

CleanExit:
    Set appGuard = Nothing
    Exit Sub
ErrHandler:
    Set appGuard = Nothing
    Infra_Error.HandleError "FeatCmd_CreateNamedRanges.Execute", Err, stateInfo
    Resume CleanExit
End Sub

' Core loop implementation designed to be callable by UI flow and Unit Tests directly
Public Function ExecuteBulkNamedRangesDirect( _
    ByVal valueRange As Range, _
    ByVal labelRange As Range, _
    ByVal scopeChoice As String, _
    ByVal targetWb As Workbook, _
    ByVal targetWs As Worksheet, _
    ByVal overwriteAllParam As Boolean, _
    ByVal skipAllParam As Boolean, _
    ByRef outCreatedCount As Long, _
    ByRef outCreatedList As String) As Boolean

    Dim tracker As Object: Set tracker = Infra_Error.Track("ExecuteBulkNamedRangesDirect")
    On Error GoTo ErrHandler
    
    Dim overwriteAll As Boolean: overwriteAll = overwriteAllParam
    Dim skipAll As Boolean: skipAll = skipAllParam
    
    outCreatedCount = 0
    outCreatedList = ""
    
    Dim expectedCount As Long: expectedCount = valueRange.Cells.CountLarge
    
    ' Gather cells into flat collections to support transposed mapping sequentially
    Dim valueCells As New Collection
    Dim labelCells As New Collection
    Dim cellObj As Range
    
    For Each cellObj In valueRange.Cells
        valueCells.Add cellObj
    Next cellObj
    
    For Each cellObj In labelRange.Cells
        labelCells.Add cellObj
    Next cellObj
    
    ' Cache existing names for speed/optimization
    Dim dictNamesByName As Object
    Dim dictNamesByCell As Object
    Set dictNamesByName = CreateObject("Scripting.Dictionary")
    Set dictNamesByCell = CreateObject("Scripting.Dictionary")
    
    dictNamesByName.CompareMode = 1 ' TextCompare (case-insensitive)
    dictNamesByCell.CompareMode = 1
    
    Dim targetNames As Names
    If scopeChoice = "Workbook" Then
        Set targetNames = targetWb.Names
    Else
        Set targetNames = targetWs.Names
    End If
    
    Dim nameObj As Name
    For Each nameObj In targetNames
        Dim localName As String: localName = GetLocalName(nameObj)
        If Not dictNamesByName.Exists(localName) Then
            Set dictNamesByName(localName) = nameObj
        End If
        
        Dim refersRange As Range
        On Error Resume Next
        Set refersRange = nameObj.RefersToRange
        On Error GoTo ErrHandler
        
        If Not refersRange Is Nothing Then
            Dim cellAddr As String: cellAddr = refersRange.Address(External:=True)
            If Not dictNamesByCell.Exists(cellAddr) Then
                Dim nameCol As Collection
                Set nameCol = New Collection
                dictNamesByCell.Add cellAddr, nameCol
            End If
            dictNamesByCell(cellAddr).Add nameObj
        End If
    Next nameObj
    
    Dim i As Long
    For i = 1 To expectedCount
        Dim labelCell As Range: Set labelCell = labelCells(i)
        Dim valueCell As Range: Set valueCell = valueCells(i)
        
        Dim labelVal As Variant: labelVal = labelCell.Value
        If IsEmpty(labelVal) Or IsNull(labelVal) Or IsError(labelVal) Then
            GoTo NextCell
        End If
        
        Dim rawLabel As String: rawLabel = Trim$(CStr(labelVal))
        If rawLabel = "" Then GoTo NextCell
        
        Dim valueVal As Variant: valueVal = valueCell.Value
        If IsEmpty(valueVal) Or IsNull(valueVal) Then
            GoTo NextCell
        End If
        
        If Not IsError(valueVal) Then
            If Trim$(CStr(valueVal)) = "" Then
                GoTo NextCell
            End If
        End If
        
        Dim sanitizedName As String: sanitizedName = SanitizeExcelName(rawLabel)
        
        ' Check if name already exists
        Dim nameExists As Boolean: nameExists = dictNamesByName.Exists(sanitizedName)
        Dim existingName As Name
        
        If nameExists Then
            Set existingName = dictNamesByName(sanitizedName)
            
            ' Check if it points to the SAME cell
            If NameRefersToRange(existingName, valueCell) Then
                ' Redundant: Already named sanitizedName. Skip silently.
                GoTo NextCell
            Else
                ' Points to a DIFFERENT cell. This is a conflict!
                Dim existingRng As Range
                On Error Resume Next
                Set existingRng = existingName.RefersToRange
                On Error GoTo ErrHandler
                
                Dim existingCellRef As String: existingCellRef = ""
                If Not existingRng Is Nothing Then
                    existingCellRef = existingRng.Parent.Name & "!" & existingRng.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                Else
                    existingCellRef = "another formula/reference"
                End If
                
                Dim targetCellRef As String
                targetCellRef = valueCell.Parent.Name & "!" & valueCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                Infra_Interaction.ShowWarning "Cannot name cell " & targetCellRef & " as '" & sanitizedName & "' because '" & sanitizedName & "' already refers to cell " & existingCellRef & "." & vbCrLf & vbCrLf & "Both cells cannot have the same name.", "Bulk Named Ranges"
                GoTo NextCell
            End If
        End If
        
        ' If we reach here, we want to create sanitizedName for valueCell.
        ' Clean up other names pointing to this valueCell (renaming support)
        Dim currentAddr As String: currentAddr = valueCell.Address(External:=True)
        If dictNamesByCell.Exists(currentAddr) Then
            Dim cellNamesCol As Collection: Set cellNamesCol = dictNamesByCell(currentAddr)
            Dim j As Long
            ' Iterate backwards when deleting/modifying collections
            For j = cellNamesCol.Count To 1 Step -1
                Dim oldNameObj As Name: Set oldNameObj = cellNamesCol(j)
                Dim oldLocalName As String: oldLocalName = GetLocalName(oldNameObj)
                If oldLocalName <> sanitizedName Then
                    On Error Resume Next
                    oldNameObj.Delete
                    On Error GoTo ErrHandler
                    
                    ' Remove from caches
                    If dictNamesByName.Exists(oldLocalName) Then
                        dictNamesByName.Remove oldLocalName
                    End If
                    cellNamesCol.Remove j
                End If
            Next j
        End If
        
        ' Create Name
        Dim newNameObj As Name
        If scopeChoice = "Workbook" Then
            Set newNameObj = targetWb.Names.Add(Name:=sanitizedName, RefersTo:=valueCell)
        Else
            Set newNameObj = targetWs.Names.Add(Name:=sanitizedName, RefersTo:=valueCell)
        End If
        
        ' Update Cache
        If Not newNameObj Is Nothing Then
            Set dictNamesByName(sanitizedName) = newNameObj
            If Not dictNamesByCell.Exists(currentAddr) Then
                Dim newCol As Collection
                Set newCol = New Collection
                dictNamesByCell.Add currentAddr, newCol
            End If
            dictNamesByCell(currentAddr).Add newNameObj
            
            ' Accumulate for custom Undo
            If outCreatedList = "" Then
                outCreatedList = newNameObj.Name
            Else
                outCreatedList = outCreatedList & ";" & newNameObj.Name
            End If
            outCreatedCount = outCreatedCount + 1
        End If
        
NextCell:
    Next i
    
    ExecuteBulkNamedRangesDirect = True

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "FeatCmd_CreateNamedRanges.ExecuteBulkNamedRangesDirect", Err
    ExecuteBulkNamedRangesDirect = False
    Resume CleanExit
End Function

' Checks if a Name object refers exactly to the target Range.
Private Function NameRefersToRange(ByVal nameObj As Name, ByVal targetRange As Range) As Boolean
    Dim tracker As Object: Set tracker = Infra_Error.Track("NameRefersToRange")
    On Error GoTo ErrHandler
    
    NameRefersToRange = False
    
    Dim refersRange As Range
    On Error Resume Next
    Set refersRange = nameObj.RefersToRange
    On Error GoTo ErrHandler
    
    If refersRange Is Nothing Then Exit Function
    
    ' Compare sheet name and cell address
    If refersRange.Parent.Parent.Name = targetRange.Parent.Parent.Name Then
        If refersRange.Parent.Name = targetRange.Parent.Name Then
            If refersRange.Address = targetRange.Address Then
                NameRefersToRange = True
            End If
        End If
    End If
    
CleanExit:
    Exit Function
ErrHandler:
    ' Silently ignore errors (e.g. if the name refers to a formula/constant)
    Resume CleanExit
End Function

' Extracts the local name without sheet prefix (e.g. "Sheet1!Name" -> "Name")
Private Function GetLocalName(ByVal nameObj As Name) As String
    Dim fullName As String: fullName = nameObj.Name
    Dim pos As Long: pos = InStr(fullName, "!")
    If pos > 0 Then
        GetLocalName = Mid$(fullName, pos + 1)
    Else
        GetLocalName = fullName
    End If
End Function


' Sanitizes raw text into a valid Excel named range string
Public Function SanitizeExcelName(ByVal rawLabel As String) As String
    Dim tracker As Object: Set tracker = Infra_Error.Track("SanitizeExcelName")
    On Error GoTo ErrHandler

    Dim result As String
    result = Trim$(rawLabel)
    
    If result = vbNullString Then
        SanitizeExcelName = "_"
        GoTo CleanExit
    End If
    
    ' 1. Replace spaces, hyphens, slashes, and common special chars with underscores
    result = Replace(result, " ", "_")
    result = Replace(result, "-", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    result = Replace(result, "%", "_")
    result = Replace(result, "$", "_")
    result = Replace(result, "&", "_")
    result = Replace(result, "(", "_")
    result = Replace(result, ")", "_")
    result = Replace(result, "+", "_")
    result = Replace(result, "=", "_")
    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")
    result = Replace(result, "{", "_")
    result = Replace(result, "}", "_")
    result = Replace(result, ",", "_")
    result = Replace(result, ";", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "'", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, "|", "_")
    result = Replace(result, "`", "_")
    result = Replace(result, "~", "_")
    result = Replace(result, "!", "_")
    result = Replace(result, "@", "_")
    result = Replace(result, "#", "_")
    result = Replace(result, "^", "_")
    result = Replace(result, "*", "_")
    
    ' 2. Strip any other non-alphanumeric, non-underscore, non-period characters
    Dim i As Long
    Dim cleanResult As String
    cleanResult = vbNullString
    
    For i = 1 To Len(result)
        Dim charCode As Integer
        Dim charStr As String
        charStr = Mid$(result, i, 1)
        charCode = Asc(charStr)
        
        ' Allow a-z, A-Z, 0-9, _, .
        If (charCode >= 97 And charCode <= 122) Or _
           (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 48 And charCode <= 57) Or _
           charStr = "_" Or charStr = "." Then
            cleanResult = cleanResult & charStr
        End If
    Next i
    
    result = cleanResult
    If result = vbNullString Then
        result = "_"
    End If
    
    ' 3. Must start with letter or underscore. If it starts with digit or period, prepend underscore
    Dim firstChar As String
    firstChar = Left$(result, 1)
    If (Asc(firstChar) >= 48 And Asc(firstChar) <= 57) Or firstChar = "." Then
        result = "_" & result
    End If
    
    ' 4. Address conflict checks (e.g. A1, XFD1048576, R, C)
    Dim upperResult As String
    upperResult = UCase$(result)
    
    If upperResult = "R" Or upperResult = "C" Then
        result = result & "_"
    Else
        ' Check if it looks like an A1 cell reference (1-3 letters followed by 1-7 digits)
        Dim isCellRef As Boolean
        isCellRef = False
        
        Dim letterCount As Long: letterCount = 0
        Dim digitCount As Long: digitCount = 0
        
        For i = 1 To Len(upperResult)
            Dim charVal As String: charVal = Mid$(upperResult, i, 1)
            Dim cCode As Integer: cCode = Asc(charVal)
            
            If cCode >= 65 And cCode <= 90 Then
                If digitCount > 0 Then
                    ' Letters after digits -> not a standard cell ref
                    isCellRef = False
                    Exit For
                End If
                letterCount = letterCount + 1
            ElseIf cCode >= 48 And cCode <= 57 Then
                digitCount = digitCount + 1
            Else
                ' Other character (like underscore/period) -> not a cell ref
                isCellRef = False
                Exit For
            End If
        Next i
        
        If letterCount >= 1 And letterCount <= 3 And digitCount >= 1 And digitCount <= 7 Then
            isCellRef = True
        End If
        
        If isCellRef Then
            result = result & "_"
        End If
    End If
    
    ' 5. Maximum length of 255
    If Len(result) > 255 Then
        result = Left$(result, 255)
    End If
    
    SanitizeExcelName = result

CleanExit:
    Exit Function
ErrHandler:
    Infra_Error.HandleError "SanitizeExcelName", Err
    Resume CleanExit
End Function
