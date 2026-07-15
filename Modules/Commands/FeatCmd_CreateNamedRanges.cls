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
    
    ' Fetch saved range preferences
    Dim lastValAddr As String
    lastValAddr = Infra_Interaction.GetUserPreference("CreateNamedRanges", "LastValueRange", "")
    Dim lastLblAddr As String
    lastLblAddr = Infra_Interaction.GetUserPreference("CreateNamedRanges", "LastLabelRange", "")
    
    If Not skipValuePrompt Then
        ' Prompt 1: Select Value Range
        Dim promptValSuccess As Boolean
        If lastValAddr <> "" Then
            promptValSuccess = Infra_Interaction.PromptRange("Select the Value Range (cells containing numbers/formulas to name):", "Select Value Range", valueRange, lastValAddr)
        Else
            promptValSuccess = Infra_Interaction.PromptRange("Select the Value Range (cells containing numbers/formulas to name):", "Select Value Range", valueRange)
        End If
        If Not promptValSuccess Then GoTo CleanExit
    End If
    
    ' Prompt 2: Select Label Range
    Dim labelRange As Range
    Dim expectedCount As Long: expectedCount = valueRange.Cells.CountLarge
    Dim labelPrompt As String
    labelPrompt = "Select the Label Range containing exactly " & expectedCount & " cell"
    If expectedCount <> 1 Then labelPrompt = labelPrompt & "s"
    labelPrompt = labelPrompt & ":"
    
    Dim promptLblSuccess As Boolean
    If lastLblAddr <> "" Then
        promptLblSuccess = Infra_Interaction.PromptRange(labelPrompt, "Select Label Range", labelRange, lastLblAddr)
    Else
        promptLblSuccess = Infra_Interaction.PromptRange(labelPrompt, "Select Label Range", labelRange)
    End If
    If Not promptLblSuccess Then GoTo CleanExit
    
    ' Validation: Cells count must match
    If valueRange.Cells.CountLarge <> labelRange.Cells.CountLarge Then
        Infra_Interaction.ShowWarning "The Value Range has " & valueRange.Cells.CountLarge & " cells, but the Label Range has " & labelRange.Cells.CountLarge & " cells." & vbCrLf & vbCrLf & "Both ranges must contain the same number of cells.", "Bulk Named Ranges"
        GoTo CleanExit
    End If
    
    ' Save user preferences for the next execution
    Infra_Interaction.SaveUserPreference "CreateNamedRanges", "LastValueRange", valueRange.Address(RowAbsolute:=True, ColumnAbsolute:=True)
    Infra_Interaction.SaveUserPreference "CreateNamedRanges", "LastLabelRange", labelRange.Address(RowAbsolute:=True, ColumnAbsolute:=True)
    
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
    
    ' Crop ranges to active worksheet bounds to prevent looping 1M+ blank rows
    Dim lastUsedRow As Long
    lastUsedRow = targetWs.UsedRange.Row + targetWs.UsedRange.Rows.Count - 1
    If lastUsedRow < 1 Then lastUsedRow = 1
    
    Dim valStartRow As Long: valStartRow = valueRange.Row
    Dim valEndRow As Long: valEndRow = valStartRow + valueRange.Rows.Count - 1
    
    If valEndRow > lastUsedRow And valueRange.Rows.Count > 1 Then
        Dim newRowCount As Long
        newRowCount = lastUsedRow - valStartRow + 1
        If newRowCount < 1 Then newRowCount = 1
        
        If newRowCount < valueRange.Rows.Count Then
            Set valueRange = valueRange.Resize(newRowCount)
            Set labelRange = labelRange.Resize(newRowCount)
        End If
    End If
    
    Dim lastUsedCol As Long
    lastUsedCol = targetWs.UsedRange.Column + targetWs.UsedRange.Columns.Count - 1
    If lastUsedCol < 1 Then lastUsedCol = 1
    
    Dim valStartCol As Long: valStartCol = valueRange.Column
    Dim valEndCol As Long: valEndCol = valStartCol + valueRange.Columns.Count - 1
    
    If valEndCol > lastUsedCol And valueRange.Columns.Count > 1 Then
        Dim newColCount As Long
        newColCount = lastUsedCol - valStartCol + 1
        If newColCount < 1 Then newColCount = 1
        
        If newColCount < valueRange.Columns.Count Then
            Set valueRange = valueRange.Resize(, newColCount)
            Set labelRange = labelRange.Resize(, newColCount)
        End If
    End If
    
    Dim expectedCount As Long: expectedCount = valueRange.Cells.CountLarge
    If expectedCount <= 0 Then
        ExecuteBulkNamedRangesDirect = True
        Exit Function
    End If
    
    ' Read values into variant arrays for fast memory checks
    Dim valArr As Variant
    Dim lblArr As Variant
    Dim numCols As Long: numCols = valueRange.Columns.Count
    
    If expectedCount > 1 Then
        valArr = valueRange.Value2
        lblArr = labelRange.Value2
    End If
    
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
        Dim fullName As String: fullName = nameObj.Name
        Dim pos As Long: pos = InStr(fullName, "!")
        Dim isRelevant As Boolean: isRelevant = True
        Dim localName As String
        
        If pos > 0 Then
            Dim sheetPrefix As String: sheetPrefix = Left$(fullName, pos - 1)
            ' Strip single quotes from sheet prefix if present
            If Left$(sheetPrefix, 1) = "'" And Right$(sheetPrefix, 1) = "'" Then
                sheetPrefix = Mid$(sheetPrefix, 2, Len(sheetPrefix) - 2)
            End If
            
            If StrComp(sheetPrefix, targetWs.Name, vbTextCompare) <> 0 Then
                isRelevant = False
            End If
            localName = Mid$(fullName, pos + 1)
        Else
            localName = fullName
        End If
        
        If isRelevant Then
            If Not dictNamesByName.Exists(localName) Then
                Set dictNamesByName(localName) = nameObj
            End If
            
            ' Check RefersTo to see if it is a range reference on targetWs
            Dim refersToFormula As String
            refersToFormula = vbNullString
            On Error Resume Next
            refersToFormula = nameObj.RefersTo
            On Error GoTo ErrHandler
            
            Dim sheetName As String
            Dim rangeAddress As String
            If Infra_CommandSupport.ParseSimpleRefersTo(refersToFormula, sheetName, rangeAddress) Then
                If StrComp(sheetName, targetWs.Name, vbTextCompare) = 0 Then
                    Dim refersRange As Range: Set refersRange = Nothing
                    On Error Resume Next
                    Set refersRange = targetWs.Range(rangeAddress)
                    On Error GoTo ErrHandler
                    
                    If Not refersRange Is Nothing Then
                        If Infra_CommandSupport.IsRangeInsideRange(refersRange, valueRange) Then
                            Dim cellAddr As String
                            cellAddr = refersRange.Address(External:=True)
                            If Not dictNamesByCell.Exists(cellAddr) Then
                                Dim nameCol As Collection: Set nameCol = New Collection
                                dictNamesByCell.Add cellAddr, nameCol
                            End If
                            dictNamesByCell(cellAddr).Add localName
                        End If
                    End If
                End If
            End If
        End If
    Next nameObj
    
    ' Get column counts for each range to support transposed/differently-shaped mappings
    Dim valCols As Long: valCols = valueRange.Columns.Count
    Dim lblCols As Long: lblCols = labelRange.Columns.Count
    
    ' Fetch the external prefix once to construct external addresses in memory
    Dim dummyAddr As String
    Dim extPrefix As String
    dummyAddr = targetWs.Cells(1, 1).Address(External:=True)
    extPrefix = Left$(dummyAddr, InStrRev(dummyAddr, "!") - 1)
    
    Dim createdNamesCol As New Collection
    
    Dim i As Long
    For i = 1 To expectedCount
        Dim lblVal As Variant
        Dim valVal As Variant
        
        Dim valR As Long, valC As Long
        If expectedCount > 1 Then
            valR = ((i - 1) \ valCols) + 1
            valC = ((i - 1) Mod valCols) + 1
            valVal = valArr(valR, valC)
            
            Dim lblR As Long, lblC As Long
            lblR = ((i - 1) \ lblCols) + 1
            lblC = ((i - 1) Mod lblCols) + 1
            lblVal = lblArr(lblR, lblC)
        Else
            valR = 1
            valC = 1
            lblVal = labelRange.Value2
            valVal = valueRange.Value2
        End If
        
        ' Skip if label is empty, null, or error
        If IsEmpty(lblVal) Or IsNull(lblVal) Or IsError(lblVal) Then
            GoTo NextCell
        End If
        
        Dim rawLabel As String: rawLabel = Trim$(CStr(lblVal))
        If rawLabel = "" Then GoTo NextCell
        
        ' Skip if value is empty or null
        If IsEmpty(valVal) Or IsNull(valVal) Then
            GoTo NextCell
        End If
        
        If Not IsError(valVal) Then
            If Trim$(CStr(valVal)) = "" Then
                GoTo NextCell
            End If
        End If
        
        Dim sanitizedName As String: sanitizedName = SanitizeExcelName(rawLabel)
        
        ' Check if name already exists
        Dim nameExists As Boolean: nameExists = dictNamesByName.Exists(sanitizedName)
        Dim existingName As Name: Set existingName = Nothing
        
        ' Calculate row and column numbers of target cell
        Dim cellRow As Long: cellRow = valStartRow + valR - 1
        Dim cellCol As Long: cellCol = valStartCol + valC - 1
        
        ' Construct external address in memory
        Dim currentAddr As String
        currentAddr = extPrefix & "!$" & Infra_CommandSupport.GetColLetter(cellCol) & "$" & cellRow
        
        Dim isRedundant As Boolean: isRedundant = False
        Dim valueCell As Range: Set valueCell = Nothing
        
        If nameExists Then
            Set existingName = dictNamesByName(sanitizedName)
            
            ' Check if it points to the SAME cell using dictNamesByCell
            If dictNamesByCell.Exists(currentAddr) Then
                Dim col As Collection: Set col = dictNamesByCell(currentAddr)
                Dim nameStr As Variant
                For Each nameStr In col
                    If StrComp(CStr(nameStr), sanitizedName, vbTextCompare) = 0 Then
                        isRedundant = True
                        Exit For
                    End If
                Next nameStr
            End If
            
            If isRedundant Then
                ' Redundant: Already named sanitizedName. Skip silently.
                GoTo NextCell
            Else
                ' Points to a DIFFERENT cell. This is a conflict!
                Dim existingRng As Range: Set existingRng = Nothing
                On Error Resume Next
                Set existingRng = existingName.RefersToRange
                On Error GoTo ErrHandler
                
                Dim existingCellRef As String: existingCellRef = ""
                If Not existingRng Is Nothing Then
                    existingCellRef = existingRng.Parent.Name & "!" & existingRng.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                Else
                    existingCellRef = "another formula/reference"
                End If
                
                Set valueCell = valueRange.Cells(i)
                Dim targetCellRef As String
                targetCellRef = valueCell.Parent.Name & "!" & valueCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                Infra_Interaction.ShowWarning "Cannot name cell " & targetCellRef & " as '" & sanitizedName & "' because '" & sanitizedName & "' already refers to cell " & existingCellRef & "." & vbCrLf & vbCrLf & "Both cells cannot have the same name.", "Bulk Named Ranges"
                GoTo NextCell
            End If
        End If
        
        ' Clean up other names pointing to this cell (renaming support)
        If dictNamesByCell.Exists(currentAddr) Then
            Dim cellNamesCol As Collection: Set cellNamesCol = dictNamesByCell(currentAddr)
            Dim j As Long
            ' Iterate backwards when deleting/modifying collections
            For j = cellNamesCol.Count To 1 Step -1
                Dim oldLocalName As String: oldLocalName = cellNamesCol(j)
                If StrComp(oldLocalName, sanitizedName, vbTextCompare) <> 0 Then
                    If dictNamesByName.Exists(oldLocalName) Then
                        Dim oldNameObj As Name: Set oldNameObj = dictNamesByName(oldLocalName)
                        On Error Resume Next
                        oldNameObj.Delete
                        On Error GoTo ErrHandler
                        
                        ' Remove from caches
                        dictNamesByName.Remove oldLocalName
                    End If
                    cellNamesCol.Remove j
                End If
            Next j
        End If
        
        ' Create Name
        Dim newNameObj As Name: Set newNameObj = Nothing
        If valueCell Is Nothing Then Set valueCell = valueRange.Cells(i)
        
        If scopeChoice = "Workbook" Then
            Set newNameObj = targetWb.Names.Add(Name:=sanitizedName, RefersTo:=valueCell)
        Else
            Set newNameObj = targetWs.Names.Add(Name:=sanitizedName, RefersTo:=valueCell)
        End If
        
        ' Update Cache
        If Not newNameObj Is Nothing Then
            Set dictNamesByName(sanitizedName) = newNameObj
            If Not dictNamesByCell.Exists(currentAddr) Then
                Dim newCol As Collection: Set newCol = New Collection
                dictNamesByCell.Add currentAddr, newCol
            End If
            dictNamesByCell(currentAddr).Add sanitizedName
            
            ' Accumulate for custom Undo
            If outCreatedList = "" Then
                outCreatedList = newNameObj.Name
            Else
                outCreatedList = outCreatedList & ";" & newNameObj.Name
            End If
            
            createdNamesCol.Add GetLocalName(newNameObj)
            outCreatedCount = outCreatedCount + 1
        End If
        
NextCell:
    Next i
    
    ' Apply names to formulas on the sheet if names were created/modified
    If createdNamesCol.Count > 0 Then
        Dim namesArray() As String
        ReDim namesArray(0 To createdNamesCol.Count - 1)
        Dim idx As Long
        For idx = 0 To createdNamesCol.Count - 1
            namesArray(idx) = createdNamesCol(idx + 1)
        Next idx
        
        Dim hasFormulas As Boolean
        hasFormulas = False
        Dim formulaCells As Range
        On Error Resume Next
        Set formulaCells = targetWs.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Not formulaCells Is Nothing Then hasFormulas = True
        On Error GoTo ErrHandler
        
        If hasFormulas Then
            On Error Resume Next
            targetWs.UsedRange.ApplyNames Names:=namesArray
            On Error GoTo ErrHandler
        End If

    End If
    
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
    
    Dim bytes() As Byte
    bytes = result
    
    Dim outBytes() As Byte
    ReDim outBytes(0 To UBound(bytes))
    
    Dim outIdx As Long
    outIdx = 0
    
    Dim lastWasUnderscore As Boolean
    lastWasUnderscore = False
    
    Dim i As Long
    For i = 0 To UBound(bytes) Step 2
        Dim b1 As Byte: b1 = bytes(i)
        Dim b2 As Byte: b2 = bytes(i + 1)
        
        If b2 = 0 Then
            ' ASCII characters
            Dim isAllowed As Boolean: isAllowed = False
            Dim isUnderscore As Boolean: isUnderscore = False
            
            ' a-z (97-122), A-Z (65-90), 0-9 (48-57), . (46)
            If (b1 >= 97 And b1 <= 122) Or (b1 >= 65 And b1 <= 90) Or (b1 >= 48 And b1 <= 57) Or b1 = 46 Then
                isAllowed = True
            ' _ (95)
            ElseIf b1 = 95 Then
                isAllowed = True
                isUnderscore = True
            ' Spaces, hyphens, slashes, and common special chars replaced with underscores
            ElseIf b1 = 32 Or b1 = 45 Or b1 = 47 Or b1 = 92 Or b1 = 37 Or b1 = 36 Or b1 = 38 Or _
                   b1 = 40 Or b1 = 41 Or b1 = 43 Or b1 = 61 Or b1 = 91 Or b1 = 93 Or b1 = 123 Or _
                   b1 = 125 Or b1 = 44 Or b1 = 59 Or b1 = 58 Or b1 = 39 Or b1 = 34 Or b1 = 60 Or _
                   b1 = 62 Or b1 = 63 Or b1 = 124 Or b1 = 96 Or b1 = 126 Or b1 = 33 Or b1 = 64 Or _
                   b1 = 35 Or b1 = 94 Or b1 = 42 Then
                isAllowed = True
                b1 = 95
                isUnderscore = True
            End If
            
            If isAllowed Then
                If isUnderscore Then
                    If Not lastWasUnderscore Then
                        outBytes(outIdx) = b1
                        outBytes(outIdx + 1) = 0
                        outIdx = outIdx + 2
                        lastWasUnderscore = True
                    End If
                Else
                    outBytes(outIdx) = b1
                    outBytes(outIdx + 1) = 0
                    outIdx = outIdx + 2
                    lastWasUnderscore = False
                End If
            End If
        End If
    Next i
    
    If outIdx > 0 Then
        ReDim Preserve outBytes(0 To outIdx - 1)
        result = outBytes
    Else
        result = "_"
    End If
    
    If result = vbNullString Then
        result = "_"
    End If
    
    ' 3. Must start with letter or underscore. If it starts with digit or period, prepend underscore
    Dim firstChar As String
    firstChar = Left$(result, 1)
    If (Asc(firstChar) >= 48 And Asc(firstChar) <= 57) Or firstChar = "." Then
        result = "_" & result
    End If
    
    ' 4. Address conflict and reserved word checks
    Dim upperResult As String
    upperResult = UCase$(result)
    
    Dim isReserved As Boolean
    Select Case upperResult
        Case "PRINT_AREA", "PRINT_TITLES", "DATABASE", "CRITERIA", "EXTRACT", "CONSOLIDATE_AREA", "SHEET_TITLE"
            isReserved = True
        Case Else
            isReserved = False
    End Select
    
    If isReserved Or upperResult = "R" Or upperResult = "C" Then
        result = result & "_"
        upperResult = UCase$(result)
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
            upperResult = UCase$(result)
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
