Attribute VB_Name = "Feat_FilterByCell"
Option Explicit

' @Module: Feat_FilterByCell
' @Category: Feature
' @Description: Quick-filter a range based on the value of the active cell.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_AppState, Infra_Config, Infra_Error

Public Sub FilterBySelectedCell()
    PushContext "FilterBySelectedCell"
    On Error GoTo ErrHandler

    Dim guard As New Infra_AppStateGuard
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim dataRange As Range
    Dim tableObject As ListObject
    Dim filterRange As Range
    Dim filterColumn As Long
    Dim filterValue As Variant
    Dim numValue As Double
    Dim epsilon As Double
    Dim startDate As Long
    Dim isDateFilter As Boolean
    Dim isErrorFilter As Boolean
    Dim isBlankFilter As Boolean
    Dim isNumericFilter As Boolean
    Dim isTextFilter As Boolean
    Dim currentFilter As AutoFilter
    Dim currentField As Filter
    Dim existingCriteria As Variant
    Dim newCriteria As Variant
    
    Dim pt As PivotTable
    Dim ptRange As Range
    Dim pivotEndCol As Long
    Dim pivotStartCol As Long
    Dim offsetCols As Long
    Dim newWidth As Long

    Dim ctx As Infra_ActionContext
    Set ctx = Infra_AppState.CaptureActionContext()
    If ctx Is Nothing Then GoTo CleanExit
    If ctx.WorksheetRef Is Nothing Then GoTo CleanExit
    
    Set ws = ctx.WorksheetRef
    Set selectedCell = ctx.ActiveCellRef
    If selectedCell Is Nothing Then GoTo CleanExit

    ' --- Determine Filter Type ---
    If IsEmpty(selectedCell.Value) Then
        isBlankFilter = True
        filterValue = "="
    ElseIf IsError(selectedCell.Value) Then
        isErrorFilter = True
        filterValue = "=" & selectedCell.Text
    ElseIf IsDate(selectedCell.Value) Then
        isDateFilter = True
        filterValue = selectedCell.Value
    ElseIf IsNumeric(selectedCell.Value) Then
        isNumericFilter = True
        numValue = selectedCell.Value
        epsilon = 0.000000001
    Else
        isTextFilter = True
        filterValue = selectedCell.Value
    End If

    On Error Resume Next
    Set tableObject = selectedCell.ListObject
    On Error GoTo ErrHandler

    If Not tableObject Is Nothing Then
        Set dataRange = tableObject.DataBodyRange
        Set filterRange = tableObject.Range
        filterColumn = selectedCell.Column - tableObject.Range.Columns(1).Column + 1
        Set currentFilter = tableObject.AutoFilter
    Else
        Set dataRange = selectedCell.CurrentRegion
        
        ' --- PivotTable Conflict Check ---
        For Each pt In ws.PivotTables
            Set ptRange = pt.TableRange2
            If Not Intersect(dataRange, ptRange) Is Nothing Then
                If Intersect(selectedCell, ptRange) Is Nothing Then
                    If selectedCell.Column > (ptRange.Column + ptRange.Columns.Count - 1) Then
                        pivotEndCol = ptRange.Column + ptRange.Columns.Count - 1
                        if dataRange.Column <= pivotEndCol Then
                            offsetCols = (pivotEndCol + 1) - dataRange.Column
                            Set dataRange = dataRange.Resize(, dataRange.Columns.Count - offsetCols).Offset(, offsetCols)
                        End If
                    ElseIf selectedCell.Column < ptRange.Column Then
                        pivotStartCol = ptRange.Column
                        If (dataRange.Column + dataRange.Columns.Count - 1) >= pivotStartCol Then
                            newWidth = pivotStartCol - dataRange.Column
                            Set dataRange = dataRange.Resize(, newWidth)
                        End If
                    End If
                End If
            End If
        Next pt

        If dataRange.Cells.Count <= 1 Then
            MsgBox "The selected cell is not part of a larger data range.", vbExclamation, Infra_Config.ADDIN_NAME
            GoTo CleanExit
        End If
        
        Set filterRange = dataRange
        Set currentFilter = ws.AutoFilter
        filterColumn = selectedCell.Column - dataRange.Columns(1).Column + 1
    End If

    ' --- Capture existing criteria ---
    On Error Resume Next
    Set currentField = Nothing
    If Not currentFilter Is Nothing Then
        If filterColumn > 0 And filterColumn <= currentFilter.Filters.Count Then
            Set currentField = currentFilter.Filters(filterColumn)
        End If
    End If
    If Not currentField Is Nothing Then
        If currentField.On Then existingCriteria = currentField.Criteria1
    End If
    On Error GoTo ErrHandler

    ' --- Apply filter ---
    Select Case True
        Case isBlankFilter
            newCriteria = "="
        Case isErrorFilter
            newCriteria = filterValue
        Case isDateFilter
            startDate = CLng(filterValue)
            filterRange.AutoFilter Field:=filterColumn, _
                                   Criteria1:=">=" & startDate, _
                                   Operator:=xlAnd, _
                                   Criteria2:="<" & startDate + 1
            GoTo CleanExit
        Case isNumericFilter
            filterRange.AutoFilter Field:=filterColumn, _
                                   Criteria1:=">=" & CStr(numValue - epsilon), _
                                   Operator:=xlAnd, _
                                   Criteria2:="<=" & CStr(numValue + epsilon)
            GoTo CleanExit
        Case isTextFilter
            newCriteria = "=" & filterValue
    End Select

    ' --- Combine existing and new criteria ---
    If Not IsEmpty(existingCriteria) Then
        Dim combinedCriteria As Collection
        Set combinedCriteria = New Collection
        
        ' Add existing criteria
        If IsArray(existingCriteria) Then
            Dim v As Variant
            For Each v In existingCriteria
                combinedCriteria.Add CStr(v)
            Next v
        Else
            combinedCriteria.Add CStr(existingCriteria)
        End If
        
        ' Add new criteria if not already present
        Dim exists As Boolean
        exists = False
        For Each v In combinedCriteria
            If v = CStr(newCriteria) Then
                exists = True
                Exit For
            End If
        Next v
        
        If Not exists Then combinedCriteria.Add CStr(newCriteria)
        
        ' Convert collection to array for AutoFilter
        Dim criteriaArray() As String
        ReDim criteriaArray(0 To combinedCriteria.Count - 1)
        Dim i As Long
        For i = 1 To combinedCriteria.Count
            criteriaArray(i - 1) = combinedCriteria(i)
        Next i
        
        filterRange.AutoFilter Field:=filterColumn, _
                               Criteria1:=criteriaArray, _
                               Operator:=xlFilterValues
    Else
        filterRange.AutoFilter Field:=filterColumn, _
                               Criteria1:=newCriteria
    End If

CleanExit:
    PopContext
    Exit Sub

ErrHandler:
    HandleError "FilterBySelectedCell", Err
End Sub
