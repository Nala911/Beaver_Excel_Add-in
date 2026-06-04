Attribute VB_Name = "Lib_UdfRegistry"
Option Explicit
Option Private Module

' @Module: Lib_UdfRegistry
' @Category: Library
' @Description: Central registry of User Defined Functions (UDFs) metadata for Beaver Add-in.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error

' Returns a collection of metadata dictionaries for all registered UDFs.
' Each dictionary contains:
'   - Name: String
'   - Description: String
'   - Category: String
'   - Syntax: String
'   - ArgumentDescriptions: Variant Array of Strings
Public Function GetAllUdfs() As Collection
    Dim tracker As Object: Set tracker = Infra_Error.Track("GetAllUdfs")
    On Error GoTo ErrHandler

    Dim registry As New Collection

    ' Register XFilter
    registry.Add GetXFilterMetadata()

    ' Register XUnpivot
    registry.Add GetXUnpivotMetadata()

    ' In the future, additional UDFs can register here:
    ' registry.Add GetAnotherFunctionMetadata()

    Set GetAllUdfs = registry

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "GetAllUdfs", Err
    Resume CleanExit
End Function

' --- PRIVATE UDF METADATA BUILDERS ---

Private Function GetXFilterMetadata() As Object
    Dim metadata As Object
    Set metadata = CreateObject("Scripting.Dictionary")
    metadata.Add "Name", "XFilter"
    metadata.Add "Description", "Filters Range_A based on existence (or non-existence) in Range_B."
    metadata.Add "Category", "User Defined"
    metadata.Add "Syntax", "XFilter(Range_A, Range_B, [code_number], [if_empty], [case_sensitive])"
    metadata.Add "ArgumentDescriptions", Array( _
        "The source range or array to filter. Comparison is done using the first column.", _
        "The reference range or array to check against.", _
        "Optional Mode: 1 = Intersection (default, In A and B), 2 = Difference (In A but not in B).", _
        "Optional value to return if no match is found (default is ""Not found"").", _
        "Optional. Set to True for case-sensitive matching; defaults to False.")
    Set GetXFilterMetadata = metadata
End Function

Private Function GetXUnpivotMetadata() As Object
    Dim metadata As Object
    Set metadata = CreateObject("Scripting.Dictionary")
    metadata.Add "Name", "XUnpivot"
    metadata.Add "Description", "Unpivots wide range data into a long 2D layout, auto-detecting numerical columns from the first data row."
    metadata.Add "Category", "User Defined"
    metadata.Add "Syntax", "XUnpivot(data_range, [attribute_header], [value_header], [skip_blanks])"
    metadata.Add "ArgumentDescriptions", Array( _
        "The source range or array to unpivot. First row must contain column headers.", _
        "Optional name for the attribute/variable column (default is ""Attribute"").", _
        "Optional name for the value column (default is ""Value"").", _
        "Optional boolean. If True, skips unpivot cells that are blank or empty (default is False).")
    Set GetXUnpivotMetadata = metadata
End Function

