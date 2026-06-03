Attribute VB_Name = "Lib_UdfRegistry"
Option Explicit

' @Module: Lib_UdfRegistry
' @Category: Library
' @Description: Central registry of User Defined Functions (UDFs) metadata for Beaver Add-in.
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Infra_Error, Lib_XFilterFunction

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
    registry.Add Lib_XFilterFunction.GetUdfMetadata()

    ' In the future, additional UDFs can register here:
    ' registry.Add Lib_AnotherFunction.GetUdfMetadata()

    Set GetAllUdfs = registry

CleanExit:
    Exit Function

ErrHandler:
    Infra_Error.HandleError "GetAllUdfs", Err
    Resume CleanExit
End Function
