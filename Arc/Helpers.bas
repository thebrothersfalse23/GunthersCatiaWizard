'===============================================================
' MODULE: helpers.bas
' PURPOSE: Safe property helpers and utility functions for late-bound property
'          access, key building, and string retrieval on CATIA objects.
'===============================================================

Private Const CATIA_TYPE_PRODUCT As String = "Product"

'---------------------------------------------------------------
' Sub: safeSet
'   Safely sets a property ("Nomenclature", "Name", "Description", "PartNumber", "Revision")
'   on a given object. Ignores errors if the property does not exist.
'
' Parameters:
'   obj      - The object on which to set the property.
'   propName - The name of the property to set.
'   value    - The value to assign to the property.
'---------------------------------------------------------------
Public Sub safeSet(ByVal obj As Object, ByVal propName As String, ByVal value As String)
    On Error Resume Next
    Select Case propName
        Case "Nomenclature":      obj.Nomenclature = value
        Case "Name":              obj.Name = value
        Case "Description":       obj.Description = value
        Case "PartNumber":        obj.PartNumber = value
        Case "Revision":          obj.Revision = value
        ' "Definition" and "ReferenceProduct" are read-only for most Product objects; do not set
    End Select
    Err.Clear
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' Function: getPropStr
'   Safely retrieves a string property ("Nomenclature", "Name", "Description", etc.)
'   from a given object. Returns an empty string if the property does not exist
'   or an error occurs.
'
' Parameters:
'   obj      - The object from which to retrieve the property.
'   propName - The name of the property to retrieve.
'
' Returns:
'   String   - The value of the specified property, or an empty string on error.
'---------------------------------------------------------------
Public Function getPropStr(ByVal obj As Object, ByVal propName As String) As String
    On Error Resume Next
    Select Case propName
        Case "Nomenclature":      getPropStr = obj.Nomenclature
        Case "Name":              getPropStr = obj.Name
        Case "Description":       getPropStr = obj.Description
        Case "PartNumber":        getPropStr = obj.PartNumber
        Case "Definition":        getPropStr = obj.Definition
        Case "Revision":          getPropStr = obj.Revision
        Case "ReferenceProduct":  getPropStr = obj.ReferenceProduct
        Case Else:                getPropStr = ""
    End Select
    If Err.Number <> 0 Then getPropStr = ""
    Err.Clear
    On Error GoTo 0
End Function

'' buildRefKey is now defined in improvedTraversal and should not be duplicated here.

'---------------------------------------------------------------
' Function: getSelectedProducts
'   Returns selected Product(s) from CATIA.
'
'   Parameters:
'     firstSelection [Boolean] - If True, returns only the first selected Product (as a Product object or Nothing).
'                                If False, returns all selected Products as a Collection.
'
'   Returns:
'     If firstSelection = True:   Product (or Nothing if none selected or first selection is not a Product)
'     If firstSelection = False:  Collection of Product objects (may be empty)
'   Note:
'     Assumes guards have already validated CATIA, document, and selection state.
'---------------------------------------------------------------
Public Function getSelectedProducts(Optional ByVal firstSelection As Boolean = False) As Variant
    Dim sel As Selection
    Dim prod As Product
    Dim result As Collection
    Set sel = CATIA.ActiveDocument.Selection

    If firstSelection Then
        If sel.Count >= 1 Then
            If TypeName(sel.Item(1).Value) = "Product" Then
                Set prod = sel.Item(1).Value
                If Not prod Is Nothing Then
                    Set getSelectedProducts = prod
                    Exit Function
                End If
            End If
        End If
        Set getSelectedProducts = Nothing
    Else
        Set result = New Collection
        Dim j As Integer
        For j = 1 To sel.Count
            Set prod = Nothing
            If TypeName(sel.Item(j).Value) = "Product" Then
                Set prod = sel.Item(j).Value
                If Not prod Is Nothing Then
                    result.Add prod
                End If
            End If
        Next j
        Set getSelectedProducts = result
    End If
End Function
'---------------------------------------------------------------
'---------------------------------------------------------------
