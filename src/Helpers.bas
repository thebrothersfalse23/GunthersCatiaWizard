'===============================================================
' MODULE: helpers.bas
' PURPOSE: Safe property helpers and utility functions for late-bound property
'          access, key building, and string retrieval on CATIA objects.
'===============================================================


'---------------------------------------------------------------
' Sub: safeSet
' Safely sets a property ("Nomenclature", "Name", "Description", "PartNumber", "Revision")
' on a given object. Ignores errors if the property does not exist.
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
' Safely retrieves a string property ("Nomenclature", "Name", "Description", etc.)
' from a given object. Returns an empty string if the property does not exist
' or an error occurs.
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

'===============================================================
' buildRefKey – Builds a stable, human-readable key for a reference
'===============================================================

'---------------------------------------------------------------
' Function: buildRefKey
' Builds a stable, human-readable key for a reference Product.
' Default: "PartNumber|DocType"; if Definition exists → "PartNumber|DocType|Definition"
'
' Parameters:
'   ref     - The reference Product object.
'   docType - The document type as a string ("ProductDocument" or "PartDocument").
'
' Returns:
'   String  - The constructed key, or "" if PartNumber is empty.
'---------------------------------------------------------------
Public Function buildRefKey(ByVal ref As Product, ByVal docType As String) As String
    On Error Resume Next
    Dim pn As String: pn = Trim$(ref.PartNumber)
    Dim defn As String: defn = Trim$(ref.Definition)
    On Error GoTo 0

    If Len(pn) = 0 Then
        buildRefKey = ""
    ElseIf Len(defn) > 0 Then
        buildRefKey = UCase$(pn) & "|" & docType & "|" & UCase$(defn)
    Else
        buildRefKey = UCase$(pn) & "|" & docType
    End If
End Function

'---------------------------------------------------------------
' Function: GetSelectedProduct
' Returns the currently selected Product in CATIA, or Nothing if not found.
'---------------------------------------------------------------
Public Function GetSelectedProduct() As Product
    On Error Resume Next
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    If sel Is Nothing Or sel.Count = 0 Then
        Set GetSelectedProduct = Nothing
        Exit Function
    End If

    Dim i As Integer
    For i = 1 To sel.Count
        Dim obj As Object
        Set obj = sel.Item(i).Value
        If TypeName(obj) = "Product" Then
            Set GetSelectedProduct = obj
            Exit Function
        End If
    Next i
    Set GetSelectedProduct = Nothing
End Function
