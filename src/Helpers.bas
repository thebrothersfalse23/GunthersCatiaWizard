'===============================================================
' MODULE: Helpers.bas
' PURPOSE: Safe property helpers and utility functions for late-bound property
'          access, key building, and string retrieval on CATIA objects.
'===============================================================

'---------------------------------------------------------------
' Sub: SafeSet
' Safely sets a string property ("Nomenclature", "Name", or "Description")
' on a given object. Ignores errors if the property does not exist.
'---------------------------------------------------------------


'---------------------------------------------------------------
' Sub: SafeSet
' Safely sets a property ("Nomenclature", "Name", "Description", "PartNumber", "Definition", "Revision", "ReferenceProduct")
' on a given object. Ignores errors if the property does not exist.
'
' Parameters:
'   obj      - The object on which to set the property.
'   propName - The name of the property to set.
'   value    - The value to assign to the property.
'---------------------------------------------------------------
Public Sub SafeSet(ByVal obj As Object, ByVal propName As String, ByVal value As String)
    On Error Resume Next
    Select Case propName
        Case "Nomenclature":      obj.Nomenclature = value
        Case "Name":              obj.Name = value
        Case "Description":       obj.Description = value
        Case "PartNumber":        obj.PartNumber = value
        Case "Definition":        obj.Definition = value
        Case "Revision":          obj.Revision = value
        Case "ReferenceProduct":  obj.ReferenceProduct = value
    End Select
    Err.Clear
    On Error GoTo 0
End Sub


'---------------------------------------------------------------
' Function: GetPropStr
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
Public Function GetPropStr(ByVal obj As Object, ByVal propName As String) As String
    On Error Resume Next
    Select Case propName
        Case "Nomenclature":      GetPropStr = obj.Nomenclature
        Case "Name":              GetPropStr = obj.Name
        Case "Description":       GetPropStr = obj.Description
        Case "PartNumber":        GetPropStr = obj.PartNumber
        Case "Definition":        GetPropStr = obj.Definition
        Case "Revision":          GetPropStr = obj.Revision
        Case "ReferenceProduct":  GetPropStr = obj.ReferenceProduct
        Case Else:                GetPropStr = ""
    End Select
    If Err.Number <> 0 Then GetPropStr = ""
    Err.Clear
    On Error GoTo 0
End Function

'===============================================================
' BuildRefKey – Builds a stable, human-readable key for a reference
'===============================================================

'---------------------------------------------------------------
' Function: BuildRefKey
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
Public Function BuildRefKey(ByVal ref As Product, ByVal docType As String) As String
    On Error Resume Next
    Dim pn As String: pn = ref.PartNumber
    Dim defn As String: defn = ref.Definition ' may be empty depending on env
    On Error GoTo 0

    If Len(pn) = 0 Then
        BuildRefKey = ""
    ElseIf Len(defn) > 0 Then
        BuildRefKey = pn & "|" & docType & "|" & defn
    Else
        BuildRefKey = pn & "|" & docType
    End If
End Function
