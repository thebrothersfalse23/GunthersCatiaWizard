
'===============================================================
' MODULE: generateTDSingle.bas
' PURPOSE: Implements single tool design logic for Gunther's Catia Wizard.
'          Renames and numbers products/parts in the selected structure using a prefix.
'          Supports copy-on-write, reference protection, and flexible numbering.
'===============================================================
Option Explicit

'--- NOTE: This module assumes early binding for CATIA types (Product, Products, etc.).
'--- If you get "User-defined type not defined", add the CATIA reference or use Object for late binding (see Traversal.bas).

'---------------------------------------------------------------
' Public Sub: generateTDSingle
'   Renames and numbers all products and parts in the selected structure using the given prefix.
'   Supports copy-on-write (createNewProduct), reference protection (protectRefDocs), and flexible numbering (startOnSelected).
'   Throws error if selection is invalid or has no children.
'---------------------------------------------------------------

Public Sub generateTDSingle(selectedProduct As Product, prefix As String, startOnSelected As Boolean, protectRefDocs As Boolean, createNewProduct As Boolean)
    On Error GoTo errHandler
    ' Assumes all guards (including children) have already been checked by UI/runner

    Dim workRoot As Product
    If createNewProduct Then
        Set workRoot = deepCopyStructure(selectedProduct)
        If workRoot Is Nothing Then
            MsgBox "Failed to duplicate product structure.", vbCritical, "Single Tool Design"
            Exit Sub
        End If
        ' Set the root's name/part number after copy
        safeSet workRoot, "PartNumber", prefix & "_COPY"
        safeSet workRoot, "Name", prefix & "_COPY"
    Else
        Set workRoot = selectedProduct
    End If

    ' Get all unique products and parts in traversal order (ordered, no duplicates)
    Dim uniqueList As Collection
    Set uniqueList = getUniques(workRoot, uoAll)

    ' Determine numbering start
    Dim numStart As Long
    If startOnSelected Then
        numStart = 1
    Else
        numStart = 0
    End If

    ' Rename/number all products and parts in order
    Dim i As Long
    For i = 1 To uniqueList.Count
        Dim item As Product
        Set item = uniqueList(i)
        Dim newName As String
        If i = 1 And numStart = 1 Then
            newName = prefix & "-" & Format(numStart, "0000")
        ElseIf i = 1 Then
            newName = prefix
        Else
            newName = prefix & "-" & Format(i, "0000")
        End If

        ' If protectRefDocs and this is a part with REF, preserve suffix
        If protectRefDocs And (InStr(1, UCase$(item.Name), "REF") > 0 Or InStr(1, UCase$(item.PartNumber), "REF") > 0) Then
            Dim refPos As Long
            refPos = InStr(1, UCase$(item.Name), "REF")
            If refPos > 1 Then
                newName = prefix & "-" & Format(i, "0000") & " " & Mid$(item.Name, refPos)
            ElseIf refPos = 1 Then
                newName = item.Name
            End If
        End If

        safeSet item, "PartNumber", newName
        safeSet item, "Name", newName
    Next i

    MsgBox "Single tool design complete!", vbInformation, "Gunther's Catia Wizard"
    Exit Sub

errHandler:
    MsgBox "Error in generateTDSingle: " & Err.Description, vbCritical, "Single Tool Design"
End Sub




