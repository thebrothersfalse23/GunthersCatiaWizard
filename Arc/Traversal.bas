
'====================================================================
' MODULE: Traversal.bas
' PURPOSE: Centralized traversal logic for CATIA Product structures.
'          Implements a single BFS traversal with mode switching for various
'          collection and manipulation tasks. All traversal modes and constants
'          are defined here as Public Const for use throughout the project.
' CATIA:   V5 (V5-6R2x compatible)
' UPDATED: 2025-08-23
' NOTES:
'   - One traversal loop; select-case dispatch per traversal mode (see constants below).
'   - Early child enqueue only for tmDeepCopyStructure (copy requires instance-first order).
'   - Late child enqueue for all other modes (prevents double-enqueue).
'   - Uses minimal, safe error scopes (On Error ... Next -> GoTo 0).
'   - Avoids deprecated/recorded artifacts; uses Products/ReferenceProduct safely.
'   - All traversal-related constants and helpers are defined here; no enums used.
'====================================================================

Option Explicit

'--- Requires reference to "CATIA V5 Type Library"
'--- If not present, go to Tools > References in VBA editor and check "CATIA V5 Type Library"
'--- Or, if using late binding, declare as Object instead of Product/Products

'-----------------------------
' Strongly-typed “enums”
'-----------------------------
Public Const tmGetUniques          As Integer = 0
Public Const tmGetProducts         As Integer = 1
Public Const tmGetParts            As Integer = 2
Public Const tmGetInstances        As Integer = 3
Public Const tmDeepCopyStructure   As Integer = 4  ' deep-copy controls traversal order

Public Const uoAll                 As Integer = 0
Public Const uoProductsOnly        As Integer = 1
Public Const uoPartsOnly           As Integer = 2

'--- CATIA type declarations (remove/comment if using late binding)
'--- If you get "User-defined type not defined", add the CATIA reference or use Object
'Dim Product As Object
'Dim Products As Object

'-----------------------------
' Core traversal (single BFS)
'-----------------------------
Public Sub traverseProduct( _
        ByVal mode As Integer, _
        ByVal root As Product, _
        Optional ByRef outRefs As Collection, _
    '--- Mode-specific state
    Select Case mode
        Case tmGetUniques, tmGetProducts, tmGetParts
            If outRefs Is Nothing Then Set outRefs = New Collection
            Set dictSeen = CreateObject("Scripting.Dictionary")
        Case tmGetInstances
            If outRefs Is Nothing Then Set outRefs = New Collection
        Case tmDeepCopyStructure
            If outRefs Is Nothing Then Set outRefs = New Collection
            ' Initialize any copy state here if needed (e.g., target container, maps)
        Case Else
            Err.Raise vbObjectError + 7000, "traverseProduct", "Unsupported traversal mode."
    End Select
            If outRefs Is Nothing Then Set outRefs = New Collection
            Set dictSeen = CreateObject("Scripting.Dictionary")
        Case tmGetInstances
            If outRefs Is Nothing Then Set outRefs = New Collection
        Case tmDeepCopyStructure
            ' Initialize any copy state here if needed (e.g., target container, maps)
        Case Else
            Err.Raise vbObjectError + 7000, "traverseProduct", "Unsupported traversal mode."
    End Select

    '--- Start BFS
    q.Add root

    Do While q.Count > 0
        Set current = q(1): q.Remove 1
        If Not current Is Nothing Then

            ' Decide enqueue policy once per node
            enqueueEarly = (mode = tmDeepCopyStructure)
            enqueueLate  = (Not enqueueEarly)

            ' Enqueue children early ONLY for DeepCopy (it often needs instance-first construction order)
            If enqueueEarly Then
                On Error Resume Next
                Set kids = current.Products
                If Err.Number = 0 And Not kids Is Nothing Then
                    For i = 1 To kids.Count
                        q.Add kids.Item(i)
                    Next
                End If
                Err.Clear
                On Error GoTo 0
            End If

            ' Resolve reference safely (some instance nodes may not resolve)
            Set ref = Nothing
            On Error Resume Next
            Set ref = current.ReferenceProduct
            On Error GoTo 0

            If Not ref Is Nothing Then
                '--- Dispatch per traversal mode to isolated handlers
                Select Case mode
                    Case tmGetUniques
                        ' collect unique refs honoring outKind
                        Call tmCollectUniques(ref, outRefs, dictSeen, outKind)

                    Case tmGetProducts
                        Call tmCollectProducts(ref, outRefs, dictSeen)

                    Case tmGetParts
                        Call tmCollectParts(ref, outRefs, dictSeen)

                    Case tmGetInstances
                        ' collect instance Products (filter by ref type via outKind)
                        Call tmCollectInstances(current, ref, outRefs, outKind)

                    Case tmDeepCopyStructure
                        ' perform one node's copy step (mapping & creation handled inside)
                        Call tmDeepCopyNode(current, ref)

                End Select
            End If

            ' Enqueue children late for all modes EXCEPT DeepCopy (avoids double-enqueue)
            If enqueueLate Then
                On Error Resume Next
                Set kids = current.Products
                If Err.Number = 0 And Not kids Is Nothing Then
                    For i = 1 To kids.Count
                        q.Add kids.Item(i)
                    Next
                End If
                Err.Clear
                On Error GoTo 0
            End If

        End If
    Loop
End Sub

'==========================================================
' tm-handlers (one-small-purpose routines, re-usable)
'==========================================================

' Unique references across the structure (Products and/or Parts per outKind)
Private Sub tmCollectUniques(ByVal ref As Product, ByRef outRefs As Collection, ByVal dictSeen As Object, ByVal outKind As Integer)
    Dim docType As String: docType = getRefDocType(ref)
    If docType = "" Then Exit Sub

    If (outKind = uoProductsOnly And docType <> "ProductDocument") Then Exit Sub
    If (outKind = uoPartsOnly    And docType <> "PartDocument") Then Exit Sub

    Dim key As String: key = buildRefKey(ref, docType)
    If Not dictSeen.Exists(key) Then
        dictSeen.Add key, True
        outRefs.Add ref
    End If
End Sub

' Only Product references (unique by ref)
Private Sub tmCollectProducts(ByVal ref As Product, ByRef outRefs As Collection, ByVal dictSeen As Object)
    If getRefDocType(ref) <> "ProductDocument" Then Exit Sub
    Dim key As String: key = buildRefKey(ref, "ProductDocument")
    If Not dictSeen.Exists(key) Then
        dictSeen.Add key, True
        outRefs.Add ref
    End If
End Sub

' Only Part references (unique by ref)
Private Sub tmCollectParts(ByVal ref As Product, ByRef outRefs As Collection, ByVal dictSeen As Object)
    If getRefDocType(ref) <> "PartDocument" Then Exit Sub
    Dim key As String: key = buildRefKey(ref, "PartDocument")
    If Not dictSeen.Exists(key) Then
        dictSeen.Add key, True
        outRefs.Add ref
    End If
End Sub

' Instance collection (current is the instance; filter by its reference type)
Private Sub tmCollectInstances(ByVal current As Product, ByVal ref As Product, ByRef outRefs As Collection, ByVal outKind As Integer)
    Dim docType As String: docType = getRefDocType(ref)
    If docType = "" Then Exit Sub

    If (outKind = uoProductsOnly And docType <> "ProductDocument") Then Exit Sub
    If (outKind = uoPartsOnly    And docType <> "PartDocument") Then Exit Sub

    outRefs.Add current  ' keep instances (e.g., for counts, placements, activation)
End Sub
' Deep-copy one node: duplicates the product structure and parts, preserving hierarchy and positions.
' Uses a mapping dictionary to avoid duplicate references and maintain structure.
' Assumes a module-level variable "gCopyMap" (Scripting.Dictionary) and "gCopyRoot" (Product) are set up before traversal.
Private Sub tmDeepCopyNode(ByVal current As Product, ByVal ref As Product)
    Dim docType As String
    docType = getRefDocType(ref)
    If docType = "" Then Exit Sub

    Dim newRef As Product
    Dim key As String
    key = buildRefKey(ref, docType)

    ' Only create a new reference once per unique reference
    If Not gCopyMap.Exists(key) Then
        ' Duplicate the reference (open as new document)
        Dim newDoc As Object
        Set newDoc = CATIA.Documents.Open(ref.Parent.FullName)
        Set newRef = newDoc.Product
        gCopyMap.Add key, newRef
    Else
        Set newRef = gCopyMap(key)
    End If

    ' Add new instance to the parent in the new structure
    Dim parentCopy As Product
    If current.Parent Is Nothing Or current Is gCopyRoot Then
        ' This is the root node; set as gCopyRoot
        Set gCopyRoot = newRef
    Else
        ' Find parent in the copy structure
        Dim parentKey As String
        parentKey = buildRefKey(current.Parent.ReferenceProduct, getRefDocType(current.Parent.ReferenceProduct))
        Set parentCopy = gCopyMap(parentKey)
        ' Add new instance to parent
        Dim newInst As Product
        Set newInst = parentCopy.Products.AddComponent(newRef)
        ' Copy position/matrix if needed
        On Error Resume Next
        newInst.ApplyTransformation current.GetTechnologicalObject("ProductPosition").GetMatrix
        On Error GoTo 0
    End If

    ' Optionally copy properties, parameters, etc. (extend as needed)
    ' Example: Copy part number, name, etc.
    On Error Resume Next
    newRef.PartNumber = ref.PartNumber
    newRef.Name = ref.Name
    On Error GoTo 0
End Sub

'==========================================================
' Helpers (small, safe, and fast)
'==========================================================

' Determine reference document type from its owning Document name
Private Function getRefDocType(ByVal ref As Product) As String
    On Error Resume Next
    Dim nm As String: nm = ref.Parent.Name  ' Document.Name (e.g., *.CATPart / *.CATProduct)
    On Error GoTo 0
    If InStr(1, nm, ".CATPart", vbTextCompare) > 0 Then
        getRefDocType = "PartDocument"
    ElseIf InStr(1, nm, ".CATProduct", vbTextCompare) > 0 Then
        getRefDocType = "ProductDocument"
    Else
        getRefDocType = ""
    End If
End Function

' Build a stable uniqueness key for a reference
Private Function buildRefKey(ByVal ref As Product, ByVal docType As String) As String
    ' Prefer PartNumber when available; Name can vary with instances.
    Dim pn As String, nm As String
    On Error Resume Next
    pn = ref.PartNumber
    nm = ref.Name
    On Error GoTo 0
    If pn = "" Then pn = nm
    buildRefKey = docType & "|" & UCase$(Trim$(pn))
End Function


'===============================================================
' Wrapper functions for traversal (formerly wrappers.bas)
' These provide user-friendly APIs for common traversal tasks.
'===============================================================

' getProducts – returns reference Products (Products only)
' Parameters:
'   root   [Product]   - Root product to traverse.
'   unique [Boolean]   - If True, returns unique references; if False, returns all (duplicates allowed).
' Returns:
'   [Collection] of reference Products (Products only).
Public Function getProducts(ByVal root As Product, Optional ByVal unique As Boolean = False) As Collection
    Dim outRefs As Collection
    Set outRefs = New Collection
    If unique Then
        traverseProduct tmGetUniques, root, outRefs, uoProductsOnly
    Else
        traverseProduct tmGetProducts, root, outRefs, uoProductsOnly
    End If
    Set getProducts = outRefs
End Function



' getParts – returns reference Products (Parts only)
' Parameters:
'   root   [Product]   - Root product to traverse.
'   unique [Boolean]   - If True, returns unique references; if False, returns all (duplicates allowed).
' Returns:
'   [Collection] of reference Products (Parts only).
Public Function getParts(ByVal root As Product, Optional ByVal unique As Boolean = False) As Collection
    Dim outRefs As Collection
    Set outRefs = New Collection
    If unique Then
        traverseProduct tmGetUniques, root, outRefs, uoPartsOnly
    Else
        traverseProduct tmGetParts, root, outRefs, uoPartsOnly
    End If
    Set getParts = outRefs
End Function

' getUniques – returns unique reference Products (ordered)
' Parameters:
'   root [Product]           - Root product to traverse.
'   kind [uniqueOutKind]     - Filter: uoAll, uoProductsOnly, or uoPartsOnly.
' Returns:
'   [Collection] of unique references (Products and/or Parts).
Public Function getUniques(ByVal root As Product, Optional ByVal kind As Integer = uoAll) As Collection
    Dim outRefs As Collection
    Set outRefs = New Collection
    traverseProduct tmGetUniques, root, outRefs, kind
    Set getUniques = outRefs
End Function

' getInstances – returns instance Products (not references)
' Parameters:
'   root [Product]           - Root product to traverse.
'   kind [uniqueOutKind]     - Filter: uoAll, uoProductsOnly, or uoPartsOnly.
' Returns:
'   [Collection] of instance Products.
Public Function getInstances(ByVal root As Product, Optional ByVal kind As Integer = uoAll) As Collection
    Dim outInst As Collection
    Set outInst = New Collection
    traverseProduct tmGetInstances, root, outInst, kind
    Set getInstances = outInst
End Function

' deepCopyStructure – duplicates the product structure and parts, preserving hierarchy and positions.
' Parameters:
'   root [Product] - Root product to deep copy.
' Returns:
'   [Product] - The root of the new copied structure.
Public Function deepCopyStructure(ByVal root As Product) As Product
    ' Requires module-level variables:
    '   gCopyMap  As Object (Scripting.Dictionary)
    '   gCopyRoot As Product
    Static gCopyMap As Object
    Static gCopyRoot As Product

    Set gCopyMap = CreateObject("Scripting.Dictionary")
    Set gCopyRoot = Nothing

    traverseProduct tmDeepCopyStructure, root

    Set deepCopyStructure = gCopyRoot
End Function