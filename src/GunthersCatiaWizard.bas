'===============================================================
' Macro:        Gunther's Catia Wizard
' Author:       [Your Name]
' Version:      v1.0 – 2025-08-22
' CATIA:        V5 (tested on V5-6R2020+; expected OK from V5R2016)
'
' DOCS INDEX
'   1) Quick Start ........................... See: GunthersCatiaWizard_Docs()
'   2) API surface (wrappers) ................ GetProducts/GetParts/GetUniques/GetInstances
'   3) Modes & contracts ..................... TraversalMode, UniqueOutKind, BuildRefKey()
'   4) Behavior guarantees ................... Ordering, uniqueness, instance/refs split
'   5) Optional write API .................... AssignInstanceData()
'   6) Full narrative & examples ............. Bottom section: "Documentation / Usage Examples"
'
' NOTE: This file is the source of truth for user docs. Do not strip comments.
'===============================================================

Option Explicit

'===============================================================
' Global Variables (kept minimal)
'===============================================================
Public prodDoc As ProductDocument     ' Active ProductDocument
Public rootProd As Product            ' Root Product of the assembly

'===============================================================
' Enumerations
'===============================================================
Public Enum TraversalMode
    tmGetUniques = 1            ' unique reference Products (deduped)
    tmGetParts = 2              ' reserved placeholder (not used by wrappers)
    tmAssignInstanceData = 3    ' explicit write API (separate from read traversals)
    tmCollectRefsAll = 4        ' all reference Products (not deduped)
    tmGetInstances = 5          ' instance Products by kind
End Enum

' Kind selector for refs/instances
Public Enum UniqueOutKind
    uoAll = 0            ' Products first, then Parts
    uoProductsOnly = 1   ' Products only
    uoPartsOnly = 2      ' Parts only
End Enum

'===============================================================
' Entry Point (guards → init → sample call)
'===============================================================
Sub CATMain()

    If Not EnsureActiveProductDocument() Then Exit Sub

    ' Example: count unique reference Products+Parts
    Dim uniqAll As Collection
    Set uniqAll = GetUniques(rootProd, uoAll)

    MsgBox "Unique references found: " & CStr(uniqAll.Count), vbInformation, "Gunther's CATIA Wizard"

    ' Keep Main clean. See GunthersCatiaWizard_Docs for full examples & usage.

End Sub

'===============================================================
' WRAPPER FUNCTIONS (clean call surface for Main/other code)
'===============================================================

' GetProducts – returns reference Products (Products only)
' unique:=True  → deduped via tmGetUniques
' unique:=False → all refs via tmCollectRefsAll
Public Function GetProducts(ByVal root As Product, Optional ByVal unique As Boolean = False) As Collection
    Dim outRefs As Collection
    If unique Then
        TraverseProduct tmGetUniques, root, outRefs, uoProductsOnly
    Else
        TraverseProduct tmCollectRefsAll, root, outRefs, uoProductsOnly
    End If
    Set GetProducts = outRefs
End Function

' GetParts – returns reference Products (Parts only)
' unique:=True  → deduped via tmGetUniques
' unique:=False → all refs via tmCollectRefsAll
Public Function GetParts(ByVal root As Product, Optional ByVal unique As Boolean = False) As Collection
    Dim outRefs As Collection
    If unique Then
        TraverseProduct tmGetUniques, root, outRefs, uoPartsOnly
    Else
        TraverseProduct tmCollectRefsAll, root, outRefs, uoPartsOnly
    End If
    Set GetParts = outRefs
End Function

' GetUniques – returns unique reference Products (ordered)
' kind: uoAll (Products→Parts), uoProductsOnly, or uoPartsOnly
Public Function GetUniques(ByVal root As Product, Optional ByVal kind As UniqueOutKind = uoAll) As Collection
    Dim outRefs As Collection
    TraverseProduct tmGetUniques, root, outRefs, kind
    Set GetUniques = outRefs
End Function

' GetInstances – returns instance Products (not references)
' kind: uoAll (Products→Parts), uoProductsOnly, or uoPartsOnly
Public Function GetInstances(ByVal root As Product, Optional ByVal kind As UniqueOutKind = uoAll) As Collection
    Dim outInst As Collection
    TraverseProduct tmGetInstances, root, outInst, kind
    Set GetInstances = outInst
End Function

'===============================================================
' Optional write API (kept separate from read traversals)
' AssignInstanceData – aligns instance/reference text fields:
'   1) Instance Description = current Name
'   2) If reference Nomenclature is empty, copy instance Name
'   3) Instance Name = reference PartNumber
'===============================================================
Public Sub AssignInstanceData(ByVal root As Product)
    Dim unused As Collection
    TraverseProduct tmAssignInstanceData, root, unused, uoAll
End Sub

'===============================================================
' TraverseProduct – Iterative queue-based traversal
'   mode:      traversal behavior
'   root:      starting Product
'   outRefs:   [Optional] Collection receiver (refs or instances depending on mode)
'   outKind:   [Optional] controls which bucket(s) to return (default uoAll)
'===============================================================
Private Sub TraverseProduct(ByVal mode As TraversalMode, _
                            ByVal root As Product, _
                            Optional ByRef outRefs As Collection, _
                            Optional ByVal outKind As UniqueOutKind = uoAll)

    If root Is Nothing Then Exit Sub

    ' Ensure Design Mode for consistent child access (no prompts; fast path)
    EnsureDesignMode root

    ' Breadth-first traversal using a simple Collection as a queue.
    Dim q As Collection: Set q = New Collection
    q.Add root

    Dim current As Product, kids As Products
    Dim ref As Product
    Dim i As Long

    ' --- Accumulators for REFERENCE Products ---
    Dim seen As Object              ' Scripting.Dictionary for uniques
    Dim prodRefs As Collection      ' refs whose parent doc is ProductDocument
    Dim partRefs As Collection      ' refs whose parent doc is PartDocument

    ' --- Accumulators for INSTANCE Products ---
    Dim instProd As Collection      ' instance products (ref parent ProductDocument)
    Dim instPart As Collection      ' instance products (ref parent PartDocument)

    If mode = tmGetUniques Then
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = 1 ' vbTextCompare
        Set prodRefs = New Collection
        Set partRefs = New Collection
    ElseIf mode = tmCollectRefsAll Then
        Set prodRefs = New Collection
        Set partRefs = New Collection
    ElseIf mode = tmGetInstances Then
        Set instProd = New Collection
        Set instPart = New Collection
    End If

    Do While q.Count > 0
        Set current = q(1): q.Remove 1
        If Not current Is Nothing Then

            ' --- ReferenceProduct read (scoped error handling) ---
            On Error Resume Next
            Set ref = current.ReferenceProduct
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo AfterModeWork
            End If
            On Error GoTo 0

            If Not ref Is Nothing Then
                Dim dt As String
                dt = TypeName(ref.Parent) ' "ProductDocument" or "PartDocument"

                Select Case mode

                    Case tmGetUniques
                        ' Build dedupe key for the reference product
                        Dim key As String
                        key = BuildRefKey(ref, dt) ' e.g., "PartNumber|DocType|Definition?"
                        If Len(key) > 0 Then
                            If Not seen.Exists(key) Then
                                seen.Add key, True
                                If dt = "ProductDocument" Then
                                    prodRefs.Add ref
                                ElseIf dt = "PartDocument" Then
                                    partRefs.Add ref
                                End If
                            End If
                        End If

                    Case tmCollectRefsAll
                        ' Collect every reference (no dedupe)
                        If dt = "ProductDocument" Then
                            prodRefs.Add ref
                        ElseIf dt = "PartDocument" Then
                            partRefs.Add ref
                        End If

                    Case tmAssignInstanceData
                        ' Be explicit: instance-side and reference-side edits
                        SafeSet current, "Description", current.Name
                        If Len(GetStringSafe(ref, "Nomenclature")) = 0 Then SafeSet ref, "Nomenclature", current.Name
                        SafeSet current, "Name", ref.PartNumber   ' rename instance to ref PartNumber

                    Case tmGetInstances
                        ' Bucket the INSTANCE by its reference document type
                        If dt = "ProductDocument" Then
                            instProd.Add current
                        ElseIf dt = "PartDocument" Then
                            instPart.Add current
                        End If

                    Case tmGetParts
                        ' Placeholder (not used by wrappers)
                        Debug.Print "Visit: "; current.PartNumber

                End Select
            End If

AfterModeWork:
            ' Enqueue children (if any). Leaf parts return Products.Count = 0.
            Set kids = current.Products
            If Not kids Is Nothing Then
                For i = 1 To kids.Count
                    q.Add kids.Item(i)
                Next i
            End If
        End If
    Loop

    ' Assemble outputs based on mode and outKind
    Dim result As Collection

    Select Case mode

        Case tmGetUniques, tmCollectRefsAll
            Set result = New Collection
            Select Case outKind
                Case uoProductsOnly
                    For i = 1 To prodRefs.Count: result.Add prodRefs.Item(i): Next i
                Case uoPartsOnly
                    For i = 1 To partRefs.Count: result.Add partRefs.Item(i): Next i
                Case Else ' uoAll: Products first, then Parts
                    For i = 1 To prodRefs.Count: result.Add prodRefs.Item(i): Next i
                    For i = 1 To partRefs.Count: result.Add partRefs.Item(i): Next i
            End Select

        Case tmGetInstances
            Set result = New Collection
            Select Case outKind
                Case uoProductsOnly
                    For i = 1 To instProd.Count: result.Add instProd.Item(i): Next i
                Case uoPartsOnly
                    For i = 1 To instPart.Count: result.Add instPart.Item(i): Next i
                Case Else ' uoAll: Products first, then Parts
                    For i = 1 To instProd.Count: result.Add instProd.Item(i): Next i
                    For i = 1 To instPart.Count: result.Add instPart.Item(i): Next i
            End Select

        Case Else
            ' Modes with no output contract: ignore
    End Select

    If Not result Is Nothing Then
        Set outRefs = result
    End If
End Sub

'===============================================================
' BuildRefKey – Builds a stable, human-readable key for a reference
' Default: "PartNumber|DocType" ; if Definition exists → "PartNumber|DocType|Definition"
'===============================================================
Private Function BuildRefKey(ByVal ref As Product, ByVal docType As String) As String
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

'===============================================================
' Safe property helpers (late-bound without CallByName)
'===============================================================
Private Sub SafeSet(ByVal obj As Object, ByVal propName As String, ByVal value As String)
    On Error Resume Next
    Select Case propName
        Case "Nomenclature": obj.Nomenclature = value
        Case "Name":         obj.Name = value
        Case "Description":  obj.Description = value
    End Select
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GetStringSafe(ByVal obj As Object, ByVal propName As String) As String
    On Error Resume Next
    Select Case propName
        Case "Nomenclature": GetStringSafe = obj.Nomenclature
        Case "Name":         GetStringSafe = obj.Name
        Case "Description":  GetStringSafe = obj.Description
        Case Else:           GetStringSafe = ""
    End Select
    If Err.Number <> 0 Then GetStringSafe = ""
    Err.Clear
    On Error GoTo 0
End Function

'===============================================================
' Work-mode & document guards
'===============================================================
Private Sub EnsureDesignMode(ByVal root As Product)
    On Error Resume Next
    root.ApplyWorkMode DESIGN_MODE
    Err.Clear
    On Error GoTo 0
End Sub

Private Function EnsureActiveProductDocument() As Boolean
    EnsureActiveProductDocument = False

    If CATIA.Documents.Count = 0 Then
        MsgBox "No document open.", vbExclamation, "Gunther's CatIA Wizard"
        Exit Function
    End If

    If TypeName(CATIA.ActiveDocument) <> "ProductDocument" Then
        MsgBox "Active document is not a CATProduct.", vbExclamation, "Gunther's CatIA Wizard"
        Exit Function
    End If

    Set prodDoc = CATIA.ActiveDocument
    Set rootProd = prodDoc.Product
    EnsureActiveProductDocument = True
End Function

'===============================================================
' Documentation / Usage Examples (keeps CATMain clean)
'===============================================================
Sub GunthersCatiaWizard_Docs()
    ' -------------------------------------------------------------------------
    ' WRAPPER FUNCTION USAGE (all return Collection)
    '
    ' GetProducts(rootProd [, unique As Boolean])  → reference Products only
    '   Dim prodsAll As Collection
    '   Set prodsAll = GetProducts(rootProd, False)  ' all refs (not deduped)
    '   Dim prodsUniq As Collection
    '   Set prodsUniq = GetProducts(rootProd, True)  ' uniques only
    '
    ' GetParts(rootProd [, unique As Boolean])      → reference Parts only
    '   Dim partsAll As Collection
    '   Set partsAll = GetParts(rootProd, False)     ' all refs (not deduped)
    '   Dim partsUniq As Collection
    '   Set partsUniq = GetParts(rootProd, True)     ' uniques only
    '
    ' GetUniques(rootProd [, kind As UniqueOutKind]) → unique refs (ordered)
    '   Dim uniqAll As Collection
    '   Set uniqAll = GetUniques(rootProd, uoAll)            ' Products→Parts
    '   Dim uniqProds As Collection
    '   Set uniqProds = GetUniques(rootProd, uoProductsOnly) ' only Products
    '   Dim uniqParts As Collection
    '   Set uniqParts = GetUniques(rootProd, uoPartsOnly)    ' only Parts
    '
    ' GetInstances(rootProd [, kind As UniqueOutKind]) → instance Products
    '   Dim instAll As Collection
    '   Set instAll = GetInstances(rootProd, uoAll)            ' Products→Parts
    '   Dim instProds As Collection
    '   Set instProds = GetInstances(rootProd, uoProductsOnly) ' only product instances
    '   Dim instParts As Collection
    '   Set instParts = GetInstances(rootProd, uoPartsOnly)    ' only part instances
    '
    ' -------------------------------------------------------------------------
    ' EXAMPLE USE-CASE (pattern)
    '
    ' Sub ExampleOfUseCase(inputProd As Product)
    '     Dim prodsToRename As Collection
    '     Set prodsToRename = GetInstances(inputProd, uoProductsOnly)
    '
    '     Dim i As Long, current As Product
    '     For i = 1 To prodsToRename.Count
    '         Set current = prodsToRename.Item(i)
    '         ' If your CATIA build exposes DescriptionInst, set it here.
    '         ' Some environments use Description or other instance fields.
    '         ' Adjust the property name to your deployment if needed.
    '         SafeSet current, "Description", "MADE BY AMCO"
    '     Next i
    ' End Sub
    '
    ' -------------------------------------------------------------------------
    ' UNDER THE HOOD
    '   - tmGetUniques      → BuildRefKey() + Dictionary to return unique refs,
    '                         ordered Products then Parts (or filtered by kind).
    '   - tmCollectRefsAll  → returns all refs (no dedupe), ordered Products then Parts.
    '   - tmGetInstances    → returns instance Products, bucketed by ref doc type.
    '   - AssignInstanceData → side-effect routine (opt-in): sets
    '       • instance.Description = instance.Name
    '       • if ref.Nomenclature="" then ref.Nomenclature = instance.Name
    '       • instance.Name = ref.PartNumber
    '
    ' NOTES
    '   - All wrappers pass ByRef collections through TraverseProduct and return them.
    '   - No UI side effects here; add exporters or viewers only when you ask.
    '   - BuildRefKey default is "PartNumber|DocType" and appends "|Definition"
    '     if available to reduce collisions across libraries.
    '   - Design Mode is applied up-front for consistent traversal depth.
    ' -------------------------------------------------------------------------
End Sub
