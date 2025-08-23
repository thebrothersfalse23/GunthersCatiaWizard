'===============================================================
' MODULE: traversal.bas
' PURPOSE: Implements the core iterative queue-based traversal logic
'          for CATIA Product structures. Handles unique/reference
'          collection, instance collection, and optional write API.
'          Used by all wrapper functions for assembly/product traversal.
'===============================================================

' traverseProduct – Iterative queue-based traversal
'---------------------------------------------------------------
' Private Sub traverseProduct(mode, root, [outRefs], [outKind])
'
' Parameters:
'   mode    - traversalMode enum, determines traversal behavior
'   root    - Starting Product object
'   outRefs - [Optional] Collection to receive output (refs or instances)
'   outKind - [Optional] uniqueOutKind enum, controls output filtering
'
' Behavior:
'   - Breadth-first traversal of the product structure
'   - Handles deduplication, instance/reference separation, and
'     optional property assignment (write API)
'   - Used internally by all public wrapper functions
'---------------------------------------------------------------

' Helper to copy one collection into another
Private Sub copyColInto(ByRef dst As Collection, ByVal src As Collection)
    If src Is Nothing Then Exit Sub
    Dim k As Long
    For k = 1 To src.Count
        dst.Add src.Item(k)
    Next
End Sub

Public Sub traverseProduct(ByVal mode As traversalMode, _
                          ByVal root As Product, _
                          Optional ByRef outRefs As Collection, _
                          Optional ByVal outKind As uniqueOutKind = uoAll)

    If root Is Nothing Then Exit Sub

    ' Ensure Design Mode for consistent child access (no prompts; fast path)
    ensureDesignMode root

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

    ' Ensure all collections are initialized to avoid object errors
    If prodRefs Is Nothing Then Set prodRefs = New Collection
    If partRefs Is Nothing Then Set partRefs = New Collection
    If instProd Is Nothing Then Set instProd = New Collection
    If instPart Is Nothing Then Set instPart = New Collection

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
                        key = buildRefKey(ref, dt) ' e.g., "PartNumber|DocType|Definition?"
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
                        safeSet current, "Description", current.Name
                        If Len(getPropStr(ref, "Nomenclature")) = 0 Then safeSet ref, "Nomenclature", current.Name
                        safeSet current, "Name", ref.PartNumber   ' rename instance to ref PartNumber

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
            ' Enqueue children (BFS)
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
    Loop

    ' Assemble outputs based on mode and outKind
    If mode = tmGetUniques Or mode = tmCollectRefsAll Then
        If outRefs Is Nothing Then Set outRefs = New Collection
        Select Case outKind
            Case uoProductsOnly
                copyColInto outRefs, prodRefs
            Case uoPartsOnly
                copyColInto outRefs, partRefs
            Case Else ' uoAll: Products first, then Parts
                copyColInto outRefs, prodRefs
                copyColInto outRefs, partRefs
        End Select
    ElseIf mode = tmGetInstances Then
        If outRefs Is Nothing Then Set outRefs = New Collection
        Select Case outKind
            Case uoProductsOnly
                copyColInto outRefs, instProd
            Case uoPartsOnly
                copyColInto outRefs, instPart
            Case Else
                copyColInto outRefs, instProd
                copyColInto outRefs, instPart
        End Select
    End If
End Sub





