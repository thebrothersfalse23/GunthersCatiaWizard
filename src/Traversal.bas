'===============================================================
' Module: Traversal.bas
' Purpose: Implements the core iterative queue-based traversal logic
'          for CATIA Product structures. Handles unique/reference
'          collection, instance collection, and optional write API.
'          Used by all wrapper functions for assembly/product traversal.
'===============================================================

' TraverseProduct – Iterative queue-based traversal
'---------------------------------------------------------------
' Private Sub TraverseProduct(mode, root, [outRefs], [outKind])
'
' Parameters:
'   mode    - TraversalMode enum, determines traversal behavior
'   root    - Starting Product object
'   outRefs - [Optional] Collection to receive output (refs or instances)
'   outKind - [Optional] UniqueOutKind enum, controls output filtering
'
' Behavior:
'   - Breadth-first traversal of the product structure
'   - Handles deduplication, instance/reference separation, and
'     optional property assignment (write API)
'   - Used internally by all public wrapper functions
'---------------------------------------------------------------

Public Sub TraverseProduct(ByVal mode As TraversalMode, _
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
                        If Len(GetPropStr(ref, "Nomenclature")) = 0 Then SafeSet ref, "Nomenclature", current.Name
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





