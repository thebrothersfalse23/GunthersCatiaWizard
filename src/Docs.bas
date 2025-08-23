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
