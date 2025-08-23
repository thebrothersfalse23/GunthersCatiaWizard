===============================================================
MODULE: Docs.bas
PURPOSE: Documentation and usage examples for Gunther's Catia Wizard.
         Keeps CATMain clean and provides reference for API usage.
===============================================================
Public Sub GunthersCatiaWizard_Docs()
    ' -------------------------------------------------------------------------
    '' WRAPPER FUNCTION USAGE (all return Collection unless noted)
    ''
    '' GetProducts(rootProd As Product, [unique As Boolean = False])  → reference Products only
    ''   Dim prodsAll As Collection
    ''   Set prodsAll = GetProducts(rootProd, False)  ' all refs (not deduped)
    ''   Dim prodsUniq As Collection
    ''   Set prodsUniq = GetProducts(rootProd, True)  ' uniques only
    ''
    '' GetParts(rootProd As Product, [unique As Boolean = False])      → reference Parts only
    ''   Dim partsAll As Collection
    ''   Set partsAll = GetParts(rootProd, False)     ' all refs (not deduped)
    ''   Dim partsUniq As Collection
    ''   Set partsUniq = GetParts(rootProd, True)     ' uniques only
    ''
    '' GetUniques(rootProd As Product, [kind As UniqueOutKind = uoAll]) → unique refs (ordered)
    ''   Dim uniqAll As Collection
    ''   Set uniqAll = GetUniques(rootProd, uoAll)            ' Products→Parts
    ''   Dim uniqProds As Collection
    ''   Set uniqProds = GetUniques(rootProd, uoProductsOnly) ' only Products
    ''   Dim uniqParts As Collection
    ''   Set uniqParts = GetUniques(rootProd, uoPartsOnly)    ' only Parts
    ''
    '' GetInstances(rootProd As Product, [kind As UniqueOutKind = uoAll]) → instance Products
    ''   Dim instAll As Collection
    ''   Set instAll = GetInstances(rootProd, uoAll)            ' Products→Parts
    ''   Dim instProds As Collection
    ''   Set instProds = GetInstances(rootProd, uoProductsOnly) ' only product instances
    ''   Dim instParts As Collection
    ''   Set instParts = GetInstances(rootProd, uoPartsOnly)    ' only part instances
    ''
    '' SafeSet(obj As Object, propName As String, value As String)
    ''   ' Safely sets a property if it exists on the object.
    ''
    '' GetPropStr(obj As Object, propName As String) As String
    ''   ' Safely gets a property value as string, or "" if not present.
    ''
    '' BuildRefKey(ref As Product, docType As String) As String
    ''   ' Returns a stable key for a reference Product (PartNumber|DocType|Definition?)
    ''
    '' EnsureActiveProductDocument() As Boolean
    ''   ' Ensures a ProductDocument is open and active, sets globals prodDoc/rootProd.
    ''
    '' EnsureDesignMode(root As Product)
    ''   ' Applies Design Mode to a Product for consistent traversal.
    ''
    '' TraverseProduct(mode As TraversalMode, root As Product, [ByRef outRefs As Collection], [outKind As UniqueOutKind = uoAll])
    ''   ' Core traversal logic, used by all wrappers.
    ''
    '' Enumerations:
    ''   TraversalMode: tmGetUniques, tmGetParts, tmAssignInstanceData, tmCollectRefsAll, tmGetInstances
    ''   UniqueOutKind: uoAll, uoProductsOnly, uoPartsOnly
    ''
    '' -------------------------------------------------------------------------
    '' EXAMPLE USE-CASE (pattern)
    ''
    '' Public Sub ExampleOfUseCase(inputProd As Product)
    ''     Dim prodsToRename As Collection
    ''     Set prodsToRename = GetInstances(inputProd, uoProductsOnly)
    ''
    ''     Dim i As Long, current As Product
    ''     For i = 1 To prodsToRename.Count
    ''         Set current = prodsToRename.Item(i)
    ''         ' If your CATIA build exposes DescriptionInst, set it here.
    ''         ' Some environments use Description or other instance fields.
    ''         ' Adjust the property name to your deployment if needed.
    ''         SafeSet current, "Description", "MADE BY AMCO"
    ''     Next i
    '' End Sub
    ''
    '' -------------------------------------------------------------------------
    '' UNDER THE HOOD
    ''   - tmGetUniques      → BuildRefKey() + Dictionary to return unique refs,
    ''                         ordered Products then Parts (or filtered by kind).
    ''   - tmCollectRefsAll  → returns all refs (no dedupe), ordered Products then Parts.
    ''   - tmGetInstances    → returns instance Products, bucketed by ref doc type.
    ''   - SafeSet           → safely sets a property if it exists on the object.
    ''   - GetPropStr        → safely gets a property value as string.
    ''   - BuildRefKey       → builds a stable key for deduplication.
    ''   - EnsureDesignMode  → applies Design Mode for traversal.
    ''   - EnsureActiveProductDocument → ensures a ProductDocument is open and sets globals.
    ''   - TraverseProduct   → core traversal, not called directly by users.
    ''
    '' NOTES
    ''   - All wrappers pass ByRef collections through TraverseProduct and return them.
    ''   - No UI side effects here; add exporters or viewers only when you ask.
    ''   - BuildRefKey default is "PartNumber|DocType" and appends "|Definition"
    ''     if available to reduce collisions across libraries.
    ''   - Design Mode is applied up-front for consistent traversal depth.
    ''   - UniqueOutKind: uoAll, uoProductsOnly, uoPartsOnly (enum for filtering).
    ''   - All collections are 1-based (VBA default).
    '' -------------------------------------------------------------------------
End Sub
