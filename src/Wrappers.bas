'===============================================================
' MODULE: Wrappers.bas
' PURPOSE: Provides wrapper functions for traversing CATIA Product structures.
'          These functions return collections of reference or instance Products/Parts,
'          with options for deduplication and filtering.
'
' FUNCTIONS:
'   - GetProducts:   Returns reference Products (Products only), deduped or all.
'   - GetParts:      Returns reference Parts (Parts only), deduped or all.
'   - GetUniques:    Returns unique references (Products and/or Parts).
'   - GetInstances:  Returns instance Products (not references), filtered by kind.
'
' All functions accept a root Product and optional parameters for deduplication or filtering.
'===============================================================

' GetProducts – returns reference Products (Products only)
' Parameters:
'   root   [Product]   - Root product to traverse.
'   unique [Boolean]   - If True, returns unique references; if False, returns all.
' Returns:
'   [Collection] of reference Products (Products only).
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
' Parameters:
'   root   [Product]   - Root product to traverse.
'   unique [Boolean]   - If True, returns unique references; if False, returns all.
' Returns:
'   [Collection] of reference Products (Parts only).
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
' Parameters:
'   root [Product]           - Root product to traverse.
'   kind [UniqueOutKind]     - Filter: uoAll, uoProductsOnly, or uoPartsOnly.
' Returns:
'   [Collection] of unique references (Products and/or Parts).
Public Function GetUniques(ByVal root As Product, Optional ByVal kind As UniqueOutKind = uoAll) As Collection
    Dim outRefs As Collection
    TraverseProduct tmGetUniques, root, outRefs, kind
    Set GetUniques = outRefs
End Function

' GetInstances – returns instance Products (not references)
' Parameters:
'   root [Product]           - Root product to traverse.
'   kind [UniqueOutKind]     - Filter: uoAll, uoProductsOnly, or uoPartsOnly.
' Returns:
'   [Collection] of instance Products.
Public Function GetInstances(ByVal root As Product, Optional ByVal kind As UniqueOutKind = uoAll) As Collection
    Dim outInst As Collection
    TraverseProduct tmGetInstances, root, outInst, kind
    Set GetInstances = outInst
End Function


